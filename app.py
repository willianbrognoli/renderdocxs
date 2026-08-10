# -*- coding: utf-8 -*-
"""Microserviço de renderização Águia - Carreiras Policiais.

Endpoints:
  GET  /health            -> status
  POST /extract           -> multipart docx -> {"text": "..."} (texto cru)
  POST /render            -> JSON estruturado -> docx diagramado (binário)

O JSON de /render segue o schema documentado no README.
"""
import io
import os
import re
import json
import tempfile
import zipfile
import subprocess

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel

import builder

app = FastAPI(title="Renderizador Águia", version="1.0")

FONTS_READY = False


def _ensure_fonts():
    global FONTS_READY
    if FONTS_READY:
        return
    try:
        builder.install_embedded_fonts()
        extra = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
        if os.path.isdir(extra):
            dest = os.path.expanduser('~/.fonts')
            os.makedirs(dest, exist_ok=True)
            for fn in os.listdir(extra):
                if fn.lower().endswith(('.ttf', '.otf')):
                    src = os.path.join(extra, fn)
                    dst = os.path.join(dest, fn)
                    if not os.path.exists(dst):
                        with open(src, 'rb') as a, open(dst, 'wb') as b:
                            b.write(a.read())
            subprocess.run(['fc-cache', '-f'], capture_output=True)
    except Exception:
        pass
    FONTS_READY = True


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    """Extrai texto e imagens de um docx cru, preservando as marcações do
    professor: runs sublinhados viram __texto__ e runs em vermelho viram
    %%texto%%. Imagens do corpo viram marcadores [IMAGEM n] no ponto exato
    do texto e são devolvidas em base64."""
    import base64 as b64
    data = await file.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        xml = z.read('word/document.xml').decode('utf-8')
    except Exception as e:
        raise HTTPException(400, f"docx inválido: {e}")
    try:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    except KeyError:
        rel_map = {}

    RUN_RE = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>', re.S)

    def _marcas(run):
        rpr = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
        rpr = rpr.group(0) if rpr else ''
        sub = re.search(r'<w:u w:val="(?!none)', rpr) is not None
        verm = False
        m = re.search(r'<w:color w:val="([0-9A-Fa-f]{6})"', rpr)
        if m:
            h = m.group(1)
            r_, g_, b_ = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            verm = r_ > 120 and g_ < 90 and b_ < 90
        return sub, verm

    paras = re.findall(r'<w:p(?: [^>]*)?/>|<w:p(?: [^>]*)?>.*?</w:p>', xml, re.S)
    lines, images, seen = [], [], {}
    for p in paras:
        partes = []
        for run in RUN_RE.findall(p):
            t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run))
            if not t:
                continue
            if t.strip():
                sub, verm = _marcas(run)
                if verm:
                    t = '%%' + t + '%%'
                elif sub:
                    t = '__' + t + '__'
            partes.append(t)
        t = ''.join(partes)
        # funde marcas de runs adjacentes: __a____b__ -> __ab__
        t = t.replace('%%%%', '').replace('____', '')
        for rid in re.findall(r'r:embed="([^"]+)"', p):
            target = rel_map.get(rid, '')
            if 'media/' not in target:
                continue
            if target in seen:
                t += ' [IMAGEM %d]' % seen[target]
                continue
            try:
                raw = z.read('word/' + target.lstrip('/'))
            except KeyError:
                continue
            if len(raw) < 3000:
                continue  # ícone/decoração
            n = len(images) + 1
            seen[target] = n
            ext = target.rsplit('.', 1)[-1].lower()
            mime = 'image/png' if ext == 'png' else 'image/jpeg' if ext in ('jpg', 'jpeg') else 'image/' + ext
            images.append({'n': n, 'mime': mime, 'base64': b64.b64encode(raw).decode('ascii')})
            t += ' [IMAGEM %d]' % n
        lines.append(t)
    text = '\n'.join(lines)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()
    return {"text": text, "chars": len(text), "images": images}


class RenderRequest(BaseModel):
    data: dict
    filename: str = "material.docx"
    paginate: bool = True


@app.post("/render")
def render(req: RenderRequest):
    _ensure_fonts()
    try:
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, re.sub(r'[^\w. \-]', '_', req.filename) or 'material.docx')
            builder.render(req.data, out, paginate=req.paginate)
            with open(out, 'rb') as fh:
                blob = fh.read()
    except KeyError as e:
        raise HTTPException(422, f"campo obrigatório ausente no JSON: {e}")
    except Exception as e:
        raise HTTPException(500, f"falha na renderização: {type(e).__name__}: {e}")
    headers = {
        'Content-Disposition': 'attachment; filename="%s"' % os.path.basename(
            re.sub(r'[^\w. \-]', '_', req.filename))
    }
    return Response(
        content=blob,
        media_type=('application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.document'),
        headers=headers)
