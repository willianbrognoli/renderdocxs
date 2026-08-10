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

    # --- RLM/matemática: equações do editor do Word (OMML) e símbolos ---
    def _mt(s):
        return ''.join(re.findall(r'<m:t[^>]*>([^<]*)</m:t>', s))

    def _omml_linear(bloco):
        """Lineariza uma equação OMML: frações viram (a)/(b), potências a^b,
        índices a_b, raízes √(x); o resto vira a sequência dos textos."""
        s = bloco
        for _ in range(12):
            antes = s
            s = re.sub(r'<m:f(?: [^>]*)?>.*?<m:num(?: [^>]*)?>(.*?)</m:num>\s*'
                       r'<m:den(?: [^>]*)?>(.*?)</m:den>\s*</m:f>',
                       lambda m: '<m:t>(%s)/(%s)</m:t>' % (_mt(m.group(1)), _mt(m.group(2))),
                       s, flags=re.S)
            s = re.sub(r'<m:sSup(?: [^>]*)?>.*?<m:e(?: [^>]*)?>(.*?)</m:e>\s*'
                       r'<m:sup(?: [^>]*)?>(.*?)</m:sup>\s*</m:sSup>',
                       lambda m: '<m:t>%s^%s</m:t>' % (_mt(m.group(1)), _mt(m.group(2))),
                       s, flags=re.S)
            s = re.sub(r'<m:sSub(?: [^>]*)?>.*?<m:e(?: [^>]*)?>(.*?)</m:e>\s*'
                       r'<m:sub(?: [^>]*)?>(.*?)</m:sub>\s*</m:sSub>',
                       lambda m: '<m:t>%s_%s</m:t>' % (_mt(m.group(1)), _mt(m.group(2))),
                       s, flags=re.S)
            s = re.sub(r'<m:rad(?: [^>]*)?>.*?<m:e(?: [^>]*)?>(.*?)</m:e>\s*</m:rad>',
                       lambda m: '<m:t>\u221a(%s)</m:t>' % _mt(m.group(1)),
                       s, flags=re.S)
            if s == antes:
                break
        return _mt(s)

    _SYMBOL_MAP = {
        0xD9: '\u2227', 0xDA: '\u2228', 0xD8: '\u00ac', 0xAE: '\u2192',
        0xAB: '\u2194', 0xDE: '\u21d2', 0xDB: '\u21d4', 0x22: '\u2200',
        0x24: '\u2203', 0xCE: '\u2208', 0xCF: '\u2209', 0xC6: '\u2205',
        0xC7: '\u2229', 0xC8: '\u222a', 0xCC: '\u2282', 0xCD: '\u2286',
        0xA3: '\u2264', 0xB3: '\u2265', 0xB9: '\u2260', 0xD6: '\u221a',
        0xBB: '\u2248', 0xA5: '\u221e', 0x40: '\u2245', 0x5E: '\u22a5',
        # equivalência, portanto, operadores e gregas frequentes em RLM
        0xBA: '\u2261', 0x5C: '\u2234', 0xB1: '\u00b1', 0xB4: '\u00d7',
        0xB8: '\u00f7', 0xC5: '\u2295', 0xC4: '\u2297', 0xB5: '\u221d',
        0xB6: '\u2202', 0xD1: '\u2207', 0xD5: '\u220f', 0xE5: '\u2211',
        0x61: '\u03b1', 0x62: '\u03b2', 0x67: '\u03b3', 0x64: '\u03b4',
        0x65: '\u03b5', 0x71: '\u03b8', 0x6c: '\u03bb', 0x6d: '\u03bc',
        0x70: '\u03c0', 0x73: '\u03c3', 0x66: '\u03c6', 0x77: '\u03c9',
        0x44: '\u0394', 0x53: '\u03a3', 0x57: '\u03a9', 0x46: '\u03a6',
    }

    def _pre(p):
        # equações OMML viram run de texto linearizado
        p = re.sub(r'<m:oMath(?: [^>]*)?>.*?</m:oMath>',
                   lambda m: ('<w:r><w:t xml:space="preserve"> %s </w:t></w:r>'
                              % _omml_linear(m.group(0))),
                   p, flags=re.S)
        # símbolos de fonte Symbol (w:sym) viram o caractere Unicode
        def _sym(m):
            try:
                code = int(m.group(1), 16) & 0xFF
            except ValueError:
                return ''
            ch = _SYMBOL_MAP.get(code)
            return '<w:t xml:space="preserve">%s</w:t>' % ch if ch else ''
        p = re.sub(r'<w:sym\b[^>]*w:char="(?:F0)?([0-9A-Fa-f]{2,4})"[^/>]*/>',
                   _sym, p)
        return p


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
        p = _pre(p)
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
