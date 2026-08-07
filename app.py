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
    """Extrai o texto de um docx cru (transcrição ou questões)."""
    data = await file.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        xml = z.read('word/document.xml').decode('utf-8')
    except Exception as e:
        raise HTTPException(400, f"docx inválido: {e}")
    paras = re.findall(r'<w:p(?: [^>]*)?/>|<w:p(?: [^>]*)?>.*?</w:p>', xml, re.S)
    lines = []
    for p in paras:
        t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))
        lines.append(t)
    text = '\n'.join(lines)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()
    return {"text": text, "chars": len(text)}


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
