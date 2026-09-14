# -*- coding: utf-8 -*-
"""Microserviço de renderização Águia - Carreiras Policiais.

Endpoints:
  GET  /health            -> status
  POST /extract           -> multipart docx OU pptx -> {"text": "...", "images": [...]}
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


# ---------------------------------------------------------------------------
# Helpers compartilhados (docx e pptx): equações OMML e fonte Symbol
# ---------------------------------------------------------------------------
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


def _extract_docx(z):
    """Extrai texto e imagens de um docx cru, preservando as marcações do
    professor: runs sublinhados viram __texto__ e runs em vermelho viram
    %%texto%%. Imagens do corpo viram marcadores [IMAGEM n] no ponto exato
    do texto e são devolvidas em base64."""
    import base64 as b64
    try:
        xml = z.read('word/document.xml').decode('utf-8')
    except Exception as e:
        raise HTTPException(400, f"docx inválido: {e}")
    try:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    except KeyError:
        rel_map = {}

    RUN_RE = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>', re.S)

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
    return {"text": text, "chars": len(text), "images": images, "kind": "docx"}


# ---------------------------------------------------------------------------
# v7.3: /extract também lê PPTX (slides do professor)
# ---------------------------------------------------------------------------
_NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
}
_R_EMBED = '{%s}embed' % _NS['r']
_R_ID = '{%s}id' % _NS['r']
# placeholders que são cromo do slide, não conteúdo
_PH_IGNORAR = {'sldNum', 'ftr', 'dt'}
# uma imagem que se repete em muitos slides é logo/decoração
_REPETE_MIN_SLIDES = 3


def _pptx_rels(z, part):
    """Mapa rId -> alvo absoluto dentro do zip para a parte informada
    (ex.: 'ppt/slides/slide3.xml')."""
    pasta, nome = part.rsplit('/', 1)
    try:
        rels = z.read('%s/_rels/%s.rels' % (pasta, nome)).decode('utf-8')
    except KeyError:
        return {}
    out = {}
    for rid, alvo in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
        if alvo.startswith('/'):
            out[rid] = alvo.lstrip('/')
        else:
            out[rid] = os.path.normpath(os.path.join(pasta, alvo)).replace('\\', '/')
    return out


def _pptx_ordem_slides(z):
    """Lista de partes de slide na ordem de apresentação."""
    try:
        pres = z.read('ppt/presentation.xml').decode('utf-8')
    except KeyError:
        raise HTTPException(400, 'pptx inválido: sem ppt/presentation.xml')
    rels = _pptx_rels(z, 'ppt/presentation.xml')
    ids = re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', pres)
    slides = [rels[r] for r in ids if r in rels]
    if not slides:  # fallback: ordem numérica dos arquivos
        slides = sorted((n for n in z.namelist()
                         if re.match(r'ppt/slides/slide\d+\.xml$', n)),
                        key=lambda n: int(re.search(r'(\d+)\.xml$', n).group(1)))
    return slides


def _pptx_texto_run(r):
    """Texto de um <a:r>, com marcas __ (sublinhado) e %% (vermelho) e
    conversão de fonte Symbol para Unicode."""
    t = ''.join(r.itertext())
    if not t:
        return ''
    rpr = r.find('a:rPr', _NS)
    if rpr is None:
        return t
    sym = rpr.find('a:sym', _NS)
    if sym is not None and 'symbol' in (sym.get('typeface') or '').lower():
        t = ''.join(_SYMBOL_MAP.get(ord(c) & 0xFF, c) for c in t)
    if not t.strip():
        return t
    verm = False
    clr = rpr.find('a:solidFill/a:srgbClr', _NS)
    if clr is not None:
        h = (clr.get('val') or '').upper()
        if len(h) == 6:
            r_, g_, b_ = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            verm = r_ > 120 and g_ < 90 and b_ < 90
    if verm:
        return '%%' + t + '%%'
    if (rpr.get('u') or 'none') != 'none':
        return '__' + t + '__'
    return t


def _pptx_paragrafo(p):
    """Um <a:p> -> uma linha de texto (equações em [FÓRMULA] ...)."""
    partes = []
    for filho in p:
        tag = filho.tag.split('}')[-1]
        if tag == 'r':
            partes.append(_pptx_texto_run(filho))
        elif tag == 'fld':
            if (filho.get('type') or '') in ('slidenum', 'datetime', 'datetime1'):
                continue
            partes.append(''.join(filho.itertext()))
        elif tag == 'br':
            partes.append(' ')
        elif tag == 'm':  # a14:m -> equação OMML
            from lxml import etree
            for om in filho.iter('{%s}oMath' % _NS['m']):
                lin = _omml_linear(etree.tostring(om, encoding='unicode'))
                if lin.strip():
                    partes.append(' [FÓRMULA] %s ' % lin.strip())
    t = ''.join(partes).replace('%%%%', '').replace('____', '')
    t = re.sub(r'[ \t]+', ' ', t).strip()
    if not t:
        return ''
    ppr = p.find('a:pPr', _NS)
    lvl = int(ppr.get('lvl') or 0) if ppr is not None else 0
    return ('  ' * lvl) + t


def _pptx_txbody(tx):
    return [l for l in (_pptx_paragrafo(p) for p in tx.findall('a:p', _NS)) if l]


def _pptx_ph_tipo(sp):
    ph = sp.find('.//p:nvPr/p:ph', _NS)
    return (ph.get('type') or 'body') if ph is not None else ''


def _pptx_pos(el):
    off = el.find('.//a:xfrm/a:off', _NS)
    if off is None:
        return (10 ** 12, 10 ** 12)
    try:
        return (int(off.get('y') or 0), int(off.get('x') or 0))
    except ValueError:
        return (10 ** 12, 10 ** 12)


def _pptx_formas(tree):
    """Formas de conteúdo do slide, em ordem de leitura (cima->baixo,
    esquerda->direita), descendo em grupos e ignorando fallbacks (a imagem
    de fallback de uma equação duplicaria a fórmula)."""
    from lxml import etree
    for fb in list(tree.iter('{%s}Fallback' % _NS['mc'])):
        fb.getparent().remove(fb)
    tree_sp = tree.find('.//p:cSld/p:spTree', _NS)
    if tree_sp is None:
        return []
    saida = []

    def visita(container):
        itens = []
        for el in container:
            tag = el.tag.split('}')[-1]
            if tag in ('sp', 'pic', 'graphicFrame'):
                itens.append((_pptx_pos(el), el))
            elif tag == 'grpSp':
                itens.append((_pptx_pos(el), el))
            elif tag == 'AlternateContent':
                ch = el.find('mc:Choice', _NS)
                if ch is not None:
                    for sub in ch:
                        itens.append((_pptx_pos(sub), sub))
        itens.sort(key=lambda x: x[0])
        for _, el in itens:
            if el.tag.split('}')[-1] == 'grpSp':
                visita(el)
            else:
                saida.append(el)
    visita(tree_sp)
    return saida


def _extract_pptx(z):
    """Slides do professor: título, corpo, tabelas, equações, imagens e
    notas do apresentador, na ordem da apresentação. Mesma resposta do
    docx ({text, images}) para o fluxo não precisar distinguir."""
    import base64 as b64
    from lxml import etree
    slides = _pptx_ordem_slides(z)

    # 1ª passada: imagens repetidas em muitos slides = logo/fundo (descarta)
    uso = {}
    arvores = []
    for parte in slides:
        try:
            tree = etree.fromstring(z.read(parte))
        except Exception:
            arvores.append(None)
            continue
        arvores.append(tree)
        rels = _pptx_rels(z, parte)
        alvos = set()
        for blip in tree.iter('{%s}blip' % _NS['a']):
            alvo = rels.get(blip.get(_R_EMBED) or '', '')
            if 'media/' in alvo:
                alvos.add(alvo)
        for a in alvos:
            uso[a] = uso.get(a, 0) + 1
    decor = {a for a, n in uso.items()
             if n >= _REPETE_MIN_SLIDES and n >= 0.4 * max(1, len(slides))}

    lines, images, seen = [], [], {}
    for idx, (parte, tree) in enumerate(zip(slides, arvores), 1):
        if tree is None:
            continue
        rels = _pptx_rels(z, parte)
        formas = _pptx_formas(tree)
        titulo, corpo = [], []
        for el in formas:
            tag = el.tag.split('}')[-1]
            if tag == 'sp':
                tipo = _pptx_ph_tipo(el)
                if tipo in _PH_IGNORAR:
                    continue
                tx = el.find('p:txBody', _NS)
                if tx is None:
                    continue
                ls = _pptx_txbody(tx)
                if tipo in ('title', 'ctrTitle'):
                    titulo.extend(ls)
                else:
                    corpo.extend(ls)
            elif tag == 'graphicFrame':
                tbl = el.find('.//a:tbl', _NS)
                if tbl is not None:
                    corpo.append('[TABELA]')
                    for tr in tbl.findall('a:tr', _NS):
                        cels = []
                        for tc in tr.findall('a:tc', _NS):
                            tx = tc.find('a:txBody', _NS)
                            cels.append(' '.join(_pptx_txbody(tx)) if tx is not None else '')
                        corpo.append('| ' + ' | '.join(cels) + ' |')
                    corpo.append('[/TABELA]')
                    continue
                # gráfico (chart) ou SmartArt: só o texto que der para ler
                ls = []
                for p in el.iter('{%s}p' % _NS['a']):
                    l = _pptx_paragrafo(p)
                    if l:
                        ls.append(l)
                if ls:
                    corpo.append('[GRÁFICO/ESQUEMA] ' + ' / '.join(ls))
            elif tag == 'pic':
                blip = el.find('.//a:blip', _NS)
                if blip is None:
                    continue
                alvo = rels.get(blip.get(_R_EMBED) or '', '')
                if 'media/' not in alvo or alvo in decor:
                    continue
                if alvo in seen:
                    corpo.append('[IMAGEM %d]' % seen[alvo])
                    continue
                try:
                    raw = z.read(alvo)
                except KeyError:
                    continue
                if len(raw) < 3000:
                    continue
                ext = alvo.rsplit('.', 1)[-1].lower()
                if ext in ('emf', 'wmf', 'svg'):
                    continue  # o builder só embute png/jpeg
                n = len(images) + 1
                seen[alvo] = n
                mime = 'image/png' if ext == 'png' else 'image/jpeg' if ext in ('jpg', 'jpeg') else 'image/' + ext
                images.append({'n': n, 'mime': mime, 'base64': b64.b64encode(raw).decode('ascii')})
                corpo.append('[IMAGEM %d]' % n)
        # notas do apresentador
        notas = []
        for rid, alvo in rels.items():
            if 'notesSlides/' in alvo:
                try:
                    nt = etree.fromstring(z.read(alvo))
                except Exception:
                    continue
                for sp in nt.iter('{%s}sp' % _NS['p']):
                    if _pptx_ph_tipo(sp) in _PH_IGNORAR:
                        continue
                    tx = sp.find('p:txBody', _NS)
                    if tx is not None:
                        notas.extend(_pptx_txbody(tx))
        cab = '===== SLIDE %d%s =====' % (idx, (': ' + ' '.join(titulo)) if titulo else '')
        bloco = [cab] + corpo
        if notas:
            bloco.append('NOTAS DO PROFESSOR: ' + ' '.join(notas))
        lines.append('\n'.join(bloco))
    text = '\n\n'.join(lines)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()
    return {"text": text, "chars": len(text), "images": images,
            "kind": "pptx", "slides": len(slides)}


@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    """Extrai texto e imagens de um .docx (transcrição/material) ou de um
    .pptx (slides do professor). Resposta: {text, chars, images, kind}."""
    data = await file.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        nomes = set(z.namelist())
    except Exception as e:
        raise HTTPException(400, f"arquivo inválido (não é docx/pptx): {e}")
    if 'ppt/presentation.xml' in nomes:
        return _extract_pptx(z)
    if 'word/document.xml' in nomes:
        return _extract_docx(z)
    raise HTTPException(400, "arquivo inválido: esperado .docx ou .pptx")


class RenderRequest(BaseModel):
    data: dict
    filename: str = "material.docx"
    paginate: bool = True
    update_fields: bool = False   # grava updateFields: Word oferece atualizar sumário ao abrir


@app.post("/render")
def render(req: RenderRequest):
    _ensure_fonts()
    try:
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, re.sub(r'[^\w. \-]', '_', req.filename) or 'material.docx')
            status = builder.render(req.data, out, paginate=req.paginate,
                                    update_fields=req.update_fields)
            with open(out, 'rb') as fh:
                blob = fh.read()
    except KeyError as e:
        raise HTTPException(422, f"campo obrigatório ausente no JSON: {e}")
    except Exception as e:
        raise HTTPException(500, f"falha na renderização: {type(e).__name__}: {e}")
    headers = {
        'Content-Disposition': 'attachment; filename="%s"' % os.path.basename(
            re.sub(r'[^\w. \-]', '_', req.filename)),
        'X-Toc-Paginated': 'true' if status.get('toc_paginated') else 'false',
    }
    if status.get('toc_error'):
        headers['X-Toc-Error'] = re.sub(r'[^\x20-\x7e]', '?', status['toc_error'])[:200]
    return Response(
        content=blob,
        media_type=('application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.document'),
        headers=headers)
