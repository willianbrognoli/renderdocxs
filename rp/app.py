# -*- coding: utf-8 -*-
"""Microserviço de renderização RP Concursos (máscara Rani Passos).

Mesmos endpoints e mesmo schema do renderizador Águia; muda apenas a máscara
(builder_rp.py + template.docx desta pasta).

Endpoints:
  GET  /health            -> status
  POST /extract           -> multipart docx, pptx OU pdf -> {"text": "...", "images": [...]}
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
import builder_rp

app = FastAPI(title="Renderizador RP Concursos", version="1.0")

FONTS_READY = False


def _ensure_fonts():
    global FONTS_READY
    if FONTS_READY:
        return
    try:
        builder.install_embedded_fonts(builder_rp.TEMPLATE)
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


# ---------------------------------------------------------------------------
# v7.4: /extract também lê PDF (material/questões do professor exportados em PDF)
# Usa PyMuPDF. Preserva o que dá para recuperar de um PDF:
#   - texto em vermelho -> %%texto%%  (mesma regra de cor do docx)
#   - sublinhado        -> __texto__  (detecta a linha/retângulo desenhado sob o texto)
#   - imagens do corpo  -> [IMAGEM n] no ponto do texto + base64 (>= 3000 bytes)
# Descarta cabeçalho/rodapé repetido (mesma linha nas margens de >= 60% das páginas)
# e números de página soltos. Junta palavras hifenizadas na quebra de linha.
# PDF escaneado (sem camada de texto): tenta OCR se o tesseract estiver no
# container; senão responde 422 explicando.
# ---------------------------------------------------------------------------
_PDF_MARGEM = 0.09          # fração da altura da página considerada cabeçalho/rodapé
_PDF_MIN_IMG = 3000         # bytes; abaixo disso é ícone/decoração (igual ao docx)
_PDF_OCR_MIN_CHARS = 40     # menos que isso em todo o PDF = sem camada de texto


def _pdf_ocr_disponivel():
    """OCR só se o tesseract estiver no PATH e tiver o idioma 'por'."""
    import shutil
    import subprocess
    if not shutil.which('tesseract'):
        return False
    try:
        out = subprocess.run(['tesseract', '--list-langs'], capture_output=True, text=True, timeout=20)
        return 'por' in (out.stdout + out.stderr).split()
    except Exception:
        return False


def _pdf_eh_vermelho(cor_int):
    r_, g_, b_ = (cor_int >> 16) & 0xFF, (cor_int >> 8) & 0xFF, cor_int & 0xFF
    return r_ > 120 and g_ < 90 and b_ < 90


def _pdf_sublinhados(page):
    """Retângulos/linhas horizontais finos desenhados na página (sublinhados
    do Word/LibreOffice saem como 're' preenchido de ~0,5-1pt de altura)."""
    out = []
    try:
        desenhos = page.get_drawings()
    except Exception:
        return out
    for d in desenhos:
        for it in d.get('items', []):
            op = it[0]
            if op == 're':
                r = it[1]
                if r.height <= 2.5 and r.width >= 3:
                    out.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
            elif op == 'l':
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) <= 1.0 and abs(p2.x - p1.x) >= 3:
                    out.append((min(p1.x, p2.x), max(p1.x, p2.x), (p1.y + p2.y) / 2))
    return out


def _pdf_span_sublinhado(bbox, subs):
    x0, y0, x1, y1 = bbox
    largura = max(x1 - x0, 1)
    for sx0, sx1, sy in subs:
        if y1 - 2.5 <= sy <= y1 + 3.5:
            inter = min(x1, sx1) - max(x0, sx0)
            if inter >= 0.5 * largura:
                return True
    return False


def _pdf_linha_marcada(ln, subs):
    """Monta o texto de uma linha (rawdict) marcando por CARACTERE:
    vermelho -> %%..%%, sublinhado -> __..__. O PyMuPDF funde spans vizinhos
    de mesma fonte/cor, então o sublinhado só é confiável no nível do char."""
    pedacos = []            # (char, marca) marca in ('', 'v', 's')
    for sp in ln.get('spans', []):
        verm = _pdf_eh_vermelho(sp.get('color', 0))
        for ch in sp.get('chars', []):
            c = ch.get('c', '')
            if not c:
                continue
            marca = ''
            if c.strip():
                if verm:
                    marca = 'v'
                elif subs and _pdf_span_sublinhado(ch['bbox'], subs):
                    marca = 's'
            pedacos.append((c, marca))
    # espaço entre dois chars com a mesma marca herda a marca (mantém "__a b__" inteiro)
    for i in range(1, len(pedacos) - 1):
        if pedacos[i][1] == '' and not pedacos[i][0].strip() and pedacos[i-1][1] == pedacos[i+1][1] != '':
            pedacos[i] = (pedacos[i][0], pedacos[i-1][1])
    out, atual = [], ''
    tag = {'v': '%%', 's': '__'}
    for c, m in pedacos:
        if m != atual:
            if atual:
                out.append(tag[atual])
            if m:
                out.append(tag[m])
            atual = m
        out.append(c)
    if atual:
        out.append(tag[atual])
    return ''.join(out).rstrip()


def _pdf_imagem_png(raw, ext):
    """Devolve (bytes, mime). PNG/JPEG passam direto; outros formatos são
    reencodados em PNG para o renderizador não engasgar."""
    import pymupdf
    ext = (ext or '').lower()
    if ext in ('png',):
        return raw, 'image/png'
    if ext in ('jpg', 'jpeg'):
        return raw, 'image/jpeg'
    try:
        pix = pymupdf.Pixmap(raw)
        if pix.n - pix.alpha >= 4:          # CMYK -> RGB
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        return pix.tobytes('png'), 'image/png'
    except Exception:
        return None, None


def _extract_pdf(data):
    import base64 as b64
    import hashlib
    try:
        import pymupdf
    except ImportError:
        raise HTTPException(500, "PyMuPDF não instalado no container (pip install pymupdf)")
    try:
        doc = pymupdf.open(stream=data, filetype='pdf')
    except Exception as e:
        raise HTTPException(400, f"pdf inválido: {e}")
    if doc.needs_pass:
        raise HTTPException(400, "pdf protegido por senha")

    # --- 1ª passada: tem camada de texto? --------------------------------
    total = sum(len(p.get_text('text').strip()) for p in doc)
    usar_ocr = False
    if total < _PDF_OCR_MIN_CHARS:
        if _pdf_ocr_disponivel():
            usar_ocr = True
        else:
            raise HTTPException(422, "pdf sem camada de texto (escaneado). Exporte o PDF a partir "
                                     "do Word ou envie o .docx; OCR (tesseract + idioma 'por') "
                                     "não está habilitado no container.")

    paginas = []          # lista de páginas; cada página = lista de (tipo, texto|imgdict, y0, y1)
    images, seen = [], {}

    for page in doc:
        H = page.rect.height
        subs = [] if usar_ocr else _pdf_sublinhados(page)
        if usar_ocr:
            try:
                tp = page.get_textpage_ocr(language='por', full=True, dpi=200)
            except Exception as e:
                raise HTTPException(422, f"OCR falhou na página {page.number + 1}: {e}")
            d = page.get_text('rawdict', textpage=tp)
        else:
            d = page.get_text('rawdict')
        itens = []
        for b in d.get('blocks', []):
            if b.get('type') == 1:                       # imagem
                raw = b.get('image')
                if not raw or len(raw) < _PDF_MIN_IMG:
                    continue
                h = hashlib.md5(raw).hexdigest()
                if h in seen:
                    itens.append(('img', ' [IMAGEM %d]' % seen[h], b['bbox'][1], b['bbox'][3]))
                    continue
                conv, mime = _pdf_imagem_png(raw, b.get('ext'))
                if not conv or len(conv) < _PDF_MIN_IMG:
                    continue
                n = len(images) + 1
                seen[h] = n
                images.append({'n': n, 'mime': mime, 'base64': b64.b64encode(conv).decode('ascii')})
                itens.append(('img', ' [IMAGEM %d]' % n, b['bbox'][1], b['bbox'][3]))
                continue
            linhas_bloco = []
            for ln in b.get('lines', []):
                t = _pdf_linha_marcada(ln, subs)
                if t.strip():
                    linhas_bloco.append((t, ln['bbox'][1], ln['bbox'][3]))
            # junta hifenização na quebra de linha dentro do bloco
            k = 0
            while k < len(linhas_bloco) - 1:
                a, b_ = linhas_bloco[k], linhas_bloco[k + 1]
                if re.search(r'[A-Za-zÀ-ÿ]-$', a[0]) and re.match(r'^[a-zà-ÿ]', b_[0]):
                    linhas_bloco[k] = (a[0][:-1] + b_[0], a[1], b_[2])
                    del linhas_bloco[k + 1]
                else:
                    k += 1
            for t, y0, y1 in linhas_bloco:
                itens.append(('txt', t, y0, y1))
        itens.sort(key=lambda it: (round(it[2], 0), 0))   # ordem de leitura (topo -> base), estável
        paginas.append((H, itens))

    # --- cabeçalho/rodapé repetido e números de página ---------------------
    n_pag = len(paginas)
    contagem = {}
    for H, itens in paginas:
        vistos = set()
        for tipo, t, y0, y1 in itens:
            if tipo != 'txt':
                continue
            if y1 <= H * _PDF_MARGEM or y0 >= H * (1 - _PDF_MARGEM):
                chave = re.sub(r'\d+', '#', t.strip())
                if chave not in vistos:
                    vistos.add(chave)
                    contagem[chave] = contagem.get(chave, 0) + 1
    limiar = max(3, int(0.6 * n_pag + 0.999)) if n_pag >= 3 else 10 ** 9
    repetidos = {k for k, v in contagem.items() if v >= limiar}

    lines = []
    for H, itens in paginas:
        for tipo, t, y0, y1 in itens:
            if tipo == 'txt':
                margem = y1 <= H * _PDF_MARGEM or y0 >= H * (1 - _PDF_MARGEM)
                if margem:
                    chave = re.sub(r'\d+', '#', t.strip())
                    if chave in repetidos or re.fullmatch(r'[\s\-–—]*(página\s*)?\d{1,4}(\s*(de|/)\s*\d{1,4})?[\s\-–—]*', t.strip(), re.I):
                        continue
                lines.append(t)
            else:
                lines.append(t.strip())
        lines.append('')                                   # quebra de página = parágrafo
    text = '\n'.join(lines)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()
    return {"text": text, "chars": len(text), "images": images,
            "kind": "pdf", "pages": n_pag, "ocr": usar_ocr}


@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    """Extrai texto e imagens de um .docx (transcrição/material) ou de um
    .pptx (slides do professor) ou .pdf (v7.4). Resposta: {text, chars, images, kind}."""
    data = await file.read()
    if data[:5] == b'%PDF-' or (file.filename or '').lower().endswith('.pdf'):
        return _extract_pdf(data)
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        nomes = set(z.namelist())
    except Exception as e:
        raise HTTPException(400, f"arquivo inválido (não é docx/pptx/pdf): {e}")
    if 'ppt/presentation.xml' in nomes:
        return _extract_pptx(z)
    if 'word/document.xml' in nomes:
        return _extract_docx(z)
    raise HTTPException(400, "arquivo inválido: esperado .docx, .pptx ou .pdf")


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
            status = builder_rp.render(req.data, out, paginate=req.paginate,
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
