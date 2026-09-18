# -*- coding: utf-8 -*-
"""Renderizador da máscara RP Concursos (Rani Passos).

Mesmo contrato JSON do renderizador Águia (builder.py): titulo, apresentacao,
capitulos[blocos], formulas, questoes, comentarios, gabarito, imagens.
A diagramação é a do "Novo Padrão 2.0 - RP": capa "ASSUNTO:", página do
professor, sumário, títulos em faixa (Ttulo/Subttulo), corpo Calibri 12 com
1,5, caixa INTERAÇÃO para diálogos, LISTA DE QUESTÕES e LISTA DE QUESTÕES
COMENTADAS (GABARITO/COMENTÁRIOS em magenta).

Reaproveita de builder.py: marcas inline (**negrito**, __sublinhado__),
LaTeX -> OMML, imagens embutidas, ordenação por banca, paginação do sumário
via LibreOffice e gravação do docx.
"""
import os
import re
import zipfile
from xml.sax.saxutils import escape

import builder as B
from builder import (make_runs, _sanitize, _boldify, _ordena_questoes_por_banca,
                     _colhe_formulas, prepare_images, _DRAWING_TPL, write_docx,
                     paginate_toc, _set_update_fields, Bookmarks)

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'template.docx')

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_W = '{%s}' % W_NS

ROXO = '510A51'      # faixa dos títulos
MAGENTA = 'BE074F'   # subtítulos, rótulos, GABARITO/COMENTÁRIOS
LILAS = '64248F'     # bordas
FUNDO = 'F3EFF9'     # fundo das caixas

# largura útil da seção de conteúdo (pgSz 11906 - margens 851/849), em twips
TEXT_W = 10206


# ---------------------------------------------------------------------------
# colheita de fragmentos do template
# ---------------------------------------------------------------------------

class Frag:
    pass


def _strip_ns_decl(s):
    return re.sub(r' xmlns:\w+="[^"]+"', '', s)


def harvest(template_path=TEMPLATE):
    from lxml import etree
    z = zipfile.ZipFile(template_path)
    xml = z.read('word/document.xml').decode('utf-8')
    root = etree.fromstring(xml.encode('utf-8'))
    body = root.find(_W + 'body')
    els = list(body)

    def txt(e):
        return ''.join(t.text or '' for t in e.iter(_W + 't'))

    def ser(e):
        return etree.tostring(e, encoding='unicode')

    f = Frag()
    f.doc_header = xml[:xml.index('<w:body>') + len('<w:body>')]
    f.doc_tail = '</w:body></w:document>'
    f.sectpr = ser(els[-1])                       # seção do conteúdo

    # capa: 1º parágrafo (imagem de fundo + caixa de texto do título + sectPr)
    f.cover = ser(els[0])

    # sumário (sdt): shell + cabeçalho + modelos de entrada + sectPr da seção 2
    sdt = next(e for e in els if e.tag == _W + 'sdt')
    s = ser(sdt)
    m = re.search(r'^(.*?<w:sdtContent>).*(</w:sdtContent>.*)$', s, re.S)
    f.toc_open, f.toc_close = m.group(1), m.group(2)
    paras = list(sdt.iter(_W + 'p'))
    f.toc_header = ser(next(p for p in paras if 'CabealhodoSumrio' in ser(p)))
    f.toc_e1 = ser(next(p for p in paras if '"Sumrio1"' in ser(p)))
    f.toc_e2 = ser(next(p for p in paras if '"Sumrio2"' in ser(p)))
    f.toc_sect = ser(next(p for p in paras if p.find('.//' + _W + 'sectPr') is not None))

    # caixa INTERAÇÃO: parágrafo do rótulo (imagem ancorada) + parágrafo com
    # os avatares e o retângulo arredondado (inline) que contém as falas
    idx = next(i for i, e in enumerate(els)
               if 'pergunta do aluno' in txt(e))
    f.inter_label = ser(els[idx - 1])
    f.inter_box = ser(els[idx])

    # imagem da página do professor (rel adicionada ao template)
    f.prof_rid = 'rIdProf'
    return f


# ---------------------------------------------------------------------------
# helpers de XML
# ---------------------------------------------------------------------------

def _p(style=None, runs='', ppr_extra='', bookmark=None):
    ppr = ''
    if style or ppr_extra:
        ppr = '<w:pPr>%s%s</w:pPr>' % ('<w:pStyle w:val="%s"/>' % style if style else '', ppr_extra)
    if bookmark:
        name, bid = bookmark
        runs = ('<w:bookmarkStart w:id="%d" w:name="%s"/>%s<w:bookmarkEnd w:id="%d"/>'
                % (bid, name, runs, bid))
    return '<w:p>%s%s</w:p>' % (ppr, runs)


def _r(text, rpr=''):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(text))


RPR_LABEL = '<w:rPr><w:rStyle w:val="GabaritoeComentrioChar"/></w:rPr>'
RPR_BOLD = '<w:rPr><w:b/><w:bCs/></w:rPr>'
RPR_CAPTION = '<w:rPr><w:i/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
RPR_GREET = '<w:rPr><w:b/><w:bCs/><w:color w:val="%s"/></w:rPr>' % MAGENTA
RPR_INTER = '<w:rPr><w:b/><w:bCs/><w:color w:val="%s"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>' % MAGENTA

_ANO_RE = re.compile(r'\s*/\s*((?:19|20)\d{2})\s*\)\s*$')


def _cab_rp(cab):
    """'(BANCA / Órgão / 2024)' -> '(BANCA / Órgão – 2024)' (padrão RP)."""
    return _ANO_RE.sub(lambda m: ' \u2013 %s)' % m.group(1), cab)


def _gab_rp(g):
    g = str(g or '').strip()
    m = re.match(r'^\s*(?:letra\s*)?([A-Ea-e])\s*$', g)
    if m:
        return 'LETRA ' + m.group(1).upper()
    return g.upper()


def _cover_with_title(cover_xml, titulo):
    """Troca o texto da caixa de título da capa (todas as caixas: DrawingML e
    fallback VML), mantendo a formatação do 1º run."""
    def fix(m):
        inner = m.group(1)
        ps = re.findall(r'<w:p\b.*?</w:p>', inner, re.S)
        if not ps:
            return m.group(0)
        p0 = ps[0]
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', p0, re.S)
        ppr = ppr.group(0) if ppr else ''
        rpr = re.search(r'<w:r\b[^>]*>(<w:rPr>.*?</w:rPr>)', p0, re.S)
        rpr = rpr.group(1) if rpr else ''
        return '<w:txbxContent>%s</w:txbxContent>' % _p(runs=_r(titulo, rpr), ppr_extra=ppr.replace('<w:pPr>', '').replace('</w:pPr>', ''))
    return re.sub(r'<w:txbxContent>(.*?)</w:txbxContent>', fix, cover_xml, flags=re.S)


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

class Builder:
    def __init__(self, frag):
        self.f = frag
        self.bm = Bookmarks()
        self.out = []
        self.bm_id = 100
        self._did = 9000
        self._img_meta, self._img_media, self._img_rels = {}, {}, []

    # ---------- básicos ----------
    def add(self, xml):
        self.out.append(xml)

    def blank(self):
        self.add('<w:p/>')

    def para(self, text, style=None, ppr_extra=''):
        for i, chunk in enumerate(str(text).split('\n')):
            if not chunk.strip() and i:
                continue
            self.add(_p(style, make_runs(chunk, '', RPR_BOLD), ppr_extra))
        self.blank()

    def _bm(self, title, level):
        name = self.bm.new(title, level)
        self.bm_id += 1
        return name, self.bm_id

    def titulo(self, texto):
        t = _sanitize(str(texto)).strip().upper()
        self.add(_p('Ttulo', _r(t), bookmark=self._bm(t, 1)))
        self.blank()

    def subtitulo(self, texto):
        t = _sanitize(str(texto)).strip().upper()
        self.add(_p('Subttulo', _r(t), bookmark=self._bm(t, 2)))
        self.blank()

    # ---------- capa, professor, sumário ----------
    def cover(self, titulo):
        self.add(_cover_with_title(self.f.cover, _sanitize(str(titulo)).strip()))

    def professor(self):
        self._did += 1
        anchor = (
            '<w:p><w:r><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            'relativeHeight="251700000" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="7560000" cy="10692000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapNone/><wp:docPr id="{did}" name="Professor"/><wp:cNvGraphicFramePr/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{did}" name="Professor"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7560000" cy="10692000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
        ).format(did=self._did, rid=self.f.prof_rid)
        self.add(anchor)
        self.add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def toc_placeholder(self):
        self.add('%%TOC%%')

    def render_toc(self):
        pages = getattr(self, '_pages', {})
        parts = [self.f.toc_open, self.f.toc_header]
        items = self.bm.items
        for i, (name, title, level) in enumerate(items):
            model = self.f.toc_e1 if level == 1 else self.f.toc_e2
            ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
            ppr = ppr.group(0) if ppr else ''
            pre = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                   '<w:r><w:instrText xml:space="preserve"> TOC \\h \\z \\t "Título;1;Subtítulo;2" </w:instrText></w:r>'
                   '<w:r><w:fldChar w:fldCharType="separate"/></w:r>') if i == 0 else ''
            post = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if i == len(items) - 1 else ''
            body = ('<w:hyperlink w:anchor="%s" w:history="1">'
                    '<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF %s \\h </w:instrText></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>%s</w:t></w:r>'
                    '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
                    '</w:hyperlink>' % (name, escape(title), name, pages.get(name, '')))
            parts.append('<w:p>%s%s%s%s</w:p>' % (ppr, pre, body, post))
        parts.append(self.f.toc_sect)     # quebra de seção: conteúdo tem margens próprias
        parts.append(self.f.toc_close)
        return ''.join(parts)

    # ---------- blocos de teoria ----------
    def rotulado(self, rotulo, texto, titulo=None):
        """Parágrafo com rótulo magenta em negrito (DICA:, BIZU:, ...) seguido
        do texto; é o equivalente RP das caixas da máscara Águia."""
        lead = rotulo + (' ' + _sanitize(str(titulo)).strip().upper() if titulo else '') + ':'
        partes = [c for c in str(texto or '').split('\n') if c.strip()] or ['']
        first = _r(lead + ' ', RPR_LABEL) + make_runs(partes[0], '', RPR_BOLD)
        self.add(_p(None, first, '<w:spacing w:before="120"/>'))
        for c in partes[1:]:
            self.add(_p(None, make_runs(c, '', RPR_BOLD)))
        self.blank()

    def dialogo(self, falas):
        """Caixa INTERAÇÃO: rótulo + retângulo arredondado com as falas."""
        falas = [f for f in (falas or []) if f and str(f.get('texto', '')).strip()]
        if not falas:
            return
        paras = []
        n_linhas = 0
        for i, fala in enumerate(falas):
            quem = str(fala.get('quem', '')).strip().lower()
            prof = ('prof' in quem) if quem else (i % 2 == 1)
            texto = _sanitize(str(fala.get('texto', '')))
            n_linhas += max(1, -(-len(texto) // 95))
            ppr = '<w:pStyle w:val="SemEspaamento"/>' + ('<w:jc w:val="right"/>' if prof else '')
            paras.append('<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (ppr, make_runs(texto, RPR_INTER, RPR_INTER)))
        # altura: linhas de 10pt com 1,5 (~15pt) + 24pt de espaçamento por fala + margens internas
        cy = int((n_linhas * 15 + len(falas) * 24 + 10) * 12700)
        cy = max(cy, 700000)
        box = self.f.inter_box
        box = re.sub(r'<w:txbxContent>.*?</w:txbxContent>',
                     '<w:txbxContent>%s</w:txbxContent>' % ''.join(paras), box, flags=re.S)
        box = re.sub(r'(<wp:inline[^>]*>\s*<wp:extent cx="\d+" cy=")\d+(")', r'\g<1>%d\2' % cy, box, count=1)
        box = re.sub(r'(<wps:spPr>.*?<a:ext cx="\d+" cy=")\d+(")', r'\g<1>%d\2' % cy, box, count=1, flags=re.S)
        box = box.replace('<a:noAutofit/>', '<a:spAutoFit/>')
        # avatar do professor (2ª imagem ancorada) desce para o fim da caixa
        anchors = re.findall(r'<wp:anchor\b.*?</wp:anchor>', box, re.S)
        if len(anchors) >= 2:
            a2 = anchors[1]
            novo = re.sub(r'(<wp:positionV[^>]*>\s*<wp:posOffset>)-?\d+(</wp:posOffset>)',
                          r'\g<1>%d\2' % (cy - 520000), a2, count=1)
            box = box.replace(a2, novo, 1)
        # ids de desenho únicos
        for xml_ in ('inter_label',):
            pass
        self.add(self._renum_ids(self.f.inter_label))
        self.add(self._renum_ids(box))
        self.blank()

    def _renum_ids(self, xml):
        def rep(m):
            self._did += 1
            return '%s%d"' % (m.group(1), self._did)
        return re.sub(r'(<wp:docPr id=")\d+"', rep, xml)

    def tabela(self, colunas, linhas, legenda=None):
        colunas = [str(c) for c in (colunas or [])]
        if not colunas:
            return
        n = len(colunas)
        w = TEXT_W // n
        grid = ''.join('<w:gridCol w:w="%d"/>' % w for _ in range(n))
        borders = ''.join('<w:%s w:val="single" w:sz="6" w:space="0" w:color="%s"/>' % (b, LILAS)
                          for b in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
        tblpr = ('<w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>%s</w:tblBorders>'
                 '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
                 '</w:tblPr>' % (TEXT_W, borders))

        def cell(text, header=False):
            rpr = ('<w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
                   if header else '<w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>')
            bold = ('<w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
                    if header else '<w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>')
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % ROXO if header else ''
            p = '<w:p><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/><w:jc w:val="%s"/></w:pPr>%s</w:p>' % (
                'center' if header else 'left', make_runs(str(text if text is not None else ''), rpr, bold))
            return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s<w:vAlign w:val="center"/></w:tcPr>%s</w:tc>' % (w, shd, p)

        rows = ['<w:tr><w:trPr><w:tblHeader/></w:trPr>%s</w:tr>' % ''.join(cell(c, True) for c in colunas)]
        for ln in (linhas or []):
            vals = list(ln) + [''] * (n - len(ln))
            rows.append('<w:tr>%s</w:tr>' % ''.join(cell(v) for v in vals[:n]))
        self.add('<w:tbl>%s<w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (tblpr, grid, ''.join(rows)))
        if legenda:
            self.add(_p(None, make_runs(str(legenda), RPR_CAPTION, RPR_CAPTION), '<w:spacing w:before="60"/><w:jc w:val="center"/>'))
        self.blank()

    def imagem(self, ref, legenda=None):
        meta = self._img_meta.get(str(ref))
        if not meta:
            return
        rid, cx, cy = meta
        self._did += 1
        self.add(_DRAWING_TPL.format(rid=rid, cx=cx, cy=cy, did=self._did))
        if legenda:
            self.add(_p(None, make_runs(str(legenda), RPR_CAPTION, RPR_CAPTION), '<w:jc w:val="center"/>'))
        self.blank()

    def resumo_formulas(self, formulas):
        linhas = []
        for fm in formulas or []:
            if not isinstance(fm, dict):
                continue
            latex = str(fm.get('latex') or '').strip().strip('$').strip()
            if not latex:
                continue
            nome = _sanitize(str(fm.get('nome') or '').strip()) or 'Fórmula'
            quando = _sanitize(str(fm.get('quando') or fm.get('descricao') or '').strip())
            linhas.append([_boldify(nome), '$' + latex + '$', quando])
        if not linhas:
            return False
        self.titulo('RESUMO DAS FÓRMULAS APRESENTADAS')
        self.para('Quadro final com todas as fórmulas trabalhadas neste material, na ordem '
                  'em que aparecem na teoria: use como checklist de revisão antes de resolver as questões.')
        self.tabela(['Fórmula', 'Expressão', 'Quando usar'], linhas)
        return True

    # ---------- questões ----------
    _IMG_MARK = re.compile(r'\[IMAGEM (\d+)\]')
    _ALT_RE = re.compile(r'^\s*([A-Ea-e])\s*[).]\s*(.*)$', re.S)
    _NUM_RE = re.compile(r'^\s*(\d+)\s*[.)\-]?\s*(.*)$', re.S)

    IND_Q = '<w:ind w:left="644" w:hanging="360"/>'
    IND_Q_CORPO = '<w:ind w:left="644"/>'
    IND_ALT = '<w:ind w:left="1440" w:hanging="360"/>'

    def _corpo_lines(self, corpo):
        corpo = B.Builder._limpa_certo_errado(corpo)
        return [str(l).replace('%%', '**') for l in corpo if str(l).strip()]

    def _linha_com_imagens(self, texto, ppr_extra, lead_runs=''):
        marcas = self._IMG_MARK.findall(texto)
        resto = self._IMG_MARK.sub('', texto).strip() if marcas else texto
        if resto or lead_runs:
            self.add(_p('PargrafodaLista', lead_runs + make_runs(resto, '', RPR_BOLD), ppr_extra))
        for ref in marcas:
            self.imagem(ref)

    def questao(self, cabecalho, corpo, certo_errado=False):
        cab, resto = B.Builder._separa_cabecalho(cabecalho)
        m = self._NUM_RE.match(cab)
        num, banca = (m.group(1), m.group(2)) if m else ('', cab)
        banca = _cab_rp(banca.strip())
        linhas = self._corpo_lines(corpo)
        if resto:
            linhas.insert(0, resto)
        # 1ª linha do corpo (comando/enunciado) vai no mesmo parágrafo do número + banca
        primeira = ''
        if linhas and not self._ALT_RE.match(linhas[0]):
            primeira = linhas.pop(0)
        lead = (_r('%s.' % num) + '<w:r><w:tab/></w:r>' if num else '') + _r(banca, RPR_BOLD) + _r(' ')
        self._linha_com_imagens(primeira, self.IND_Q, lead)
        for l in linhas:
            am = self._ALT_RE.match(l)
            if am:
                self.add(_p('PargrafodaLista',
                            _r('%s)' % am.group(1).lower()) + '<w:r><w:tab/></w:r>' + make_runs(am.group(2).strip(), '', RPR_BOLD),
                            self.IND_ALT))
            else:
                self._linha_com_imagens(l, self.IND_Q_CORPO)
        self.add(_p('PargrafodaLista', '', '<w:ind w:left="0"/>'))

    def comentario(self, cabecalho, corpo, gabarito, comentario):
        self.questao(cabecalho, corpo)
        # o questao() fecha com parágrafo vazio; GABARITO vem logo em seguida
        self.out.pop()
        self.add(_p('GabaritoeComentrio', _r('GABARITO: ' + _gab_rp(gabarito), RPR_LABEL),
                    '<w:spacing w:before="120" w:after="120"/><w:ind w:left="644"/>'))
        partes = [c for c in str(comentario or '').replace('%%', '**').split('\n') if c.strip()] or ['']
        self.add(_p('PargrafodaLista', _r('COMENTÁRIOS: ', RPR_LABEL) + make_runs(partes[0], '', RPR_BOLD), self.IND_Q_CORPO))
        for c in partes[1:]:
            self.add(_p('PargrafodaLista', make_runs(c, '', RPR_BOLD), self.IND_Q_CORPO))
        self.add(_p('PargrafodaLista', '', '<w:ind w:left="0"/>'))

    # ---------- montagem ----------
    def compose(self):
        parts = []
        for e in self.out:
            parts.append(self.render_toc() if e == '%%TOC%%' else e)
        return self.f.doc_header + ''.join(parts) + self.f.sectpr + self.f.doc_tail


# ---------------------------------------------------------------------------
# documento a partir do JSON
# ---------------------------------------------------------------------------

_CAP_PREFIX = re.compile(r'^\s*cap[ií]tulo\s*\d+\s*[\-\u2013\u2014:.]?\s*', re.I)


def build_document(data, frag=None, prebuilt=None):
    data = _ordena_questoes_por_banca(data)
    f = frag or harvest()
    b = prebuilt or Builder(f)
    b.cover(data.get('titulo') or 'Material')
    b.professor()
    b.toc_placeholder()

    b.titulo('INTRODUÇÃO')
    for i, p in enumerate(data.get('apresentacao', []) or []):
        if i == 0 and re.match(r'^\s*fala[,!\s]', str(p), re.I) and len(str(p)) < 80:
            b.add(_p(None, _r(_sanitize(str(p)).strip(), RPR_GREET)))
            b.blank()
        else:
            b.para(p)

    for cap in data.get('capitulos', []) or []:
        b.titulo(_CAP_PREFIX.sub('', str(cap.get('titulo') or '')) or 'CAPÍTULO')
        for blk in cap.get('blocos', []) or []:
            t = blk.get('tipo')
            if t == 'paragrafo':
                b.para(blk.get('texto', ''))
            elif t == 'subtitulo':
                b.subtitulo(blk.get('texto', ''))
            elif t == 'tabela':
                b.tabela(blk.get('colunas'), blk.get('linhas'), blk.get('legenda'))
            elif t == 'mnemonico':
                b.rotulado('BIZU', blk.get('texto', ''), blk.get('titulo'))
            elif t == 'dica':
                b.rotulado('DICA', blk.get('texto', ''))
            elif t == 'lei':
                b.rotulado('LETRA DA LEI', blk.get('texto', ''), blk.get('fonte'))
            elif t == 'jurisprudencia':
                ref = ', '.join(x for x in (blk.get('tribunal'), blk.get('referencia')) if x)
                txt = blk.get('texto', '')
                if blk.get('observacao'):
                    txt += '\n' + str(blk['observacao'])
                b.rotulado('JURISPRUDÊNCIA', txt, ref)
            elif t == 'aprofundando':
                b.rotulado('APROFUNDANDO', blk.get('texto', ''), blk.get('titulo'))
            elif t == 'dialogo':
                b.dialogo(blk.get('falas'))
            elif t == 'divergencia':
                txt = '\n'.join('**%s:** %s' % (p.get('rotulo', ''), p.get('texto', ''))
                                for p in (blk.get('posicoes') or []))
                b.rotulado('DIVERGÊNCIA', txt, blk.get('pergunta'))
            elif t == 'revisao':
                txt = '\n'.join('**%s:** %s' % (l.get('tema', ''), '; '.join(l.get('itens') or []))
                                for l in (blk.get('linhas') or []))
                b.rotulado('REVISÃO', txt)
            elif t == 'imagem':
                b.imagem(blk.get('ref'), blk.get('legenda'))

    modo = data.get('resumo_formulas', 'auto')
    if modo is not False:
        formulas = list(data.get('formulas') or [])
        if not formulas and modo == 'auto':
            formulas = _colhe_formulas(data)
        b.resumo_formulas(formulas)

    questoes = data.get('questoes', []) or []
    if questoes:
        b.titulo('LISTA DE QUESTÕES')
        for q in questoes:
            b.questao(q.get('cabecalho', ''), q.get('corpo', []), q.get('certo_errado', False))
    comentarios = data.get('comentarios', []) or []
    if comentarios:
        b.titulo('LISTA DE QUESTÕES COMENTADAS')
        for i, c in enumerate(comentarios):
            corpo = c.get('corpo') or (questoes[i].get('corpo', []) if i < len(questoes) else [])
            b.comentario(c.get('cabecalho', ''), corpo, c.get('gabarito', ''), c.get('comentario', ''))
    return b


def render(data, out_path, template_path=TEMPLATE, paginate=True, update_fields=False):
    f = harvest(template_path)
    b = Builder(f)
    b._img_meta, b._img_media, b._img_rels = prepare_images(data.get('imagens') or {}, template_path)
    b = build_document(data, f, prebuilt=b)
    write_docx(b, out_path, template_path)
    status = {'toc_paginated': False, 'toc_error': None}
    if paginate:
        try:
            if paginate_toc(b, out_path, template_path):
                status['toc_paginated'] = True
            else:
                status['toc_error'] = 'LibreOffice não gerou o PDF de medição'
        except Exception as e:
            status['toc_error'] = '%s: %s' % (type(e).__name__, e)
    if update_fields or (paginate and not status['toc_paginated']):
        try:
            _set_update_fields(out_path)
        except Exception:
            pass
    return status
