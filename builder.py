# -*- coding: utf-8 -*-
"""
Renderizador Águia - Carreiras Policiais
Monta o docx final sobre o template (documento modelo com estilos, fontes
embutidas, cabeçalhos, rodapés e capa), clonando fragmentos XML colhidos do
próprio modelo e injetando o conteúdo estruturado (JSON).
"""
import re
import io
import os
import copy
import shutil
import struct
import zipfile
import tempfile
import subprocess
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.docx")

P_RE = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>', re.S)
T_RE = re.compile(r'<w:t[^>]*>([^<]*)</w:t>')
RUN_RE = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>', re.S)


def _text_of(x):
    return ''.join(T_RE.findall(x))


def _paras(x):
    return P_RE.findall(x)


def _runs(p):
    return RUN_RE.findall(p)


def _rpr(run):
    m = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
    return m.group(0) if m else ''


def _strip_bold(rpr):
    return re.sub(r'<w:b(?:Cs)?(?: [^>]*)?/>', '', rpr)


def _add_bold(rpr):
    if not rpr:
        return '<w:rPr><w:b/></w:rPr>'
    if re.search(r'<w:b\b', rpr):
        return rpr
    return rpr.replace('<w:rPr>', '<w:rPr><w:b/>', 1)


class Frag:
    """Fragmentos colhidos do template."""
    pass


def _split_body(xml):
    """Divide o corpo em elementos de nível superior com um parser real,
    devolvendo cada um como string serializada (suporta w:p aninhado em
    caixas de texto)."""
    from lxml import etree
    root = etree.fromstring(xml.encode('utf-8'))
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    body_el = root.find(W + 'body')
    elems = []
    sectpr = ''
    for child in body_el:
        s = etree.tostring(child, encoding='unicode')
        if child.tag == W + 'sectPr':
            sectpr = s
        else:
            elems.append(s)
    return elems, sectpr


def harvest(template_path=TEMPLATE):
    z = zipfile.ZipFile(template_path)
    xml = z.read('word/document.xml').decode('utf-8')
    header = xml[:xml.index('<w:body>') + len('<w:body>')]
    elems, sectpr_real = _split_body(xml)

    f = Frag()
    f.doc_header = header
    f.doc_tail = '</w:body></w:document>'
    f.sectpr = sectpr_real

    def is_tbl(e):
        return e.lstrip().startswith('<w:tbl')

    def find_tbl(substr, fill=None):
        for e in elems:
            if is_tbl(e) and substr in _text_of(e) and \
                    (fill is None or ('w:fill="%s"' % fill) in e):
                return e
        raise KeyError(substr)

    def find_p(pred):
        for e in elems:
            if not is_tbl(e) and pred(e):
                return e
        raise KeyError('para')

    # ---- capa -------------------------------------------------------------
    f.cover_bg = elems[0]                       # arte de fundo
    f.cover_art = elems[1]                      # 2 caixas de texto com o título

    # ---- sumário (vive dentro de um <w:sdt>) ------------------------------
    sdt = next(e for e in elems if e.lstrip().startswith('<w:sdt')
               and 'CabealhodoSumrio' in e)
    m = re.search(r'^(.*?<w:sdtContent>).*(</w:sdtContent>.*)$', sdt, re.S)
    f.toc_shell_open, f.toc_shell_close = m.group(1), m.group(2)
    sdt_paras = _paras(sdt)
    f.toc_header = next(p for p in sdt_paras if 'CabealhodoSumrio' in p)
    f.toc_e1 = next(p for p in sdt_paras if '"Sumrio1"' in p)
    f.toc_e2 = next(p for p in sdt_paras if '"Sumrio2"' in p)

    # ---- estruturais ------------------------------------------------------
    f.blank = '<w:p/>'
    f.pagebreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    f.h1 = find_p(lambda p: '"Ttulo1"' in p and 'PODER' in _text_of(p))
    f.banner_t5 = find_tbl('APRESENTAÇÃO', '3B3B3B')        # Ttulo5 escuro
    f.banner_t2 = find_tbl('CAPÍTULO 1', '3B3B3B')          # Ttulo2 escuro
    f.subtitle = find_tbl('A quem o poder disciplinar')     # Ttulo4 claro
    f.body_p = find_p(lambda p: 'Este material trata' in _text_of(p))
    f.caption = find_p(lambda p: 'O alcance do poder disciplinar' in _text_of(p))

    # ---- boxes ------------------------------------------------------------
    f.box_mnemonico = find_tbl('MNEMÔNICO', 'F3F0E6')
    f.box_dica = find_tbl('DICA RÁPIDA', 'FBF9F1')
    f.box_lei = find_tbl('LETRA DA LEI', 'F2EFE4')
    f.box_juris = find_tbl('JURISPRUDÊNCIA', 'FAF4E6')
    f.box_aprof = find_tbl('Aprofundando', 'FBF6E4')
    f.box_dialogo = find_tbl('DIÁLOGO EM SALA', 'F7F1DE')
    f.box_diverg = find_tbl('DIVERGÊNCIA', 'F8F2F0')
    f.box_revisao = find_tbl('O QUE ESTUDEI')
    f.tabela = find_tbl('CATEGORIA')
    f.gabarito_tbl = find_tbl('QUESTÃO', 'F3F1E7')

    # ---- questões ---------------------------------------------------------
    f.q_cab = find_p(lambda p: 'CabealhoQuesto' in p)
    f.q_corpo = find_p(lambda p: 'CorpoQuestoteste' in p and _text_of(p).strip()
                       and 'Errado (' not in _text_of(p))
    f.q_marker = find_p(lambda p: 'CorpoQuestoteste' in p and
                        'Errado (' in _text_of(p).replace('\xa0', ' '))
    f.q_corpo_blank = find_p(lambda p: 'CorpoQuestoteste' in p and
                             not _text_of(p).strip())
    f.q_espaco = find_p(lambda p: 'EspaoEntreQuestes' in p)
    f.c_gab = find_p(lambda p: 'CitaoIntensa' in p and
                     'GABARITO:' in _text_of(p))
    f.c_com = find_p(lambda p: 'CitaoIntensa' in p and
                     'COMENTÁRIO:' in _text_of(p))
    f.esp_cap = find_p(lambda p: 'EspaoCapitulos' in p)
    return f


# ---------------------------------------------------------------------------
# geração de runs com **negrito** e highlight FGV
# ---------------------------------------------------------------------------

BLOQUEIO_ZERO = True  # travessão proibido em conteúdo gerado


def _sanitize(text):
    # travessão em conteúdo vindo do modelo (títulos de capítulo) é permitido;
    # o pipeline de IA já entrega sem travessão no corpo. Aqui só normaliza nbsp.
    return text.replace('\u00a0', ' ')


def make_runs(text, base_rpr, bold_rpr=None, highlight_fgv=True):
    """Converte texto com **negrito** em runs; FGV recebe highlight amarelo."""
    if bold_rpr is None:
        bold_rpr = _add_bold(base_rpr)
    out = []
    parts = re.split(r'(\*\*.*?\*\*)', _sanitize(text))
    for part in parts:
        if not part:
            continue
        bold = part.startswith('**') and part.endswith('**') and len(part) > 4
        txt = part[2:-2] if bold else part
        rpr = bold_rpr if bold else base_rpr
        if highlight_fgv and 'FGV' in txt:
            for piece in re.split(r'(FGV)', txt):
                if not piece:
                    continue
                if piece == 'FGV':
                    hp = rpr.replace('<w:rPr>', '<w:rPr><w:highlight w:val="yellow"/>', 1) \
                        if rpr else '<w:rPr><w:highlight w:val="yellow"/></w:rPr>'
                    out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                               % (hp, escape(piece)))
                else:
                    out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                               % (rpr, escape(piece)))
        else:
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                       % (rpr, escape(txt)))
    return ''.join(out)


def retext_para(para_xml, text, keep_first_n_runs=0, highlight_fgv=True):
    """Reaproveita o pPr e o rPr do parágrafo modelo, trocando o texto.
    Runs em negrito do modelo viram o bold_rpr; o primeiro run "normal" vira
    o base_rpr."""
    ppr = re.search(r'<w:pPr>.*?</w:pPr>', para_xml, re.S)
    ppr = ppr.group(0) if ppr else ''
    runs = _runs(para_xml)
    base_rpr, bold_rpr = '', None
    for r in runs:
        rp = _rpr(r)
        if re.search(r'<w:b\b', rp):
            if bold_rpr is None:
                bold_rpr = rp
        elif not base_rpr:
            base_rpr = rp
    if not base_rpr and bold_rpr:
        base_rpr = _strip_bold(bold_rpr)
    kept = ''.join(runs[:keep_first_n_runs])
    body = make_runs(text, base_rpr, bold_rpr, highlight_fgv)
    return '<w:p>%s%s%s</w:p>' % (ppr, kept, body)


def _patch_first_text(xml_frag, new_text):
    """Substitui o texto do fragmento mantendo tudo mais (primeiro w:t recebe
    o novo texto, demais w:t do mesmo parágrafo são esvaziados)."""
    done = [False]

    def rep(m):
        if not done[0]:
            done[0] = True
            return '<w:t xml:space="preserve">%s</w:t>' % escape(new_text)
        return '<w:t xml:space="preserve"></w:t>'
    return T_RE.sub(lambda m: rep(m), xml_frag)


class Bookmarks:
    def __init__(self):
        self.n = 0
        self.items = []  # (name, title, level)

    def new(self, title, level):
        self.n += 1
        name = '_TocRani2_2%03d' % self.n
        self.items.append((name, title, level))
        return name


def _with_bookmark(frag, title, bm_name, bm_id):
    """Troca o texto do título dentro do fragmento e regrava o bookmark."""
    frag = re.sub(r'<w:bookmarkStart [^/]*/>', '', frag)
    frag = re.sub(r'<w:bookmarkEnd [^/]*/>', '', frag)
    # injeta bookmark ao redor dos runs do parágrafo de título
    def inj(m):
        p = m.group(0)
        p = _patch_first_text(p, title)
        p = re.sub(r'(</w:pPr>)',
                   r'\1<w:bookmarkStart w:name="%s" w:id="%d"/>' % (bm_name, bm_id),
                   p, count=1)
        p = p.replace('</w:p>', '<w:bookmarkEnd w:id="%d"/></w:p>' % bm_id)
        return p
    return P_RE.sub(inj, frag, count=1)


# ---------------------------------------------------------------------------
# builders de bloco
# ---------------------------------------------------------------------------

class Builder:
    def __init__(self, frag):
        self.f = frag
        self.bm = Bookmarks()
        self.out = []
        self.bm_id = 100

    # ---------- helpers ----------
    def add(self, xml):
        self.out.append(xml)

    def para(self, text):
        self.add(retext_para(self.f.body_p, text))

    def blank(self):
        self.add(self.f.blank)

    def _box_generic(self, box, texts_by_para, clone_last_for=None):
        """texts_by_para: lista alinhada aos parágrafos do box; None = manter.
        clone_last_for: lista de textos extras clonando o último parágrafo."""
        ps = _paras(box)
        new_ps = []
        for i, p in enumerate(ps):
            t = texts_by_para[i] if i < len(texts_by_para) else None
            if t is _DROP:
                continue
            new_ps.append(retext_para(p, t) if t is not None else p)
        if clone_last_for:
            model = ps[len(texts_by_para) - 1] if texts_by_para else ps[-1]
            for t in clone_last_for:
                new_ps.append(retext_para(model, t))
        # remonta a tabela com os novos parágrafos na mesma célula
        cell = re.search(r'<w:tc>.*?</w:tc>', box, re.S).group(0)
        tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', cell, re.S).group(0)
        new_cell = '<w:tc>%s%s</w:tc>' % (tcpr, ''.join(new_ps))
        return re.sub(r'<w:tc>.*?</w:tc>', lambda m: new_cell, box, count=1, flags=re.S)

    # ---------- estrutura ----------
    def cover(self, titulo):
        self.add(self.f.cover_bg)

        def fix_box(m):
            inner = m.group(0)
            done = [False]

            def rep(mm):
                if not done[0]:
                    done[0] = True
                    return ('<w:t xml:space="preserve">%s</w:t>'
                            % escape(titulo))
                return '<w:t xml:space="preserve"></w:t>'
            return T_RE.sub(rep, inner)

        art = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', fix_box,
                     self.f.cover_art, flags=re.S)
        self.add(art)

    def toc_placeholder(self):
        self._toc_index = len(self.out)
        self.add('%%TOC%%')
        self.add(self.f.blank)
        self.add(self.f.pagebreak)

    def h1(self, titulo):
        name = self.bm.new(titulo, 1)
        self.bm_id += 1
        p = self.f.h1
        p = re.sub(r'<w:bookmarkStart [^/]*/>', '', p)
        p = re.sub(r'<w:bookmarkEnd [^/]*/>', '', p)
        p = _patch_first_text(p, titulo)
        p = re.sub(r'(</w:pPr>)', r'\1<w:bookmarkStart w:name="%s" w:id="%d"/>'
                   % (name, self.bm_id), p, count=1)
        p = p.replace('</w:p>', '<w:bookmarkEnd w:id="%d"/></w:p>' % self.bm_id)
        self.add(p)

    def banner5(self, titulo):
        name = self.bm.new(titulo, 2)
        self.bm_id += 1
        self.add(_with_bookmark(self.f.banner_t5, titulo, name, self.bm_id))

    def banner2(self, titulo, page_break=True):
        if page_break:
            self.add(self.f.pagebreak)
            self.blank()
        name = self.bm.new(titulo, 2)
        self.bm_id += 1
        self.add(_with_bookmark(self.f.banner_t2, titulo, name, self.bm_id))
        self.blank()

    def subtitle(self, titulo):
        frag = P_RE.sub(lambda m: _patch_first_text(m.group(0), titulo),
                        self.f.subtitle, count=1)
        self.add(frag)
        self.blank()

    # ---------- boxes ----------
    def mnemonico(self, titulo, texto):
        self.add(self._box_generic(self.f.box_mnemonico, [None, titulo, texto]))
        self.blank()

    def dica(self, texto):
        self.add(self._box_generic(self.f.box_dica, [None, texto]))
        self.blank()

    def lei(self, fonte, texto):
        parts = [t for t in texto.split('\n') if t.strip()]
        self.add(self._box_generic(self.f.box_lei, [None, fonte, parts[0]],
                                   clone_last_for=parts[1:]))
        self.blank()

    def jurisprudencia(self, tribunal, referencia, texto, observacao=None):
        ps = _paras(self.f.box_juris)
        texts = ['JURISPRUDÊNCIA — %s' % tribunal, referencia, texto]
        if observacao:
            texts.append(observacao)
        else:
            texts.append(_DROP)
        # label mantém rPr, só troca o texto
        box = self.f.box_juris
        new_ps = []
        for i, p in enumerate(ps):
            t = texts[i] if i < len(texts) else _DROP
            if t is _DROP:
                continue
            new_ps.append(retext_para(p, t))
        cell = re.search(r'<w:tc>.*?</w:tc>', box, re.S).group(0)
        tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', cell, re.S).group(0)
        new_cell = '<w:tc>%s%s</w:tc>' % (tcpr, ''.join(new_ps))
        self.add(re.sub(r'<w:tc>.*?</w:tc>', lambda m: new_cell, box,
                        count=1, flags=re.S))
        self.blank()

    def aprofundando(self, titulo, texto):
        parts = [t for t in texto.split('\n') if t.strip()]
        self.add(self._box_generic(self.f.box_aprof,
                                   ['Aprofundando: %s' % titulo, parts[0]],
                                   clone_last_for=parts[1:]))
        self.blank()

    def dialogo(self, falas):
        ps = _paras(self.f.box_dialogo)
        aluno_model, prof_model = ps[1], ps[2]
        new_ps = [ps[0]]
        for fala in falas:
            quem = fala.get('quem', 'Aluno')
            model = aluno_model if quem.lower().startswith('a') else prof_model
            ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
            ppr = ppr.group(0) if ppr else ''
            runs = _runs(model)
            lead_rpr = _rpr(runs[0])
            base_rpr = _rpr(runs[1]) if len(runs) > 1 else _strip_bold(lead_rpr)
            lead = '<w:r>%s<w:t xml:space="preserve">%s: </w:t></w:r>' % (
                lead_rpr, escape(quem))
            body = make_runs(fala['texto'], base_rpr)
            new_ps.append('<w:p>%s%s%s</w:p>' % (ppr, lead, body))
        box = self.f.box_dialogo
        cell = re.search(r'<w:tc>.*?</w:tc>', box, re.S).group(0)
        tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', cell, re.S).group(0)
        new_cell = '<w:tc>%s%s</w:tc>' % (tcpr, ''.join(new_ps))
        self.add(re.sub(r'<w:tc>.*?</w:tc>', lambda m: new_cell, box,
                        count=1, flags=re.S))
        self.blank()

    def divergencia(self, pergunta, posicoes):
        ps = _paras(self.f.box_diverg)
        model = ps[2]
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
        ppr = ppr.group(0) if ppr else ''
        runs = _runs(model)
        lead_rpr = _rpr(runs[0])
        base_rpr = _rpr(runs[1]) if len(runs) > 1 else _strip_bold(lead_rpr)
        new_ps = [ps[0], retext_para(ps[1], pergunta)]
        for pos in posicoes:
            lead = '<w:r>%s<w:t xml:space="preserve">%s: </w:t></w:r>' % (
                lead_rpr, escape(pos['rotulo']))
            body = make_runs(pos['texto'], base_rpr)
            new_ps.append('<w:p>%s%s%s</w:p>' % (ppr, lead, body))
        box = self.f.box_diverg
        cell = re.search(r'<w:tc>.*?</w:tc>', box, re.S).group(0)
        tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', cell, re.S).group(0)
        new_cell = '<w:tc>%s%s</w:tc>' % (tcpr, ''.join(new_ps))
        self.add(re.sub(r'<w:tc>.*?</w:tc>', lambda m: new_cell, box,
                        count=1, flags=re.S))
        self.blank()

    # ---------- tabelas ----------
    def tabela(self, colunas, linhas, legenda=None):
        tbl = self.f.tabela
        header_pr = re.search(r'<w:tblPr>.*?</w:tblPr>', tbl, re.S).group(0)
        rows = re.findall(r'<w:tr(?: [^>]*)?>.*?</w:tr>', tbl, re.S)
        hrow, brow_a = rows[0], rows[1]
        brow_b = rows[2] if len(rows) > 2 else rows[1]

        def rebuild_row(row_model, values):
            cells = re.findall(r'<w:tc>.*?</w:tc>', row_model, re.S)
            trpr = re.search(r'<w:trPr>.*?</w:trPr>', row_model, re.S)
            trpr = trpr.group(0) if trpr else ''
            n = len(values)
            new_cells = []
            for i, v in enumerate(values):
                model = cells[min(i, len(cells) - 1)]
                tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', model, re.S).group(0)
                if n != len(cells):
                    total = 10469
                    tcpr = re.sub(r'<w:tcW w:w="\d+"',
                                  '<w:tcW w:w="%d"' % (total // n), tcpr)
                pmodel = _paras(model)[0]
                new_cells.append('<w:tc>%s%s</w:tc>' % (tcpr,
                                 retext_para(pmodel, v)))
            return '<w:tr>%s%s</w:tr>' % (trpr, ''.join(new_cells))

        grid = re.search(r'<w:tblGrid>.*?</w:tblGrid>', tbl, re.S).group(0)
        n = len(colunas)
        old_cols = re.findall(r'<w:gridCol [^/]*/>', grid)
        if n != len(old_cols):
            grid = '<w:tblGrid>%s</w:tblGrid>' % (
                ('<w:gridCol w:w="%d"/>' % (10469 // n)) * n)
        parts = [header_pr, grid, rebuild_row(hrow, colunas)]
        for i, linha in enumerate(linhas):
            parts.append(rebuild_row(brow_a if i % 2 == 0 else brow_b, linha))
        self.add('<w:tbl>%s</w:tbl>' % ''.join(parts))
        if legenda:
            self.add(retext_para(self.f.caption, legenda))
        self.blank()

    def revisao(self, titulo, linhas):
        tbl = self.f.box_revisao
        rows = re.findall(r'<w:tr(?: [^>]*)?>.*?</w:tr>', tbl, re.S)
        hrow, brow = rows[0], rows[1]

        def build_row(model, tema, itens, extra_first=None):
            cells = re.findall(r'<w:tc>.*?</w:tc>', model, re.S)
            trpr = re.search(r'<w:trPr>.*?</w:trPr>', model, re.S)
            trpr = trpr.group(0) if trpr else ''
            new_cells = []
            # célula 0: (título +) tema
            c0 = cells[0]
            tcpr0 = re.search(r'<w:tcPr>.*?</w:tcPr>', c0, re.S).group(0)
            ps0 = _paras(c0)
            c0new = [tcpr0]
            if extra_first is not None:
                c0new.append(retext_para(ps0[0], extra_first))
                tema_model = ps0[1] if len(ps0) > 1 else ps0[0]
            else:
                tema_model = ps0[0]
            c0new.append(retext_para(tema_model, tema))
            new_cells.append('<w:tc>%s</w:tc>' % ''.join(c0new))
            # célula 1: seta (verbatim)
            new_cells.append(cells[1])
            # célula 2: bullets
            c2 = cells[2]
            tcpr2 = re.search(r'<w:tcPr>.*?</w:tcPr>', c2, re.S).group(0)
            bmodel = _paras(c2)[0]
            bl = ''.join(retext_para(bmodel, '\u2022  %s' % it) for it in itens)
            new_cells.append('<w:tc>%s%s</w:tc>' % (tcpr2, bl))
            return '<w:tr>%s%s</w:tr>' % (trpr, ''.join(new_cells))

        header_pr = re.search(r'<w:tblPr>.*?</w:tblPr>', tbl, re.S).group(0)
        grid = re.search(r'<w:tblGrid>.*?</w:tblGrid>', tbl, re.S).group(0)
        parts = [header_pr, grid]
        parts.append(build_row(hrow, linhas[0]['tema'], linhas[0]['itens'],
                               extra_first=titulo))
        for linha in linhas[1:]:
            parts.append(build_row(brow, linha['tema'], linha['itens']))
        self.add('<w:tbl>%s</w:tbl>' % ''.join(parts))
        self.blank()

    # ---------- imagens ----------
    def imagem(self, ref, legenda=None):
        meta = getattr(self, '_img_meta', {}).get(str(ref))
        if not meta:
            return  # imagem referenciada mas não enviada: ignora sem quebrar
        rid, cx, cy = meta
        self._img_did = getattr(self, '_img_did', 9000) + 1
        self.add(_DRAWING_TPL.format(rid=rid, cx=cx, cy=cy, did=self._img_did))
        if legenda:
            self.add(retext_para(self.f.caption, legenda))
        self.blank()

    # ---------- questões ----------
    def questao(self, cabecalho, corpo, certo_errado=False):
        self.add(retext_para(self.f.q_cab, cabecalho))
        for linha in corpo:
            self.add(retext_para(self.f.q_corpo, linha))
        if certo_errado:
            self.add(self.f.q_marker)
        self.add(self.f.q_corpo_blank)
        self.add(self.f.q_espaco)

    def comentario(self, cabecalho, corpo, gabarito, comentario):
        self.add(retext_para(self.f.q_cab, cabecalho))
        for linha in corpo:
            self.add(retext_para(self.f.q_corpo, linha))
        self.add(retext_para(self.f.c_gab, 'GABARITO: %s' % gabarito))
        self.add(retext_para(self.f.c_com, 'COMENTÁRIO: %s' % comentario))
        self.add(self.f.q_espaco)

    def gabarito(self, entradas):
        tbl = self.f.gabarito_tbl
        rows = re.findall(r'<w:tr(?: [^>]*)?>.*?</w:tr>', tbl, re.S)
        hrow, brow_a = rows[0], rows[1]
        brow_b = rows[2] if len(rows) > 2 else rows[1]
        header_pr = re.search(r'<w:tblPr>.*?</w:tblPr>', tbl, re.S).group(0)
        grid = re.search(r'<w:tblGrid>.*?</w:tblGrid>', tbl, re.S).group(0)
        n = len(entradas)
        nrows = (n + 2) // 3
        cols = []
        idx = 0
        for c in range(3):
            take = min(nrows, n - idx)
            cols.append(entradas[idx:idx + take])
            idx += take

        def build_row(model, vals):
            cells = re.findall(r'<w:tc>.*?</w:tc>', model, re.S)
            trpr = re.search(r'<w:trPr>.*?</w:trPr>', model, re.S)
            trpr = trpr.group(0) if trpr else ''
            new_cells = []
            for i in range(6):
                model_c = cells[min(i, len(cells) - 1)]
                tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', model_c, re.S).group(0)
                pm = _paras(model_c)[0]
                new_cells.append('<w:tc>%s%s</w:tc>' %
                                 (tcpr, retext_para(pm, vals[i])))
            return '<w:tr>%s%s</w:tr>' % (trpr, ''.join(new_cells))

        parts = [header_pr, grid,
                 build_row(hrow, ['QUESTÃO', 'GABARITO'] * 3)]
        for r in range(nrows):
            vals = []
            for c in range(3):
                if r < len(cols[c]):
                    e = cols[c][r]
                    vals += [str(e['n']), str(e['g'])]
                else:
                    vals += ['', '']
            parts.append(build_row(brow_a if r % 2 == 0 else brow_b, vals))
        self.add('<w:tbl>%s</w:tbl>' % ''.join(parts))

    # ---------- sumário ----------
    def render_toc(self):
        e1, e2 = self.f.toc_e1, self.f.toc_e2

        def entry(model, title, page, first=False, last=False):
            ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
            ppr = ppr.group(0) if ppr else ''
            pre = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                   '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" '
                   '\\h \\z \\u </w:instrText></w:r>'
                   '<w:r><w:fldChar w:fldCharType="separate"/></w:r>') if first else ''
            post = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if last else ''
            name = title[0]
            body = ('<w:hyperlink w:history="1" w:anchor="%s">'
                    '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                    '<w:r><w:tab/></w:r>'
                    '<w:r><w:t>%s</w:t></w:r></w:hyperlink>'
                    % (name, escape(title[1]), page))
            return '<w:p>%s%s%s%s</w:p>' % (ppr, pre, body, post)

        items = self.bm.items
        parts = [self.f.toc_shell_open, self.f.toc_header]
        for i, (name, title, level) in enumerate(items):
            model = e1 if level == 1 else e2
            page = self._pages.get(name, '') if hasattr(self, '_pages') else ''
            parts.append(entry(model, (name, title), page,
                               first=(i == 0), last=(i == len(items) - 1)))
        parts.append(self.f.toc_shell_close)
        return ''.join(parts)

    def compose(self):
        xml_body = []
        for i, e in enumerate(self.out):
            if e == '%%TOC%%':
                xml_body.append(self.render_toc())
            else:
                xml_body.append(e)
        return self.f.doc_header + ''.join(xml_body) + self.f.sectpr + self.f.doc_tail


_DROP = object()




# ---------------------------------------------------------------------------
# imagens embutidas no material
# ---------------------------------------------------------------------------

def _img_dims(data):
    """Largura/altura em pixels de PNG ou JPEG; fallback 800x600."""
    try:
        if data[:8] == b'\x89PNG\r\n\x1a\n':
            w = int.from_bytes(data[16:20], 'big')
            h = int.from_bytes(data[20:24], 'big')
            return w, h
        if data[:2] == b'\xff\xd8':
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                    h = int.from_bytes(data[i + 5:i + 7], 'big')
                    w = int.from_bytes(data[i + 7:i + 9], 'big')
                    return w, h
                seg = int.from_bytes(data[i + 2:i + 4], 'big')
                i += 2 + seg
    except Exception:
        pass
    return 800, 600


MAX_CX_EMU = 6096000  # ~16.1 cm, largura útil da página


def prepare_images(imagens, template_path=TEMPLATE):
    """Prepara imagens: decide rIds livres, dimensões e nomes de mídia.
    imagens: dict ref(str) -> {"base64": ..., "mime": "image/png"}.
    Devolve (meta ref->(rid, cx, cy), media_files, rels_extra)."""
    import base64 as _b64
    if not imagens:
        return {}, {}, []
    z = zipfile.ZipFile(template_path)
    rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    next_rid = max(used) + 1 if used else 100
    meta, media, rels_extra = {}, {}, []
    for ref in sorted(imagens.keys(), key=lambda x: int(x)):
        info = imagens[ref]
        raw = _b64.b64decode(info['base64'])
        mime = info.get('mime', 'image/png')
        ext = 'png' if 'png' in mime else ('jpeg' if ('jpeg' in mime or 'jpg' in mime) else 'png')
        w, h = _img_dims(raw)
        cx = int(w * 9525)
        cy = int(h * 9525)
        if cx > MAX_CX_EMU:
            cy = int(cy * MAX_CX_EMU / cx)
            cx = MAX_CX_EMU
        fname = 'media/imagemAG%s.%s' % (ref, ext)
        rid = 'rId%d' % next_rid
        next_rid += 1
        meta[str(ref)] = (rid, cx, cy)
        media['word/' + fname] = raw
        rels_extra.append(
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/image" Target="%s"/>' % (rid, fname))
    return meta, media, rels_extra


_DRAWING_TPL = (
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
    '<wp:inline distT="0" distB="0" distL="0" distR="0" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:docPr id="{did}" name="ImagemAG{did}"/>'
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:nvPicPr><pic:cNvPr id="{did}" name="ImagemAG{did}"/><pic:cNvPicPr/></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


# ---------------------------------------------------------------------------
# documento completo a partir do JSON
# ---------------------------------------------------------------------------

def build_document(data, frag=None, prebuilt=None):
    f = frag or harvest()
    b = prebuilt or Builder(f)
    b.cover(data['titulo'])
    b.toc_placeholder()
    b.h1(data['titulo'])
    b.banner5('APRESENTAÇÃO')
    b.blank()
    for p in data.get('apresentacao', []):
        b.para(p)

    for cap in data.get('capitulos', []):
        b.banner2(cap['titulo'])
        for blk in cap.get('blocos', []):
            t = blk['tipo']
            if t == 'paragrafo':
                b.para(blk['texto'])
            elif t == 'subtitulo':
                b.subtitle(blk['texto'])
            elif t == 'tabela':
                b.tabela(blk['colunas'], blk['linhas'], blk.get('legenda'))
            elif t == 'mnemonico':
                b.mnemonico(blk['titulo'], blk['texto'])
            elif t == 'dica':
                b.dica(blk['texto'])
            elif t == 'lei':
                b.lei(blk['fonte'], blk['texto'])
            elif t == 'jurisprudencia':
                b.jurisprudencia(blk['tribunal'], blk['referencia'],
                                 blk['texto'], blk.get('observacao'))
            elif t == 'aprofundando':
                b.aprofundando(blk['titulo'], blk['texto'])
            elif t == 'dialogo':
                b.dialogo(blk['falas'])
            elif t == 'divergencia':
                b.divergencia(blk['pergunta'], blk['posicoes'])
            elif t == 'revisao':
                b.revisao(blk.get('titulo', 'HORA DE REVISAR'), blk['linhas'])
            elif t == 'imagem':
                b.imagem(blk.get('ref'), blk.get('legenda'))

    b.banner2('QUESTÕES PARA PRATICAR')
    for q in data.get('questoes', []):
        b.questao(q['cabecalho'], q['corpo'], q.get('certo_errado', False))

    b.banner2('QUESTÕES COMENTADAS')
    for q in data.get('comentarios', []):
        b.comentario(q['cabecalho'], q.get('corpo', []), q['gabarito'],
                     q['comentario'])

    b.banner2('GABARITO')
    b.add(f.esp_cap)
    b.gabarito(data['gabarito'])
    return b


def write_docx(builder, out_path, template_path=TEMPLATE, pages=None):
    if pages:
        builder._pages = pages
    doc_xml = builder.compose()
    shutil.copy(template_path, out_path)
    extra = dict(getattr(builder, '_img_media', {}))
    extra['word/document.xml'] = doc_xml.encode('utf-8')
    rels_extra = getattr(builder, '_img_rels', [])
    if rels_extra:
        z = zipfile.ZipFile(template_path)
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rels = rels.replace('</Relationships>', ''.join(rels_extra) + '</Relationships>')
        extra['word/_rels/document.xml.rels'] = rels.encode('utf-8')
    _replace_many_in_zip(out_path, extra)


def _replace_many_in_zip(path, files):
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        done = set()
        for item in zin.infolist():
            if item.filename in files:
                zout.writestr(item, files[item.filename])
                done.add(item.filename)
            else:
                zout.writestr(item, zin.read(item.filename))
        for name, content in files.items():
            if name not in done:
                zout.writestr(name, content)
    os.replace(tmp, path)


def _replace_in_zip(path, name, content):
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == name:
                zout.writestr(item, content)
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# fontes embutidas (deobfuscação) + paginação do sumário via LibreOffice
# ---------------------------------------------------------------------------

def install_embedded_fonts(template_path=TEMPLATE):
    """Extrai as fontes .odttf do template, desofusca e instala no sistema
    para o LibreOffice paginar com a métrica correta."""
    z = zipfile.ZipFile(template_path)
    try:
        rels = z.read('word/_rels/fontTable.xml.rels').decode('utf-8')
        ftable = z.read('word/fontTable.xml').decode('utf-8')
    except KeyError:
        return
    keys = dict(re.findall(r'w:fontKey="\{([0-9A-Fa-f-]+)\}"[^>]*r:id="([^"]+)"',
                           ftable))
    if not keys:
        pairs = re.findall(r'r:id="([^"]+)"[^>]*w:fontKey="\{([0-9A-Fa-f-]+)\}"',
                           ftable)
        keys = {k: rid for rid, k in pairs}
    rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    font_dir = os.path.expanduser('~/.fonts')
    os.makedirs(font_dir, exist_ok=True)
    for key, rid in keys.items():
        target = rel_map.get(rid)
        if not target:
            continue
        data = bytearray(z.read('word/' + target))
        guid = key.replace('-', '')
        kb = bytes.fromhex(guid)[::-1]
        for i in range(32):
            data[i] ^= kb[i % 16]
        out = os.path.join(font_dir, os.path.basename(target) + '.ttf')
        with open(out, 'wb') as fh:
            fh.write(bytes(data))
    subprocess.run(['fc-cache', '-f', font_dir], capture_output=True)


def paginate_toc(builder, docx_path, template_path=TEMPLATE):
    """Converte para PDF com LibreOffice, localiza a página de cada título e
    regrava o sumário com os números reais."""
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf',
             '--outdir', td, docx_path],
            capture_output=True, timeout=300)
        pdf = os.path.join(td, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
        if not os.path.exists(pdf):
            return False
        txt = subprocess.run(['pdftotext', '-layout', pdf, '-'],
                             capture_output=True, text=True).stdout
        pages_txt = txt.split('\f')
        toc_page = next((pi for pi, p in enumerate(pages_txt, 1)
                         if 'SUM' in p and 'RIO' in p), 2)
        pages = {}
        for name, title, level in builder.bm.items:
            norm = re.sub(r'\s+', ' ', title).strip()
            for pi, ptxt in enumerate(pages_txt, 1):
                if pi <= toc_page:
                    continue
                if norm in re.sub(r'\s+', ' ', ptxt):
                    pages[name] = str(pi)
                    break
        builder._pages = pages
        write_docx(builder, docx_path, template_path, pages=pages)
        return True


def render(data, out_path, template_path=TEMPLATE, paginate=True):
    f = harvest(template_path)
    b = Builder(f)
    imagens = data.get('imagens') or {}
    b._img_meta, b._img_media, b._img_rels = prepare_images(imagens, template_path)
    b = build_document(data, f, prebuilt=b)
    write_docx(b, out_path, template_path)
    if paginate:
        try:
            install_embedded_fonts(template_path)
            paginate_toc(b, out_path, template_path)
        except Exception:
            pass
    return out_path
