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
    try:
        f.styles = z.read('word/styles.xml').decode('utf-8')
    except KeyError:
        f.styles = ''

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


_TOKEN_RE = re.compile(r'(\*\*.+?\*\*|__.+?__|%%.+?%%)', re.S)

# Símbolos lógicos/matemáticos que nunca recebem negrito. O caso clássico é a
# disjunção exclusiva __∨__ (∨ sublinhado): a marca __ vinha virando negrito +
# sublinhado, mas o glifo deve sair em peso normal, só com o sublinhado.
_SYM_CHARS = set('∨∧⊻¬~→↔⇒⇔∀∃∈∉∅∩∪⊂⊆⊃⊇≤≥≠≈≡∴√±×÷⊕⊗⊥∝∞∂∇∏∑ ()')


def _is_symbol_only(txt):
    """True quando o conteúdo do token é composto apenas por símbolos
    lógicos/matemáticos (sem letras nem dígitos)."""
    t = txt.strip()
    return bool(t) and all(ch in _SYM_CHARS for ch in t)


def _with_props(rpr, bold=False, underline=False, color=None):
    """Deriva um rPr acrescentando negrito/sublinhado/cor ao existente."""
    inner = re.sub(r'^<w:rPr>|</w:rPr>$', '', rpr) if rpr else ''
    if color is not None:
        inner = re.sub(r'<w:color [^/]*/>', '', inner)
        inner += '<w:color w:val="%s"/>' % color
    if bold and '<w:b' not in inner:
        inner = '<w:b/>' + inner
    if underline:
        # remove qualquer u herdado (inclusive w:val="none") e força single
        inner = re.sub(r'<w:u [^/]*/>', '', inner)
        inner += '<w:u w:val="single"/>'
    return '<w:rPr>%s</w:rPr>' % inner if inner else ''


def make_runs(text, base_rpr, bold_rpr=None, highlight_fgv=False):
    """Converte texto com marcas inline em runs:
    **negrito**  __negrito sublinhado__  %%negrito vermelho%%
    Marcas __ e %% aninhadas dentro de **...** também são processadas
    (caso _boldify: rótulos e cabeçalhos de tabela). FGV recebe highlight."""
    if bold_rpr is None:
        bold_rpr = _add_bold(base_rpr)
    out = []
    for part in _TOKEN_RE.split(_sanitize(text)):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            # ** é sempre negrito, inclusive para símbolos (comportamento
            # original). A exceção de símbolo vale apenas para a marca __.
            inner = part[2:-2]
            if '__' in inner or '%%' in inner:
                # parse aninhado: o miolo mantém o contexto de negrito e as
                # marcas internas ganham sublinhado/vermelho por cima dele
                for sub in _TOKEN_RE.split(inner):
                    if not sub:
                        continue
                    if sub.startswith('__') and sub.endswith('__') and len(sub) > 4:
                        stxt = sub[2:-2]
                        # dentro de negrito o símbolo mantém o peso do
                        # contexto (cabeçalho todo bold) e ganha só a barra
                        srpr = _with_props(bold_rpr, underline=True)
                    elif sub.startswith('%%') and sub.endswith('%%') and len(sub) > 4:
                        stxt = sub[2:-2]
                        # sem vermelho: %% rebaixa para negrito, mantendo o peso do contexto
                        srpr = _with_props(bold_rpr, bold=True)
                    else:
                        stxt, srpr = sub, bold_rpr
                    out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                               % (srpr, escape(stxt)))
                continue
            txt, rpr = inner, bold_rpr
        elif part.startswith('__') and part.endswith('__') and len(part) > 4:
            txt = part[2:-2]
            if _is_symbol_only(txt):
                # ex.: __∨__ (disjunção exclusiva): sublinhado, peso normal
                rpr = _with_props(base_rpr, underline=True)
            else:
                rpr = _with_props(bold_rpr, bold=True, underline=True)
        elif part.startswith('%%') and part.endswith('%%') and len(part) > 4:
            # sem vermelho: %% vira apenas negrito
            txt, rpr = part[2:-2], _with_props(base_rpr, bold=True)
        else:
            txt, rpr = part, base_rpr
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


def retext_para(para_xml, text, keep_first_n_runs=0, highlight_fgv=False):
    """Reaproveita o pPr e o rPr do parágrafo modelo, trocando o texto.
    Runs em negrito do modelo viram o bold_rpr; o primeiro run "normal" vira
    o base_rpr."""
    ppr = re.search(r'<w:pPr>.*?</w:pPr>', para_xml, re.S)
    ppr = ppr.group(0) if ppr else ''
    runs = _runs(para_xml)
    base_rpr, bold_rpr = '', None
    base_found = False
    for r in runs:
        rp = _rpr(r)
        if re.search(r'<w:b\b', rp):
            if bold_rpr is None:
                bold_rpr = rp
        elif not base_found and T_RE.search(r):
            # run normal COM texto: seu rPr (mesmo vazio = herda o estilo)
            # é a formatação de corpo correta. rPr vazio NÃO é "não achei".
            base_rpr = rp
            base_found = True
    if not base_found and bold_rpr:
        # só quando o parágrafo modelo é 100% negrito (ex.: rótulos)
        base_rpr = _strip_bold(bold_rpr)
    kept = ''.join(runs[:keep_first_n_runs])
    body = make_runs(text, base_rpr, bold_rpr, highlight_fgv)
    return '<w:p>%s%s%s</w:p>' % (ppr, kept, body)


def _boldify(text):
    """Marca o texto inteiro como negrito (para rótulos/títulos cujo parágrafo
    modelo é todo bold e perderia o negrito no retext)."""
    if text is None:
        return text
    t = str(text)
    if not t.strip() or '**' in t:
        return t
    return '**%s**' % t


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


def _force_jc(frag, val='both'):
    """Garante <w:jc w:val="..."/> em todo parágrafo com texto do fragmento.
    A máscara especifica título e subtítulo de capítulo JUSTIFICADOS, mas os
    parágrafos correspondentes do template vêm sem w:jc (default esquerda)."""
    def fix(m):
        p = m.group(0)
        if not T_RE.search(p):
            return p
        if '<w:jc ' in p:
            return re.sub(r'<w:jc [^/]*/>', '<w:jc w:val="%s"/>' % val, p)
        jc = '<w:jc w:val="%s"/>' % val
        mp = re.search(r'<w:pPr>(.*?)</w:pPr>', p, re.S)
        if mp:
            inner = mp.group(1)
            if '<w:rPr>' in inner:
                inner = inner.replace('<w:rPr>', jc + '<w:rPr>', 1)
            else:
                inner = inner + jc
            return p[:mp.start()] + '<w:pPr>' + inner + '</w:pPr>' + p[mp.end():]
        return re.sub(r'<w:p((?: [^>]*)?)>',
                      r'<w:p\1><w:pPr>' + jc + '</w:pPr>', p, count=1)
    return P_RE.sub(fix, frag)


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

    # Largura útil da caixa de título da capa, em caracteres, na fonte cheia.
    # Calibrado pelos títulos de 2 linhas do modelo ("SENTENÇAS, PROPOSIÇÕES E
    # / MODIFICADORES LÓGICOS"): ~24-26 caracteres por linha. A faixa da capa
    # comporta 2 linhas na fonte cheia; títulos maiores são reduzidos.
    _COVER_CPL = 26.0
    _COVER_MAX_LINES = 2.05   # orçamento de altura em "linhas cheias"
    _COVER_MIN_SCALE = 0.50

    @staticmethod
    def _wrap_count(text, cpl):
        """Número de linhas de um texto com quebra gulosa por palavra."""
        words = str(text).split()
        if not words:
            return 1
        lines, cur = 1, 0
        for w in words:
            need = len(w) if cur == 0 else cur + 1 + len(w)
            if cur and need > cpl:
                lines += 1
                cur = len(w)
            else:
                cur = need
        return lines

    @classmethod
    def _cover_scale(cls, titulo):
        """Maior fator de escala da fonte (1.0 = cheia) em que o título cabe
        no orçamento de altura da faixa da capa. Fonte menor = mais caracteres
        por linha e linhas mais baixas."""
        s = 1.0
        while s > cls._COVER_MIN_SCALE:
            lines = cls._wrap_count(titulo, cls._COVER_CPL / s)
            if lines * s <= cls._COVER_MAX_LINES:
                return s
            s = round(s - 0.05, 2)
        return cls._COVER_MIN_SCALE

    def _style_sz(self, style_id):
        """Tamanho de fonte (w:sz, meios-pontos) definido em um estilo do
        styles.xml do template. None quando não encontrado."""
        st = getattr(self.f, 'styles', '') or ''
        m = re.search(r'<w:style [^>]*w:styleId="%s".*?</w:style>'
                      % re.escape(style_id), st, re.S)
        if not m:
            return None
        s = re.search(r'<w:sz\b[^>]*?w:val="(\d+)"', m.group(0))
        return int(s.group(1)) if s else None

    def _scale_txbx(self, art, scale):
        """Reduz proporcionalmente o título da capa para caber na faixa,
        sempre gravando o tamanho reduzido de forma EXPLÍCITA nos runs
        (conversores como o LibreOffice/Stirling ignoram autofit de caixa
        de texto, então nada de normAutofit).
        Caminho 1: runs com w:sz/w:szCs próprios -> reduz esses valores.
        Caminho 2: fonte herdada de um pStyle (ex.: TtuloCapa) -> lê o
        tamanho base no styles.xml e injeta w:sz/w:szCs inline reduzidos
        em cada run da caixa, sobrepondo o estilo."""
        if scale >= 0.999:
            return art

        def shrink_box(m):
            box = m.group(0)
            hit = [False]

            def sz(mm):
                hit[0] = True
                val = max(4, int(round(int(mm.group(1)) * scale / 2.0)) * 2)
                return mm.group(0).replace('w:val="%s"' % mm.group(1),
                                           'w:val="%d"' % val)

            box = re.sub(r'<w:sz\b[^>]*?w:val="(\d+)"[^>]*?/>', sz, box)
            box = re.sub(r'<w:szCs\b[^>]*?w:val="(\d+)"[^>]*?/>', sz, box)

            def line(mm):
                val = max(40, int(round(int(mm.group(1)) * scale)))
                return mm.group(0).replace('w:line="%s"' % mm.group(1),
                                           'w:line="%d"' % val)

            box = re.sub(r'<w:spacing\b[^>]*?w:line="(\d+)"[^>]*?'
                         r'w:lineRule="(?:exact|atLeast)"[^>]*?/>', line, box)

            if not hit[0]:
                # sem tamanho inline: resolve o base pelo pStyle da caixa
                base = None
                mps = re.search(r'<w:pStyle w:val="([^"]+)"', box)
                if mps:
                    base = self._style_sz(mps.group(1))
                base = base or 44
                val = max(4, int(round(base * scale / 2.0)) * 2)
                szxml = ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                         % (val, val))

                def fix_run(rm):
                    run = rm.group(0)
                    if '<w:rPr>' in run:
                        return run.replace('<w:rPr>', '<w:rPr>' + szxml, 1)
                    return re.sub(r'(<w:r(?: [^>]*)?>)',
                                  lambda o: o.group(1) +
                                  '<w:rPr>' + szxml + '</w:rPr>',
                                  run, count=1)
                box = RUN_RE.sub(fix_run, box)
            return box

        return re.sub(r'<w:txbxContent>.*?</w:txbxContent>', shrink_box,
                      art, flags=re.S)

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
        # título longo (3+ linhas na fonte cheia): reduz a fonte das caixas
        # de título para caber na faixa da capa sem corte
        art = self._scale_txbx(art, self._cover_scale(titulo))
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

    def trim_blanks(self):
        """Remove parágrafos vazios sobrando no fim do fluxo (eles são a
        causa clássica de página em branco quando uma tabela termina cravada
        no pé da página e o próximo elemento abre página nova)."""
        _kn = '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'
        while self.out and self.out[-1] in (self.f.blank, self.f.pagebreak, _kn):
            self.out.pop()

    def banner2(self, titulo, page_break=True):
        if page_break:
            # v7.1: quebra embutida no respiro via pageBreakBefore. Se o
            # conteúdo anterior encheu a página exata, o pageBreakBefore no
            # topo da página seguinte vira no-op e NÃO gera página em branco
            # (o parágrafo-quebra avulso gerava).
            self.trim_blanks()
            self.add('<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>')
        name = self.bm.new(titulo, 2)
        self.bm_id += 1
        self.add(_force_jc(_with_bookmark(self.f.banner_t2, titulo, name, self.bm_id)))
        self.blank()

    def subtitle(self, titulo):
        """O subtítulo do template é uma tabela de 1 célula (borda dourada),
        e tabela não tem keepNext: título ficava órfão no pé da página.
        Convertemos para parágrafo com pBdr idêntico (mesma borda/cor do
        template) e keepNext encadeado até o conteúdo seguinte."""
        frag = P_RE.sub(lambda m: _patch_first_text(m.group(0), titulo),
                        self.f.subtitle, count=1)
        m_bord = re.search(r'<w:top w:val="single" w:color="([0-9A-Fa-f]{6})" '
                           r'w:sz="(\d+)"', frag)
        cor, sz = (m_bord.group(1), m_bord.group(2)) if m_bord else ('C9A227', '10')
        m_p = re.search(r'<w:p(?: [^>]*)?>.*?</w:p>', frag, re.S)
        p = m_p.group(0) if m_p else '<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % escape(titulo)
        borda = ('<w:pBdr>' + ''.join(
            '<w:%s w:val="single" w:sz="%s" w:space="8" w:color="%s"/>'
            % (lado, sz, cor) for lado in ('top', 'left', 'bottom', 'right'))
            + '</w:pBdr>')
        # respiro assimétrico: separa do assunto anterior (antes) e gruda no
        # conteúdo do próprio subtítulo (depois)
        espac = '<w:spacing w:before="360" w:after="120"/>'
        # a borda fica w:space="8" (160 twips) para fora do texto; recuar o
        # parágrafo em +160 à esquerda e -160 à direita alinha a LINHA da caixa
        # à margem do corpo do texto (que tem recuo zero)
        recuo = '<w:ind w:left="160" w:right="160"/>'
        extras = '<w:keepNext/>' + espac + recuo + borda
        if '<w:pPr>' in p:
            if '<w:pStyle' in p:
                p = re.sub(r'(<w:pStyle [^/]*/>)', r'\1' + extras, p, count=1)
            else:
                p = p.replace('<w:pPr>', '<w:pPr>' + extras, 1)
        else:
            p = re.sub(r'<w:p((?: [^>]*)?)>',
                       r'<w:p\1><w:pPr>' + extras + '</w:pPr>', p, count=1)
        self.add(_force_jc(p))

    # ---------- boxes ----------
    def mnemonico(self, titulo, texto):
        box = self._box_generic(self.f.box_mnemonico, [None, _boldify(titulo), texto])
        self.add(box.replace('<w:color w:val="7A2E2E"', '<w:color w:val="000000"'))
        self.blank()

    def dica(self, texto):
        # sem vermelho na dica: %%alerta%% rebaixa para negrito preto
        texto = str(texto).replace('%%', '**')
        self.add(self._box_generic(self.f.box_dica, [None, texto]))
        self.blank()

    def lei(self, fonte, texto):
        parts = [t for t in texto.split('\n') if t.strip()]
        self.add(self._box_generic(self.f.box_lei, [None, _boldify(fonte), parts[0]],
                                   clone_last_for=parts[1:]))
        self.blank()

    def jurisprudencia(self, tribunal, referencia, texto, observacao=None):
        ps = _paras(self.f.box_juris)
        texts = [_boldify('JURISPRUDÊNCIA — %s' % tribunal), _boldify(referencia), texto]
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
                                   [_boldify('Aprofundando: %s' % titulo), parts[0]],
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
        new_ps = [ps[0], retext_para(ps[1], _boldify(pergunta))]
        for pos in posicoes:
            lead = '<w:r>%s<w:t xml:space="preserve">%s: </w:t></w:r>' % (
                lead_rpr, escape(pos['rotulo']))
            body = make_runs(pos['texto'], base_rpr)
            new_ps.append('<w:p>%s%s%s</w:p>' % (ppr, lead, body))
        box = self.f.box_diverg
        cell = re.search(r'<w:tc>.*?</w:tc>', box, re.S).group(0)
        tcpr = re.search(r'<w:tcPr>.*?</w:tcPr>', cell, re.S).group(0)
        new_cell = '<w:tc>%s%s</w:tc>' % (tcpr, ''.join(new_ps))
        out = re.sub(r'<w:tc>.*?</w:tc>', lambda m: new_cell, box,
                     count=1, flags=re.S)
        self.add(out.replace('<w:color w:val="7A2E2E"', '<w:color w:val="000000"'))
        self.blank()

    # ---------- tabelas ----------
    def tabela(self, colunas, linhas, legenda=None):
        tbl = self.f.tabela
        # A borda interna vertical (insideV) fina (sz 8 = 1pt) picota na
        # rasterização de alguns leitores de PDF (cai no sub-pixel e some por
        # trechos). Subir só ela para sz 12 (1,5pt) mantém o visual e elimina
        # o defeito, sem tocar nas demais bordas.
        tbl = re.sub(r'(<w:insideV w:val="single" w:color="[0-9A-Fa-f]+" w:sz=")\d+(")',
                     r'\g<1>12\g<2>', tbl)
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
        parts = [header_pr, grid, rebuild_row(hrow, [_boldify(c) for c in colunas])]
        for i, linha in enumerate(linhas):
            parts.append(rebuild_row(brow_a if i % 2 == 0 else brow_b, linha))
        self.add('<w:tbl>%s</w:tbl>' % ''.join(parts))
        if legenda:
            self.add(retext_para(self.f.caption, legenda))
        self.blank()

    def revisao(self, titulo, linhas):
        """Layout v4 (PDF de referência): moldura dourada do template com o
        título centrado, e por dentro uma tabela de 2 colunas: chip escuro com
        o tema em CAIXA ALTA centrado + itens com seta dourada, em fonte de
        corpo e linhas compactas."""
        from lxml import etree
        W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        wrapper = (self.f.doc_header + self.f.box_revisao +
                   '</w:body></w:document>')
        root = etree.fromstring(wrapper.encode('utf-8'))
        body = root.find(W + 'body')
        outer = body.find(W + 'tbl')
        outer_cell = outer.find(W + 'tr').find(W + 'tc')
        inner = outer_cell.find(W + 'tbl')

        # 1) título: primeiro w:p da célula externa, texto novo e centrado
        for p in outer_cell.findall(W + 'p'):
            ts = p.findall('.//' + W + 't')
            if ts:
                ts[0].text = _sanitize(titulo)
                for t in ts[1:]:
                    t.text = ''
                ppr = p.find(W + 'pPr')
                if ppr is None:
                    ppr = etree.SubElement(p, W + 'pPr')
                    p.remove(ppr)
                    p.insert(0, ppr)
                jc = ppr.find(W + 'jc')
                if jc is None:
                    jc = etree.SubElement(ppr, W + 'jc')
                jc.set(W + 'val', 'center')
                break

        # 1b) moldura externa: estica até a largura do texto comum (sectPr)
        textw = 0
        m_sz = re.search(r'<w:pgSz [^/]*w:w="(\d+)"', self.f.sectpr or '')
        m_mg = re.search(r'<w:pgMar [^/]*/>', self.f.sectpr or '')
        if m_sz and m_mg:
            mg = m_mg.group(0)
            left = re.search(r'w:left="(\d+)"', mg)
            right = re.search(r'w:right="(\d+)"', mg)
            if left and right:
                textw = int(m_sz.group(1)) - int(left.group(1)) - int(right.group(1))
        if textw > 0:
            tblpr = outer.find(W + 'tblPr')
            if tblpr is not None:
                tw = tblpr.find(W + 'tblW')
                if tw is None:
                    tw = etree.SubElement(tblpr, W + 'tblW')
                tw.set(W + 'w', str(textw))
                tw.set(W + 'type', 'dxa')
            og = outer.find(W + 'tblGrid')
            if og is not None:
                cols = og.findall(W + 'gridCol')
                if len(cols) == 1:
                    cols[0].set(W + 'w', str(textw))
            tcpr = outer_cell.find(W + 'tcPr')
            if tcpr is not None:
                tcw = tcpr.find(W + 'tcW')
                if tcw is None:
                    tcw = etree.SubElement(tcpr, W + 'tcW')
                tcw.set(W + 'w', str(textw))
                tcw.set(W + 'type', 'dxa')

        # 2) largura interna: baseia na moldura, com respiro simétrico
        grid = inner.find(W + 'tblGrid')
        inner_total = sum(int(g.get(W + 'w', '0') or 0)
                          for g in grid.findall(W + 'gridCol'))
        outer_grid = outer.find(W + 'tblGrid')
        outer_total = sum(int(g.get(W + 'w', '0') or 0)
                          for g in outer_grid.findall(W + 'gridCol')) \
            if outer_grid is not None else 0
        total = (textw or outer_total or inner_total or 10229) - 700
        tema_w = 2900
        item_w = total - tema_w

        _sp = ('<w:pPr><w:spacing w:before="20" w:after="20" w:line="240" '
               'w:lineRule="auto"/><w:ind w:left="240" w:hanging="240"/>%s'
               '</w:pPr>')
        rows = []
        n_linhas = len(linhas)
        for idx_l, linha in enumerate(linhas):
            itens = ''.join(
                '<w:p>%s<w:r><w:rPr><w:b/><w:color w:val="000000"/></w:rPr>'
                '<w:t xml:space="preserve">\u21d2 </w:t></w:r>%s</w:p>'
                % (_sp % '<w:jc w:val="left"/>',
                   make_runs(str(it), '', highlight_fgv=False))
                for it in linha['itens'])
            tema_p = ('<w:p><w:pPr><w:spacing w:before="20" w:after="20" '
                      'w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>'
                      '</w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/>'
                      '<w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">%s'
                      '</w:t></w:r></w:p>'
                      % escape(_sanitize(str(linha['tema'])).upper()))
            b_top = ('' if idx_l == 0 else
                     '<w:top w:val="single" w:sz="24" w:color="F1E9D2"/>')
            b_bot = ('' if idx_l == n_linhas - 1 else
                     '<w:bottom w:val="single" w:sz="24" w:color="F1E9D2"/>')
            tema_tc = ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>'
                       '<w:tcBorders>%s%s'
                       '<w:right w:val="single" w:sz="48" w:color="F1E9D2"/>'
                       '</w:tcBorders>'
                       '<w:shd w:val="clear" w:color="auto" w:fill="3B3B3B"/>'
                       '<w:vAlign w:val="center"/></w:tcPr>%s</w:tc>'
                       % (tema_w, b_top, b_bot, tema_p))
            item_tc = ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>'
                       '<w:shd w:val="clear" w:color="auto" w:fill="F1E9D2"/>'
                       '<w:vAlign w:val="center"/>'
                       '<w:tcMar><w:left w:w="140" w:type="dxa"/></w:tcMar>'
                       '</w:tcPr>%s</w:tc>' % (item_w, itens))
            rows.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>%s%s</w:tr>'
                        % (tema_tc, item_tc))
        # linha-respiro invisível: garante o vão inferior da moldura
        _p0 = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
               'w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>')
        rows.append('<w:tr><w:trPr><w:cantSplit/>'
                    '<w:trHeight w:val="300" w:hRule="exact"/></w:trPr>'
                    '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>'
                    '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>'
                    '</w:tr>' % (tema_w, _p0, item_w, _p0))

        new_inner_xml = ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                         '<w:jc w:val="center"/>'
                         '<w:tblLayout w:type="fixed"/>'
                         '<w:tblCellMar><w:top w:w="30" w:type="dxa"/>'
                         '<w:left w:w="60" w:type="dxa"/>'
                         '<w:bottom w:w="30" w:type="dxa"/>'
                         '<w:right w:w="60" w:type="dxa"/></w:tblCellMar>'
                         '</w:tblPr>'
                         '<w:tblGrid><w:gridCol w:w="%d"/><w:gridCol w:w="%d"/>'
                         '</w:tblGrid>%s</w:tbl>'
                         % (total, tema_w, item_w, ''.join(rows)))
        new_inner = etree.fromstring(
            (self.f.doc_header + new_inner_xml + '</w:body></w:document>')
            .encode('utf-8')).find(W + 'body').find(W + 'tbl')
        outer_cell.replace(inner, new_inner)

        # respiro inferior simétrico: parágrafo espaçador após a tabela interna
        SP = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
              'w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>')
        idx_inner = list(outer_cell).index(new_inner)
        resto = list(outer_cell)[idx_inner + 1:]
        for el in resto:
            if el.tag == W + 'p':
                outer_cell.remove(el)
        sp_el = etree.fromstring(
            (self.f.doc_header + SP + '</w:body></w:document>')
            .encode('utf-8')).find(W + 'body').find(W + 'p')
        outer_cell.append(sp_el)

        self.add(etree.tostring(outer, encoding='unicode'))
        self.blank()

    def _revisao_legacy(self, titulo, linhas):
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
    _IMG_MARK = re.compile(r'\[IMAGEM (\d+)\]')

    def _linha_questao(self, modelo, linha):
        """Renderiza uma linha de questão; marcadores [IMAGEM N] viram a
        imagem real no ponto correspondente (questões com texto de apoio
        visual, charge, campanha etc.)."""
        s = str(linha)
        marcas = self._IMG_MARK.findall(s)
        if not marcas:
            self.add(retext_para(modelo, s, highlight_fgv=False))
            return
        resto = self._IMG_MARK.sub('', s).strip()
        if resto:
            self.add(retext_para(modelo, resto, highlight_fgv=False))
        for ref in marcas:
            self.imagem(ref)

    def questao(self, cabecalho, corpo, certo_errado=False):
        self.add(retext_para(self.f.q_cab, cabecalho, highlight_fgv=False))
        for linha in corpo:
            self._linha_questao(self.f.q_corpo, linha)
        self.add(self.f.q_corpo_blank)
        self.add(self.f.q_espaco)

    def _lead_para(self, model, lead_text, body_text):
        """Parágrafo com rótulo em negrito (rPr clonado do primeiro run do
        modelo, ex.: "GABARITO: ") seguido do corpo em formatação normal,
        com suporte a **negrito** e highlight FGV."""
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
        ppr = ppr.group(0) if ppr else ''
        runs = _runs(model)
        lead_rpr = _rpr(runs[0]) if runs else '<w:rPr><w:b/></w:rPr>'
        base_rpr = ''
        for r in runs[1:]:
            rp = _rpr(r)
            if not re.search(r'<w:b\b', rp) and T_RE.search(r):
                base_rpr = rp
                break
        lead = '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (
            lead_rpr, escape(_sanitize(lead_text)))
        body = make_runs(body_text, base_rpr, highlight_fgv=False)
        return '<w:p>%s%s%s</w:p>' % (ppr, lead, body)

    def comentario(self, cabecalho, corpo, gabarito, comentario):
        self.add(retext_para(self.f.q_cab, cabecalho, highlight_fgv=False))
        for linha in corpo:
            self._linha_questao(self.f.q_corpo, linha)
        self.add(self._lead_para(self.f.c_gab, 'GABARITO: ', gabarito))
        self.add(self._lead_para(self.f.c_com, 'COMENTÁRIO: ', comentario))
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
                 build_row(hrow, [_boldify('QUESTÃO'), _boldify('GABARITO')] * 3)]
        for r in range(nrows):
            vals = []
            for c in range(3):
                if r < len(cols[c]):
                    e = cols[c][r]
                    vals += [_boldify(str(e['n'])), _boldify(str(e['g']))]
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
                b.revisao('O QUE ESTUDEI', blk['linhas'])
            elif t == 'imagem':
                b.imagem(blk.get('ref'), blk.get('legenda'))

    b.banner2('QUESTÕES PARA PRATICAR')
    for q in data.get('questoes', []):
        b.questao(q['cabecalho'], q['corpo'], q.get('certo_errado', False))

    b.banner2('QUESTÕES COMENTADAS')
    for q in data.get('comentarios', []):
        b.comentario(q['cabecalho'], q.get('corpo', []), q['gabarito'],
                     q['comentario'])

    b.banner2('GABARITO FINAL')
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

def _sfnt_info(path):
    """Lê nomes (família legada, família tipográfica, nome completo, subfamília)
    e estilo (bold/itálico) de um TTF/OTF pelo name table e OS/2."""
    with open(path, 'rb') as fh:
        data = fh.read()
    try:
        num = struct.unpack('>H', data[4:6])[0]
        tables = {}
        for i in range(num):
            off = 12 + 16 * i
            tag = data[off:off + 4].decode('latin1')
            toff, tlen = struct.unpack('>II', data[off + 8:off + 16])
            tables[tag] = (toff, tlen)

        def _name(nid):
            toff, _ = tables['name']
            cnt, stroff = struct.unpack('>HH', data[toff + 2:toff + 6])
            best = None
            for i in range(cnt):
                r = toff + 6 + 12 * i
                pid, eid, lid, nid_, ln, so = struct.unpack('>6H', data[r:r + 12])
                if nid_ != nid:
                    continue
                raw = data[toff + stroff + so:toff + stroff + so + ln]
                try:
                    val = raw.decode('utf-16-be') if pid in (0, 3) else raw.decode('latin1')
                except Exception:
                    continue
                score = 0 if (pid, lid) == (3, 0x409) else 1
                if best is None or score < best[0]:
                    best = (score, val)
            return (best[1] if best else '').strip()

        fam1, fam16 = _name(1), _name(16)
        full4 = _name(4)
        sub = (_name(17) or _name(2)).strip()
        bold = italic = False
        if 'OS/2' in tables:
            toff, _ = tables['OS/2']
            fs = struct.unpack('>H', data[toff + 62:toff + 64])[0]
            italic = bool(fs & 0x01)
            bold = bool(fs & 0x20)
        s = sub.lower()
        bold = bold or 'bold' in s
        italic = italic or 'italic' in s or 'oblique' in s
        return {'fam1': fam1, 'fam16': fam16, 'full': full4, 'sub': sub,
                'bold': bold, 'italic': italic}
    except Exception:
        return {'fam1': '', 'fam16': '', 'full': '', 'sub': '',
                'bold': False, 'italic': False}


def _retag_font(data, family, sub, bold, italic):
    """Reescreve o name table (e flags de estilo) da fonte para que ela se
    registre exatamente com a família que o documento usa. Necessário quando o
    doc usa o nome do estilo como família (ex.: 'Winner Sans Wide Bold'):
    sem isso o LibreOffice registra a embutida pelo nome interno original e
    não encontra a família pedida."""
    data = bytearray(data)
    num = struct.unpack('>H', data[4:6])[0]
    tables = {}
    for i in range(num):
        off = 12 + 16 * i
        tag = data[off:off + 4].decode('latin1')
        toff, tlen = struct.unpack('>II', data[off + 8:off + 16])
        tables[tag] = (off, toff, tlen)
    full = family if sub.lower() == 'regular' else family + ' ' + sub
    ps = (re.sub(r'[^A-Za-z0-9]', '', family) + '-' +
          re.sub(r'[^A-Za-z0-9]', '', sub))
    entries = {1: family, 2: sub, 4: full, 6: ps, 16: family, 17: sub}
    recs, strings = [], b''
    for nid in sorted(entries):
        s = entries[nid].encode('utf-16-be')
        recs.append(struct.pack('>6H', 3, 1, 0x409, nid, len(s), len(strings)))
        strings += s
    name = (struct.pack('>HHH', 0, len(recs), 6 + 12 * len(recs)) +
            b''.join(recs) + strings)
    if len(name) % 4:
        name += b'\x00' * (4 - len(name) % 4)
    off, _, _ = tables['name']
    newoff = len(data)
    data += name
    data[off + 8:off + 16] = struct.pack('>II', newoff, len(name))
    if 'OS/2' in tables:
        _, t, _ = tables['OS/2']
        data[t + 4:t + 6] = struct.pack('>H', 700 if bold else 400)
        fs = (0x01 if italic else 0) | (0x20 if bold else 0)
        if not bold and not italic:
            fs |= 0x40
        data[t + 62:t + 64] = struct.pack('>H', fs)
    if 'head' in tables:
        _, t, _ = tables['head']
        mac = (1 if bold else 0) | (2 if italic else 0)
        data[t + 44:t + 46] = struct.pack('>H', mac)
    return bytes(data)


def _brand_font_dirs():
    dirs = []
    for d in (os.path.join(HERE, 'fonts'),
              '/usr/share/fonts/truetype/aguia'):
        if os.path.isdir(d):
            dirs.append(d)
    return dirs


def embed_brand_fonts(docx_path):
    """Embute as fontes reais da marca (pasta fonts/ do serviço) dentro do
    docx gerado, para o arquivo renderizar certo em qualquer máquina e na
    conversão para PDF (o template só embute as fontes Microsoft)."""
    import uuid
    dirs = _brand_font_dirs()
    if not dirs:
        return 0
    z = zipfile.ZipFile(docx_path)
    ftable = z.read('word/fontTable.xml').decode('utf-8')
    try:
        rels = z.read('word/_rels/fontTable.xml.rels').decode('utf-8')
    except KeyError:
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships"></Relationships>')
    ctypes = z.read('[Content_Types].xml').decode('utf-8')
    settings = z.read('word/settings.xml').decode('utf-8')
    z.close()

    doc_fonts = re.findall(r'<w:font w:name="([^"]+)"', ftable)
    norm = lambda s: re.sub(r'\s+', ' ', s).strip().lower()
    doc_map = {norm(n): n for n in doc_fonts}

    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)] or [0]
    next_rid = max(used) + 1
    extra, added = {}, 0
    ja = [int(m) for m in re.findall(r'fontAG(\d+)\.odttf', rels)]
    seq = max(ja) if ja else 100
    filled = set()  # (familia_doc, slot) já preenchidos
    NEUTRAS = {'', 'regular', 'normal', 'roman', 'italic', 'bold',
               'bold italic', 'italic bold', 'oblique', 'bold oblique'}
    arquivos = []
    for d in dirs:
        for fn in sorted(os.listdir(d)):
            if fn.lower().endswith(('.ttf', '.otf')):
                arquivos.append(os.path.join(d, fn))
    for caminho in arquivos:
        info = _sfnt_info(caminho)
        sub_n = norm(info['sub'])
        # 1) match pela família pura (o doc usa a família e alterna b/i)
        target, merged = None, False
        for k in (info['fam16'], info['fam1']):
            if k and norm(k) in doc_map:
                target = doc_map[norm(k)]
                break
        if target is not None and sub_n not in NEUTRAS:
            target = None  # peso extra (Light, Medium...) não ocupa slot
        # 2) match pelo nome completo / família+subfamília (o doc usa o nome
        #    do estilo como se fosse a família, ex.: "Winner Sans Wide Bold")
        if target is None:
            candidatos = [info['full']]
            for f in (info['fam16'], info['fam1']):
                if f and info['sub']:
                    candidatos.append(f + ' ' + info['sub'])
            for k in candidatos:
                if k and norm(k) in doc_map:
                    target = doc_map[norm(k)]
                    merged = True
                    break
        if target is None:
            continue
        if merged:
            slots = ['embedRegular', 'embedBold']  # a face É a família do doc
        else:
            slots = ['embed' + ('BoldItalic' if info['bold'] and info['italic']
                                else 'Bold' if info['bold']
                                else 'Italic' if info['italic']
                                else 'Regular')]
        slots = [s for s in slots if (target, s) not in filled]
        if not slots:
            continue
        blk = re.search(r'<w:font w:name="%s"[^>]*>.*?</w:font>'
                        % re.escape(target), ftable, re.S)
        if not blk:
            continue
        bloco = blk.group(0)
        raw0 = open(caminho, 'rb').read()
        for slot in slots:
            bloco = re.sub(r'<w:%s [^/]*/>' % slot, '', bloco)
            guid = str(uuid.uuid4()).upper()
            if merged:
                # a face precisa se registrar com a família exata do doc
                sub_slot = 'Bold' if slot == 'embedBold' else 'Regular'
                base = _retag_font(raw0, target, sub_slot,
                                   bold=(slot == 'embedBold'), italic=False)
            else:
                base = raw0
            raw = bytearray(base)
            kb = bytes.fromhex(guid.replace('-', ''))[::-1]
            for i in range(min(32, len(raw))):
                raw[i] ^= kb[i % 16]
            seq += 1
            fname = 'fontAG%d.odttf' % seq
            rid = 'rId%d' % next_rid
            next_rid += 1
            bloco = bloco.replace(
                '</w:font>',
                '<w:%s r:id="%s" w:fontKey="{%s}"/></w:font>' % (slot, rid, guid))
            rels = rels.replace(
                '</Relationships>',
                '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/font" Target="fonts/%s"/>'
                '</Relationships>' % (rid, fname))
            extra['word/fonts/' + fname] = bytes(raw)
            filled.add((target, slot))
            added += 1
        ftable = ftable.replace(blk.group(0), bloco)
    if not added:
        return 0
    if 'Extension="odttf"' not in ctypes:
        ctypes = ctypes.replace(
            '</Types>',
            '<Default Extension="odttf" ContentType="application/'
            'vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>')
        extra['[Content_Types].xml'] = ctypes.encode('utf-8')
    if '<w:embedTrueTypeFonts' not in settings:
        settings = re.sub(r'(<w:settings[^>]*>)',
                          r'\1<w:embedTrueTypeFonts/>', settings, count=1)
        extra['word/settings.xml'] = settings.encode('utf-8')
    extra['word/fontTable.xml'] = ftable.encode('utf-8')
    extra['word/_rels/fontTable.xml.rels'] = rels.encode('utf-8')
    _replace_many_in_zip(docx_path, extra)
    return added


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
        # Busca em ORDEM de documento com fronteira monotônica: cada seção
        # está numa página >= a da anterior. Match primário: o título ocupa
        # uma linha inteira (é assim que banners e H1 saem no PDF); isso evita
        # que "GABARITO" case com as linhas "GABARITO: X" das comentadas.
        pages = {}
        floor = toc_page + 1
        total = len(pages_txt)

        def _lines(ptxt):
            return [re.sub(r'\s+', ' ', ln).strip() for ln in ptxt.splitlines()]

        for name, title, level in builder.bm.items:
            norm = re.sub(r'\s+', ' ', title).strip()
            found = None
            for pi in range(floor, total + 1):
                if norm in _lines(pages_txt[pi - 1]):
                    found = pi
                    break
            if found is None:  # fallback: substring, ainda respeitando a ordem
                for pi in range(floor, total + 1):
                    if norm in re.sub(r'\s+', ' ', pages_txt[pi - 1]):
                        found = pi
                        break
            if found is not None:
                pages[name] = str(found)
                floor = found
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
    try:
        embed_brand_fonts(out_path)
    except Exception:
        pass
    return out_path
