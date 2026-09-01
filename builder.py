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


# ---------------------------------------------------------------------------
# Fórmulas: LaTeX -> OMML (equação nativa do Word)
# Detecta $...$, $$...$$, \(...\), \[...\] e também LaTeX "cru" (sem
# delimitadores) no meio do texto, e injeta <m:oMath> no lugar do run.
# Se as bibliotecas não estiverem instaladas ou a conversão falhar, o texto
# sai literal, como antes (nunca quebra a renderização).
#   pip install latex2mathml mathml2omml
# ---------------------------------------------------------------------------
try:
    import latex2mathml.converter as _l2m
    import mathml2omml as _m2o
    _MATH_OK = True
except Exception:
    _MATH_OK = False

_M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
_MATH_CACHE = {}

_MATH_DELIM_RE = re.compile(
    r'(\$\$.+?\$\$'
    r'|(?<!\$)\$[^$\n]+?\$'
    r'|\\\(.+?\\\)'
    r'|\\\[.+?\\\])', re.S)

_LATEX_CMD_RE = re.compile(
    r'\\(?:dfrac|tfrac|frac|sqrt|bar|hat|vec|overline|underline|sum|prod|'
    r'int|cdot|times|div|pm|mp|leq|le|geq|ge|neq|ne|approx|equiv|infty|'
    r'alpha|beta|gamma|delta|epsilon|theta|lambda|mu|sigma|pi|phi|omega|'
    r'Delta|Sigma|Omega|Pi|log|ln|sin|cos|tan|binom|text|left|right)\b')

_MATH_CHARS = set('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
                  '0123456789 _^{}\\+-*/=()|<>!%'
                  '\u00b7\u03a3\u03a0\u0394\u03a9\u03c3\u03c0\u03b8'
                  '\u03bb\u00b5\u03bc\u03b1\u03b2\u03b3\u03b4\u03c6'
                  '\u03c9\u221e\u2264\u2265\u2260\u2248\u221a')


def _ilha_latex(text, cmd_start):
    """Expande, a partir de um comando LaTeX cru, a "ilha" de fórmula ao
    redor (ex.: 'x = \\frac{...}{n}' dentro de uma frase), sem engolir
    palavras do texto corrido."""
    n = len(text)

    def _pontuacao_ok(i):
        # '.' e ',' só entram na fórmula entre dígitos (decimais: 0,5 / 3.14)
        return 0 < i < n - 1 and text[i - 1].isdigit() and text[i + 1].isdigit()

    # ---- direita ----
    j = cmd_start
    depth = 0
    while j < n:
        c = text[j]
        if c == '{':
            depth += 1
            j += 1
            continue
        if c == '}':
            if depth == 0:
                break
            depth -= 1
            j += 1
            continue
        if depth > 0:
            j += 1
            continue
        if c in '.,':
            if not _pontuacao_ok(j):
                break
            j += 1
            continue
        if c.isalpha() and c.isascii():
            k = j
            while k < n and text[k].isalpha() and text[k].isascii():
                k += 1
            prev = text[j - 1] if j > 0 else ''
            if (k - j) >= 3 and prev != '\\':
                break  # palavra do texto corrido: fórmula termina antes dela
            if (k - j) < 3 and prev != '\\':
                # letra(s) solta(s): se logo depois vier palavra de prosa
                # ("... {n} e pronto"), é conjunção, não variável
                k2 = k
                while k2 < n and text[k2] == ' ':
                    k2 += 1
                if k2 < n and text[k2].isalpha() and text[k2].isascii():
                    break
            j = k
            continue
        if c not in _MATH_CHARS:
            break
        j += 1
    fim = j

    # ---- esquerda ----
    i = cmd_start
    while i > 0:
        c = text[i - 1]
        if c in '{}':
            break
        if c in '.,':
            if not _pontuacao_ok(i - 1):
                break
            i -= 1
            continue
        if c.isalpha() and c.isascii():
            k = i - 1
            while k > 0 and text[k - 1].isalpha() and text[k - 1].isascii():
                k -= 1
            if (i - k) >= 3:
                break
            i = k
            continue
        if c not in _MATH_CHARS:
            break
        i -= 1
    return i, fim


def _apara_ilha(ilha):
    """Limpa as pontas da ilha e garante chaves equilibradas."""
    ilha = ilha.strip()
    while ilha and ilha[-1] in '+-*/=(, ':
        ilha = ilha[:-1]
    while ilha and ilha[0] in '*/=+), ':
        ilha = ilha[1:]
    while ilha:
        d, estourou = 0, False
        for ch in ilha:
            if ch == '{':
                d += 1
            elif ch == '}':
                d -= 1
                if d < 0:
                    estourou = True
                    break
        if d == 0 and not estourou:
            break
        ilha = ilha[:-1].rstrip()
    return ilha.strip()


def _segmentos_cru(text):
    """Encontra LaTeX sem delimitadores dentro de texto corrido."""
    segs = []
    pos = 0     # início do texto ainda não emitido
    busca = 0   # de onde procurar o próximo comando
    while True:
        m = _LATEX_CMD_RE.search(text, busca)
        if not m:
            break
        ini, fim = _ilha_latex(text, m.start())
        trecho = text[ini:fim]
        limpo = _apara_ilha(trecho)
        if (len(limpo) >= 4 and len(limpo) <= 300 and ini >= pos
                and _LATEX_CMD_RE.search(limpo)):
            a = ini + trecho.find(limpo)
            b = a + len(limpo)
            if a > pos:
                segs.append(('txt', text[pos:a]))
            segs.append(('math', limpo))
            pos = b
            busca = max(b, m.end())
        else:
            busca = m.end()
    if pos < len(text):
        segs.append(('txt', text[pos:]))
    return segs


def _segmentos_math(text):
    """Divide o texto em segmentos [('txt'|'math', trecho)]."""
    segs = []
    pos = 0
    for m in _MATH_DELIM_RE.finditer(text):
        if m.start() > pos:
            segs.extend(_segmentos_cru(text[pos:m.start()]))
        raw = m.group(0)
        latex = raw[2:-2] if raw.startswith(('$$', '\\(', '\\[')) else raw[1:-1]
        segs.append(('math', latex.strip()))
        pos = m.end()
    if pos < len(text):
        segs.extend(_segmentos_cru(text[pos:]))
    return segs


_GROUPCHR_FIX_RE = re.compile(
    r'(<m:groupChrPr>(?:<m:\w+(?: [^>]*)?/>)*)</m:groupChr>')


def _conserta_omml(omml):
    """Corrige bug conhecido do mathml2omml 0.0.2: <m:groupChrPr> fechado
    como </m:groupChr> (acentos tipo \\bar/\\hat)."""
    return _GROUPCHR_FIX_RE.sub(r'\1</m:groupChrPr>', omml)


# ---------------------------------------------------------------------------
# v7.3: acentos matemáticos (\bar, \hat, \vec, \tilde, \dot, \overline...).
# latex2mathml gera <mover>; o mathml2omml 0.0.2 converte <mover> em
# <m:groupChr> (chave de grupo, tipo \overbrace) ou <m:limUpp> (limite
# superior, tipo o "n" em cima do somatório). Word e LibreOffice empilham
# esses construtos com o vão de uma chave/limite: é o "tracinho da média
# flutuando acima do x" reportado pelo cliente. O construto correto para
# acento é <m:acc> com o caractere COMBINANTE (U+0305 barra, U+0302 chapéu,
# U+20D7 vetor...) e, para \overline/\underline de expressões, <m:bar>.
# Tudo que não é acento (\overbrace, \lim_{x\to 0}, somatórios) fica intacto.
# ---------------------------------------------------------------------------
_M_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
_MQ = '{%s}' % _M_URI

# caractere "espaçador" emitido pelo latex2mathml -> combinante da galeria
# de acentos do Word (Cambria Math renderiza colado à base, no lugar certo)
_ACENTO_COMBINANTE = {
    '\u00AF': '\u0305',   # ¯  \bar               -> COMBINING OVERLINE
    '\u0304': '\u0305',   # macron combinante (se vier)
    '\u005E': '\u0302',   # ^  \hat \widehat
    '\u02C6': '\u0302',   # ˆ
    '\u007E': '\u0303',   # ~  \tilde \widetilde
    '\u02DC': '\u0303',   # ˜
    '\u02D9': '\u0307',   # ˙  \dot
    '\u00A8': '\u0308',   # ¨  \ddot
    '\u20DB': '\u20DB',   # ⃛  \dddot (já combinante)
    '\u2192': '\u20D7',   # →  \vec \overrightarrow
    '\u2190': '\u20D6',   # ←  \overleftarrow
    '\u2194': '\u20E1',   # ↔  \overleftrightarrow
    '\u02D8': '\u0306',   # ˘  \breve
    '\u02C7': '\u030C',   # ˇ  \check
    '\u00B4': '\u0301',   # ´  \acute
    '\u0060': '\u0300',   # `  \grave
    '\u02DA': '\u030A',   # ˚  \mathring
}
# barra "longa" que o latex2mathml usa para \overline / \underline
_BARRA_LONGA = {'\u2015', '\u203E', '\u0332', '\u005F'}


def _omml_lim_char(el):
    """Se <el> contém exatamente um <m:r> com um único caractere, devolve-o."""
    runs = el.findall('.//' + _MQ + 'r')
    if len(runs) != 1:
        return None
    ts = runs[0].findall(_MQ + 't')
    if len(ts) != 1:
        return None
    txt = (ts[0].text or '').strip()
    return txt if len(txt) == 1 else None


def _omml_novo_acc(etree, chr_comb, base_e):
    acc = etree.Element(_MQ + 'acc')
    pr = etree.SubElement(acc, _MQ + 'accPr')
    c = etree.SubElement(pr, _MQ + 'chr')
    c.set(_MQ + 'val', chr_comb)
    acc.append(base_e)
    return acc


def _omml_nova_bar(etree, pos, base_e):
    bar = etree.Element(_MQ + 'bar')
    pr = etree.SubElement(bar, _MQ + 'barPr')
    p = etree.SubElement(pr, _MQ + 'pos')
    p.set(_MQ + 'val', pos)
    bar.append(base_e)
    return bar


def _normaliza_acentos_omml(omml):
    """Reescreve groupChr[top]/limUpp/limLow de acento em m:acc / m:bar.
    Recebe o <m:oMath ...> já com xmlns:m. Em qualquer erro devolve a
    entrada intacta (nunca quebra a renderização)."""
    try:
        from lxml import etree
        root = etree.fromstring(omml.encode('utf-8'))
    except Exception:
        return omml
    mudou = False

    def _troca(velho, novo):
        velho.getparent().replace(velho, novo)

    # 1) groupChr com pos=top e chr de acento -> acc (ou bar p/ barra longa)
    for g in list(root.iter(_MQ + 'groupChr')):
        pr = g.find(_MQ + 'groupChrPr')
        e = g.find(_MQ + 'e')
        if pr is None or e is None:
            continue
        chr_el = pr.find(_MQ + 'chr')
        pos_el = pr.find(_MQ + 'pos')
        chr_val = chr_el.get(_MQ + 'val') if chr_el is not None else None
        pos_val = pos_el.get(_MQ + 'val') if pos_el is not None else 'bot'
        if pos_val != 'top' or chr_val is None:
            continue
        if chr_val in _ACENTO_COMBINANTE:
            _troca(g, _omml_novo_acc(etree, _ACENTO_COMBINANTE[chr_val], e))
            mudou = True
        elif chr_val in _BARRA_LONGA:
            _troca(g, _omml_nova_bar(etree, 'top', e))
            mudou = True

    # 2) limUpp cujo "lim" é um único caractere de acento -> acc / bar top
    for lu in list(root.iter(_MQ + 'limUpp')):
        e = lu.find(_MQ + 'e')
        lim = lu.find(_MQ + 'lim')
        if e is None or lim is None:
            continue
        ch = _omml_lim_char(lim)
        if ch is None:
            continue
        if ch in _ACENTO_COMBINANTE:
            _troca(lu, _omml_novo_acc(etree, _ACENTO_COMBINANTE[ch], e))
            mudou = True
        elif ch in _BARRA_LONGA:
            _troca(lu, _omml_nova_bar(etree, 'top', e))
            mudou = True

    # 3) limLow cujo "lim" é a barra longa -> bar bot (\underline)
    for ll in list(root.iter(_MQ + 'limLow')):
        e = ll.find(_MQ + 'e')
        lim = ll.find(_MQ + 'lim')
        if e is None or lim is None:
            continue
        if _omml_lim_char(lim) in _BARRA_LONGA:
            _troca(ll, _omml_nova_bar(etree, 'bot', e))
            mudou = True

    if not mudou:
        return omml
    return etree.tostring(root, encoding='unicode')


def _omml_valido(omml):
    """Valida o XML antes de injetar no documento; inválido -> descarta."""
    try:
        from lxml import etree
        etree.fromstring(omml.encode('utf-8'))
        return True
    except ImportError:
        return True     # sem lxml não dá para validar: segue o melhor esforço
    except Exception:
        return False


def _latex_to_omml(latex):
    """LaTeX -> <m:oMath> (equação nativa). None se não der para converter."""
    if not _MATH_OK or not latex:
        return None
    if latex in _MATH_CACHE:
        return _MATH_CACHE[latex]
    omml = None
    try:
        mathml = _l2m.convert(latex)
        omml = _m2o.convert(mathml)
        if omml.startswith('<m:oMath>'):
            omml = _conserta_omml(omml)
            omml = omml.replace('<m:oMath>', '<m:oMath %s>' % _M_NS, 1)
            if not _omml_valido(omml):
                omml = None
            else:
                omml = _normaliza_acentos_omml(omml)   # v7.3: \bar, \hat...
        else:
            omml = None
    except Exception:
        omml = None
    _MATH_CACHE[latex] = omml
    return omml


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


_TOKEN_RE = re.compile(r'(\*\*.+?\*\*|__.+?__|%%.+?%%)', re.S)


def _with_props(rpr, bold=False, underline=False, color=None):
    """Deriva um rPr acrescentando negrito/sublinhado/cor ao existente."""
    inner = re.sub(r'^<w:rPr>|</w:rPr>$', '', rpr) if rpr else ''
    if color is not None:
        inner = re.sub(r'<w:color [^/]*/>', '', inner)
        inner += '<w:color w:val="%s"/>' % color
    if bold and '<w:b' not in inner:
        inner = '<w:b/>' + inner
    if underline and '<w:u ' not in inner:
        inner += '<w:u w:val="single"/>'
    return '<w:rPr>%s</w:rPr>' % inner if inner else ''


def _runs_texto(text, base_rpr, bold_rpr, highlight_fgv):
    """Marcas inline -> runs (texto já saneado)."""
    out = []
    for part in _TOKEN_RE.split(text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            txt, rpr = part[2:-2], bold_rpr
        elif part.startswith('__') and part.endswith('__') and len(part) > 4:
            txt, rpr = part[2:-2], _with_props(bold_rpr, bold=True,
                                               underline=True)
        elif part.startswith('%%') and part.endswith('%%') and len(part) > 4:
            txt, rpr = part[2:-2], _with_props(base_rpr, bold=True,
                                               color='C00000')
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


def make_runs(text, base_rpr, bold_rpr=None, highlight_fgv=True):
    """Converte texto com marcas inline em runs:
    **negrito**  __negrito sublinhado__  %%negrito vermelho%%
    FGV recebe highlight amarelo. Trechos em LaTeX ($...$, \\(...\\) ou
    LaTeX cru com \\frac, \\bar etc.) viram equações nativas do Word
    (m:oMath); se a conversão falhar, saem como texto literal."""
    if bold_rpr is None:
        bold_rpr = _add_bold(base_rpr)
    out = []
    for kind, seg in _segmentos_math(_sanitize(text)):
        if kind == 'math':
            omml = _latex_to_omml(seg)
            if omml:
                out.append(omml)
                continue
        out.append(_runs_texto(seg, base_rpr, bold_rpr, highlight_fgv))
    return ''.join(out)


def retext_para(para_xml, text, keep_first_n_runs=0, highlight_fgv=True):
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
    # v7.2.4: capa com título longo. A faixa do template comporta 2 linhas a
    # 22pt; títulos que precisariam de 3+ linhas têm a fonte reduzida até
    # caberem em 2 linhas (pedido do cliente). Se nem no piso couber,
    # fallback: estica a faixa mantendo o centro.
    _COVER_CHARS_POR_LINHA = 26   # calibrado na largura útil da caixa a 22pt
    _COVER_SZ_BASE = 44           # 22pt (estilo TtuloCapa)
    _COVER_SZ_MIN = 28            # 14pt: piso da redução
    _COVER_EMU_POR_LINHA = 350000  # ~0,97 cm por linha extra (fallback)

    @classmethod
    def _cover_linhas(cls, titulo, limite=None):
        limite = limite or cls._COVER_CHARS_POR_LINHA
        linhas, atual = 1, ''
        for w in str(titulo).split():
            t = (atual + ' ' + w).strip()
            if len(t) <= limite or not atual:
                atual = t
            else:
                linhas += 1
                atual = w
        return linhas

    @classmethod
    def _cover_fit_sz(cls, titulo):
        """Maior tamanho de fonte (half-points) que faz o título caber em
        2 linhas; None quando o tamanho do estilo (22pt) já basta."""
        if cls._cover_linhas(titulo) <= 2:
            return None
        sz = cls._COVER_SZ_BASE
        while sz > cls._COVER_SZ_MIN:
            sz -= 2
            limite = int(cls._COVER_CHARS_POR_LINHA * cls._COVER_SZ_BASE / sz)
            if cls._cover_linhas(titulo, limite) <= 2:
                return sz
        return cls._COVER_SZ_MIN

    @staticmethod
    def _cover_set_sz(box_inner, sz):
        """Grava o tamanho de fonte em todos os runs da caixa de título."""
        tag = '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)

        def rep(m):
            run = m.group(0)
            if '<w:rPr>' in run:
                run = re.sub(r'<w:sz w:val="\d+"/>|<w:szCs w:val="\d+"/>', '', run)
                return run.replace('<w:rPr>', '<w:rPr>' + tag, 1)
            return re.sub(r'^(<w:r(?: [^>]*)?>)', r'\1<w:rPr>%s</w:rPr>' % tag,
                          run, count=1)
        return re.sub(r'<w:r(?: [^>]*)?>.*?</w:r>', rep, box_inner, flags=re.S)

    @staticmethod
    def _cover_grow(art, extra):
        """Aumenta em `extra` EMU a altura de cada shape ancorada da arte do
        título e sobe o offset vertical em extra/2 (centro preservado)."""
        def patch(m):
            bloco = m.group(0)
            me = re.search(r'<wp:extent cx="\d+" cy="(\d+)"/>', bloco)
            if not me:
                return bloco
            cy0 = me.group(1)
            bloco = bloco.replace('cy="%s"' % cy0, 'cy="%d"' % (int(cy0) + extra))
            bloco = re.sub(
                r'(<wp:positionV[^>]*><wp:posOffset>)(\d+)(</w?p?:?posOffset>)',
                lambda mm: mm.group(1) + str(max(0, int(mm.group(2)) - extra // 2))
                + mm.group(3),
                bloco)
            return bloco
        return re.sub(r'<wp:anchor\b.*?</wp:anchor>', patch, art, flags=re.S)

    def cover(self, titulo):
        self.add(self.f.cover_bg)
        sz = self._cover_fit_sz(titulo)

        def fix_box(m):
            inner = m.group(0)
            done = [False]

            def rep(mm):
                if not done[0]:
                    done[0] = True
                    return ('<w:t xml:space="preserve">%s</w:t>'
                            % escape(titulo))
                return '<w:t xml:space="preserve"></w:t>'
            inner = T_RE.sub(rep, inner)
            if sz:
                inner = self._cover_set_sz(inner, sz)
            return inner

        art = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', fix_box,
                     self.f.cover_art, flags=re.S)
        if sz:
            limite = int(self._COVER_CHARS_POR_LINHA * self._COVER_SZ_BASE / sz)
            linhas = self._cover_linhas(titulo, limite)
            if linhas > 2:
                art = self._cover_grow(
                    art, (linhas - 2) * self._COVER_EMU_POR_LINHA)
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
        self.add(_with_bookmark(self.f.banner_t2, titulo, name, self.bm_id))
        self.blank()

    # v7.2.6: subtítulo no tamanho do corpo (12pt). O estilo Ttulo4 do
    # template é 13pt em Winner Sans Wide (fonte larga), que aparenta ser
    # bem maior que o texto; o tamanho é forçado nos runs.
    _SUBTITLE_SZ = 24

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
        # v7.2.8: a moldura é desenhada `space` pontos para FORA do texto,
        # mais a espessura da linha; sem recuo ela vazava além das margens
        # da coluna. O recuo compensa exatamente esse deslocamento, deixando
        # a borda externa alinhada com o texto do corpo.
        recuo_tw = int(round((8 + int(sz) / 8.0) * 20))  # space(8pt)+linha, em twips
        ind = '<w:ind w:left="%d" w:right="%d"/>' % (recuo_tw, recuo_tw)
        extras = '<w:keepNext/>' + espac + ind + borda
        if '<w:pPr>' in p:
            if '<w:pStyle' in p:
                p = re.sub(r'(<w:pStyle [^/]*/>)', r'\1' + extras, p, count=1)
            else:
                p = p.replace('<w:pPr>', '<w:pPr>' + extras, 1)
        else:
            p = re.sub(r'<w:p((?: [^>]*)?)>',
                       r'<w:p\1><w:pPr>' + extras + '</w:pPr>', p, count=1)
        # força o tamanho de fonte do corpo em todos os runs do subtítulo
        tag = ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
               % (self._SUBTITLE_SZ, self._SUBTITLE_SZ))

        def _sz(mm):
            run = mm.group(0)
            run = re.sub(r'<w:sz w:val="\d+"/>|<w:szCs w:val="\d+"/>', '', run)
            if '<w:rPr>' in run:
                return run.replace('<w:rPr>', '<w:rPr>' + tag, 1)
            return re.sub(r'^(<w:r(?: [^>]*)?>)', r'\1<w:rPr>%s</w:rPr>' % tag,
                          run, count=1)
        p = re.sub(r'<w:r(?: [^>]*)?>.*?</w:r>', _sz, p, flags=re.S)
        self.add(p)

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

    # v7.2: linha que é só uma "opção" de Certo/Errado (ex.: "C) Certo",
    # "E) Errado", "( ) Certo", "Certo."). Questão de Certo/Errado nunca
    # lista opções; elas são removidas quando o corpo traz o par completo.
    _CE_LINHA = re.compile(
        r'^\s*\(?\s*[A-Ea-e]?\s*[).:\-]?\s*\(?\s*(certo|errado)\s*\)?\s*\.?\s*$',
        re.I)

    @classmethod
    def _limpa_certo_errado(cls, corpo):
        """Se o corpo trouxer linhas de opção 'Certo' E 'Errado', remove-as
        (mantendo texto de apoio e afirmativa). Devolve o corpo filtrado."""
        linhas = [str(x if x is not None else '') for x in (corpo or [])]
        marca = [bool(cls._CE_LINHA.match(l.strip())) for l in linhas]
        tem_c = any(m and re.search(r'certo', l, re.I)
                    for l, m in zip(linhas, marca))
        tem_e = any(m and re.search(r'errado', l, re.I)
                    for l, m in zip(linhas, marca))
        if not (tem_c and tem_e):
            return linhas
        return [l for l, m in zip(linhas, marca) if not m]

    def _linha_questao(self, modelo, linha):
        """Renderiza uma linha de questão; marcadores [IMAGEM N] viram a
        imagem real no ponto correspondente (questões com texto de apoio
        visual, charge, campanha etc.). Sem vermelho nas seções de questões:
        %%alerta%% rebaixa para negrito preto."""
        s = str(linha).replace('%%', '**')
        marcas = self._IMG_MARK.findall(s)
        if not marcas:
            self.add(retext_para(modelo, s, highlight_fgv=False))
            return
        resto = self._IMG_MARK.sub('', s).strip()
        if resto:
            self.add(retext_para(modelo, resto, highlight_fgv=False))
        for ref in marcas:
            self.imagem(ref)

    # v7.2.6 (padrão CEBRASPE, imagem de referência do cliente):
    # - cabecalho COM comando ("... julgue o item...") -> caixa cobre número +
    #   banca + comando; enunciado/afirmativa descem no corpo, sem caixa.
    # - cabecalho SEM comando -> caixa só em "NN. (BANCA / Órgão / Ano)" e
    #   qualquer texto excedente desce para o corpo.
    # - Se a IA mandar a questão inteira no cabecalho (corpo vazio), a 1ª
    #   frase fica na caixa quando for comando (julgue/assinale); o resto cai
    #   no corpo.
    # frases de comando que sobem para a caixa destacada do cabeçalho:
    # "julgue o item...", "assinale a alternativa...", "marque a opção...",
    # "indique...", "aponte...", "analise...", "é correto afirmar",
    # "está correto o que se afirma"
    _COMANDO_RE = re.compile(
        r'\bjulgue\b|\bassinale\b|\bmarque\b|\bindique\b|\baponte\b|'
        r'\banalise\b|\bé\s+correto\b|\best[áa]\s+corret[oa]\b', re.I)
    _ALT_RE = re.compile(r'^\s*[A-Ea-e]\s*[).]\s+')
    # linha de ITEM de lista dentro da questão (afirmativas V/F, itens
    # romanos/numerados, bullets): nunca é enunciado, nunca ganha caixa.
    # Ex.: "( ) Todo cidadão...", "(V) ...", "I. ...", "II) ...", "1. ..."
    _ITEM_RE = re.compile(
        r'^\s*(\(\s*[VvFf]?\s*\)|[IVXLCDM]+\s*[.)\-\u2013\u2014]\s|'
        r'\d+\s*[.)\-\u2013\u2014]\s|[\u2022\u25aa\u2023*\-\u2013\u2014]\s)')
    # letra da alternativa em minúsculo, como na máscara: "A) ..." -> "a) ..."
    _ALT_MIN_RE = re.compile(r'^(\s*)([A-E])(\s*[).])')

    @staticmethod
    def _separa_cabecalho(cab):
        cab = str(cab or '')
        if not re.match(r'^\s*\d+\s*[.)\-]?\s*\(', cab):
            return cab, ''
        i = cab.find('(')
        depth = 0
        for k in range(i, len(cab)):
            if cab[k] == '(':
                depth += 1
            elif cab[k] == ')':
                depth -= 1
                if depth == 0:
                    return cab[:k + 1].rstrip(), cab[k + 1:].strip()
        return cab, ''

    @classmethod
    def _monta_sequencia(cls, cabecalho, corpo):
        """Devolve [(tipo, linha)] com tipo 'caixa' (sombreado) ou 'corpo'.

        Regra v7.2.9 (definida pelo cliente):
        - Certo/Errado: caixa = numero + banca + comando; afirmativa fora.
        - Multipla escolha: caixa = 1a linha (numero + banca + comando) e o
          ENUNCIADO (ultima linha antes das alternativas); texto de apoio e
          fonte (linha entre parenteses) ficam FORA, entre as duas caixas.
          Alternativas sempre fora."""
        corpo = [str(l if l is not None else '') for l in (corpo or [])]
        # padrão da máscara: letra das alternativas em minúsculo (a, b, c...)
        corpo = [cls._ALT_MIN_RE.sub(
            lambda m: m.group(1) + m.group(2).lower() + m.group(3), l)
            for l in corpo]
        cab, resto = cls._separa_cabecalho(cabecalho)
        linha_cab = (cab + (' ' + resto if resto else '')).strip()
        tem_alt = any(cls._ALT_RE.match(l) for l in corpo)
        if tem_alt:
            pre, alts = [], []
            achou_alt = False
            for l in corpo:
                if not achou_alt and cls._ALT_RE.match(l):
                    achou_alt = True
                (alts if achou_alt else pre).append(l)
            # comando na 1ª linha ("Assinale...") sobe para a caixa do
            # cabeçalho, como no ramo Certo/Errado
            while pre and not pre[0].strip():
                pre.pop(0)
            if pre:
                prim = pre[0].strip()
                if (cls._COMANDO_RE.search(prim) and len(prim) <= 240
                        and '[IMAGEM' not in prim):
                    linha_cab = (linha_cab + ' ' + prim).strip()
                    pre = pre[1:]
            # enunciado = ultima linha pre-alternativas, exceto fonte "(...)",
            # linha com imagem ou ITEM de lista (afirmativa V/F, item romano
            # ou numerado, bullet): item nunca ganha caixa
            enunciado = None
            while pre:
                ult = pre[-1].strip()
                if not ult:
                    pre.pop()
                    continue
                eh_fonte = bool(re.match(r'^\(.*\)[.\s]*$', ult))
                eh_item = bool(cls._ITEM_RE.match(ult))
                if not eh_fonte and not eh_item and '[IMAGEM' not in ult:
                    enunciado = pre.pop().strip()
                break
            seq = [('caixa', linha_cab)]
            seq += [('corpo', l) for l in pre if str(l).strip()]
            if enunciado:
                seq.append(('caixa', enunciado))
            seq += [('corpo', l) for l in alts]
            return seq
        # ---- Certo/Errado / sem alternativas ----
        # comando veio como 1ª linha do corpo -> sobe para a caixa, junto ao
        # cabeçalho (mesmo parágrafo, como na máscara)
        if corpo and len([l for l in corpo if l.strip()]) >= 2 and not resto:
            prim = corpo[0].strip()
            if (prim and cls._COMANDO_RE.search(prim) and len(prim) <= 240
                    and '[IMAGEM' not in prim):
                linha_cab = (linha_cab + ' ' + prim).strip()
                corpo = corpo[1:]
        if not corpo and resto:
            # tudo veio no cabecalho: 1a frase fica na caixa se for comando
            m = re.match(r'([^.?!:]*[.?!:])\s*(.*)', resto, re.S)
            if m and cls._COMANDO_RE.search(m.group(1)):
                seq = [('caixa', (cab + ' ' + m.group(1)).strip())]
                r2 = m.group(2).strip()
                if r2:
                    seq.append(('corpo', r2))
                return seq
            return [('caixa', cab), ('corpo', resto)]
        return [('caixa', linha_cab)] + [('corpo', l) for l in corpo]

    def _render_sequencia(self, seq):
        for tipo, linha in seq:
            if tipo == 'caixa':
                self.add(retext_para(self.f.q_cab, linha, highlight_fgv=False))
            else:
                self._render_linha_corpo(linha)

    def _render_linha_corpo(self, linha):
        self._linha_questao(self.f.q_corpo, linha)

    def questao(self, cabecalho, corpo, certo_errado=False):
        corpo = self._limpa_certo_errado(corpo)
        self._render_sequencia(self._monta_sequencia(cabecalho, corpo))
        # padrão v4 (PDF de referência): sem linhas "Certo (   )/Errado (   )"
        self.add(self.f.q_corpo_blank)
        self.add(self.f.q_espaco)

    @staticmethod
    def _force_bold(rpr):
        """Garante <w:b/> e cor preta no rPr (rótulos GABARITO:/COMENTÁRIO:
        sempre em negrito preto, independente do run do template)."""
        if not rpr:
            return '<w:rPr><w:b/><w:color w:val="000000"/></w:rPr>'
        rpr = re.sub(r'<w:color [^/>]*/>', '', rpr)
        rpr = rpr.replace('</w:rPr>', '<w:color w:val="000000"/></w:rPr>', 1)
        if not re.search(r'<w:b\b(?![A-Za-z])', rpr):
            rpr = rpr.replace('<w:rPr>', '<w:rPr><w:b/>', 1)
        return rpr

    def _lead_para(self, model, lead_text, body_text):
        """Parágrafo com rótulo em negrito (rPr clonado do primeiro run do
        modelo, ex.: "GABARITO: ") seguido do corpo em formatação normal,
        com suporte a **negrito** e highlight FGV."""
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', model, re.S)
        ppr = ppr.group(0) if ppr else ''
        runs = _runs(model)
        lead_rpr = self._force_bold(_rpr(runs[0]) if runs else '')
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

    _SHD_RE = re.compile(r'<w:shd [^>]*?w:fill="([0-9A-Fa-f]{6})"[^>]*/>')

    def _caixa_comentario(self, paras_xml):
        """Envolve os parágrafos do comentário numa tabela de célula única,
        sem bordas, com o sombreamento na CÉLULA (pedido do revisor): assim o
        fundo estica junto com o texto quando alguém edita o comentário no
        Word, sem quebrar o alinhamento. O tom do fundo é herdado do
        sombreamento que o template usava nos parágrafos."""
        m = self._SHD_RE.search(paras_xml)
        fill = m.group(1) if m else None
        if fill:
            # o fundo agora é da célula: tira o dos parágrafos p/ não duplicar
            paras_xml = self._SHD_RE.sub('', paras_xml)
        nenhuma = ('<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        tcpr = '<w:tcW w:w="5000" w:type="pct"/>'
        if fill:
            tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill
        return (
            '<w:tbl><w:tblPr>'
            '<w:tblW w:w="5000" w:type="pct"/>'
            '<w:tblBorders>%s</w:tblBorders>'
            '<w:tblCellMar>'
            '<w:top w:w="113" w:type="dxa"/><w:left w:w="142" w:type="dxa"/>'
            '<w:bottom w:w="113" w:type="dxa"/><w:right w:w="142" w:type="dxa"/>'
            '</w:tblCellMar>'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="10469"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr>%s</w:tcPr>%s</w:tc></w:tr>'
            '</w:tbl>' % (nenhuma, tcpr, paras_xml))

    def comentario(self, cabecalho, corpo, gabarito, comentario):
        corpo = self._limpa_certo_errado(corpo)
        self._render_sequencia(self._monta_sequencia(cabecalho, corpo))
        # v7.2: GABARITO + resposta em negrito e CAIXA ALTA, no rótulo
        gab_txt = re.sub(r'[*_%]', '', str(gabarito)).strip().upper()
        p_gab = self._lead_para(self.f.c_gab, 'GABARITO: ' + gab_txt, '')
        # texto do comentário sai limpo: remove **negrito**, __sublinhado__ e
        # %%vermelho%% — só os rótulos GABARITO:/COMENTÁRIO: ficam em negrito
        comentario = re.sub(r'\*\*|__|%%', '', str(comentario))
        p_com = self._lead_para(self.f.c_com, 'COMENTÁRIO: ', comentario)
        # v10: bloco inteiro numa tabela de 1 célula sem bordas (fundo na
        # célula acompanha edições sem desalinhar)
        self.add(self._caixa_comentario(p_gab + p_com))
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

_ANO_SEG_RE = re.compile(r'\b(19\d{2}|20\d{2})\b')


def _normaliza_cabecalho_banca(cab):
    """Força o cabeçalho no padrão 'NN. (BANCA / Órgão / Ano)'.

    O arquivo fonte costuma trazer campos extras dentro dos parênteses
    (cargo, órgão duplicado): 'NN. (BANCA / Órgão / Órgão / Cargo / Ano)'.
    Mantém o 1º campo (banca), o campo logo após a banca (órgão) e o último
    que contiver ano; descarta o restante. Parênteses aninhados no nome da
    banca/órgão são respeitados. Texto após o fecha-parênteses é preservado.
    """
    s = str(cab or '')
    m = re.match(r'^(\s*\d+\s*[.)\-]?\s*)\(', s)
    if not m:
        return s
    i = s.find('(')
    depth, fim = 0, -1
    for k in range(i, len(s)):
        if s[k] == '(':
            depth += 1
        elif s[k] == ')':
            depth -= 1
            if depth == 0:
                fim = k
                break
    if fim < 0:
        return s
    inner, resto = s[i + 1:fim], s[fim + 1:]
    # split de nível superior por '/'
    segs, buf, depth = [], '', 0
    for ch in inner:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch == '/' and depth == 0:
            segs.append(buf.strip())
            buf = ''
        else:
            buf += ch
    if buf.strip():
        segs.append(buf.strip())
    if len(segs) <= 3:
        return s
    banca = segs[0]
    ano = ''
    for sg in reversed(segs[1:]):
        if _ANO_SEG_RE.search(sg):
            ano = sg
            break
    meio = [sg for sg in segs[1:] if sg != ano]
    orgao = meio[0] if meio else ''
    partes = [banca] + ([orgao] if orgao else []) + ([ano] if ano else [])
    return '%s(%s)%s' % (m.group(1), ' / '.join(partes), resto)


# ordem oficial das bancas nas seções de questões (definida pelo cliente):
# CEBRASPE/CESPE, FGV, VUNESP, FCC, AOCP/Instituto AOCP, IBADE, FUNDATEC,
# IBFC, IDECAN, FUMARC, Instituto AVALIA, FEPESE; demais em ordem alfabética.
_ORDEM_BANCAS = [('CEBRASPE', 'CESPE'), ('FGV',), ('VUNESP',), ('FCC',),
                 ('AOCP',), ('IBADE',), ('FUNDATEC',), ('IBFC',),
                 ('IDECAN',), ('FUMARC',), ('AVALIA',), ('FEPESE',)]


def _ordena_questoes_por_banca(data):
    """Reordena questoes/comentarios/gabarito por banca e ano.

    Ordem das bancas: a lista oficial _ORDEM_BANCAS (CEBRASPE/CESPE, FGV,
    VUNESP, FCC, AOCP, IBADE, FUNDATEC, IBFC, IDECAN, FUMARC, AVALIA,
    FEPESE); demais em ordem alfabética depois delas. Dentro de cada
    banca, ano decrescente (mais novo primeiro);
    empates preservam a ordem original (sort estável). Depois de ordenar,
    renumera os cabeçalhos de questoes e comentarios e reconstrói o
    gabarito com a numeração nova.
    """
    questoes = list(data.get('questoes') or [])
    if not questoes:
        return data
    comentarios = list(data.get('comentarios') or [])
    gabarito = list(data.get('gabarito') or [])

    def _banca_de(cab):
        m = re.search(r'\(\s*([^/)\n]+?)\s*[/)]', str(cab or ''))
        return (m.group(1).strip().upper() if m else '')

    def _ano_de(cab):
        anos = re.findall(r'\b(19\d{2}|20\d{2})\b', str(cab or ''))
        return int(anos[-1]) if anos else 0

    def _rank_banca(banca):
        b = re.sub(r'[^A-Z0-9]+', '', banca)
        for pos, chaves in enumerate(_ORDEM_BANCAS):
            if any(c in b for c in chaves):
                return (pos, '')
        return (len(_ORDEM_BANCAS), banca)  # demais: alfabético pelo nome

    # gabarito antigo indexado pelo n (1-based) -> resposta
    g_por_n = {}
    for g in gabarito:
        try:
            g_por_n[int(g.get('n'))] = g.get('g', '')
        except (TypeError, ValueError):
            pass

    idx = list(range(len(questoes)))
    idx.sort(key=lambda i: (_rank_banca(_banca_de(questoes[i].get('cabecalho'))),
                            -_ano_de(questoes[i].get('cabecalho')),
                            i))

    def _renumera(cab, novo_n):
        s = _normaliza_cabecalho_banca(str(cab or ''))
        novo, feito = re.subn(r'^\s*\d+\s*[.)\-]?', '%d.' % novo_n, s, count=1)
        return novo if feito else ('%d. %s' % (novo_n, s.strip()))

    novas_q, novos_c, novo_g = [], [], []
    for pos, i in enumerate(idx):
        n = pos + 1
        q = dict(questoes[i])
        q['cabecalho'] = _renumera(q.get('cabecalho'), n)
        novas_q.append(q)
        if i < len(comentarios) and comentarios[i]:
            c = dict(comentarios[i])
            c['cabecalho'] = _renumera(c.get('cabecalho'), n)
            novos_c.append(c)
            resp = g_por_n.get(i + 1, c.get('gabarito', ''))
        else:
            resp = g_por_n.get(i + 1, '')
        m = re.match(r'^\s*letra\s*([A-Ea-e])\s*$', str(resp or ''))
        novo_g.append({'n': n, 'g': m.group(1).upper() if m else str(resp or '')})

    data['questoes'] = novas_q
    if comentarios:
        data['comentarios'] = novos_c
    data['gabarito'] = novo_g
    return data


def build_document(data, frag=None, prebuilt=None):
    data = _ordena_questoes_por_banca(data)
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


def _lo_math_disponivel():
    """True se o soffice local tem o componente Math (libsmlo). Sem ele, a
    conversão docx->PDF descarta as equações (páginas do sumário também podem
    desviar em materiais com muitas fórmulas)."""
    for d in ('/usr/lib/libreoffice/program', '/usr/lib64/libreoffice/program',
              '/opt/libreoffice/program', '/usr/lib/libreoffice/program/../program'):
        if os.path.isdir(d):
            try:
                return any(n.startswith('libsm') for n in os.listdir(d))
            except OSError:
                pass
    return None  # LibreOffice não localizado: sem veredito


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
    if _MATH_OK and _lo_math_disponivel() is False:
        import sys
        print('[builder] AVISO: LibreOffice local SEM libreoffice-math: '
              'PDFs gerados por este soffice descartam as equações e a '
              'paginação do sumário pode desviar. Instale libreoffice-math '
              'na imagem (e confira também o container do conversor de PDF).',
              file=sys.stderr)
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
