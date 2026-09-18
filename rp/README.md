# Renderizador RP Concursos (máscara Rani Passos)

Microserviço irmão do renderizador Águia: **mesmo contrato JSON**, máscara
diferente. Os workflows do n8n (`... - rani`) mandam exatamente o mesmo payload
do pipeline da Águia; só muda a `rendererUrl`.

Não altera nada do renderizador Águia (`builder.py` da raiz do repo). Esta
pasta traz uma cópia do `builder.py` (usada como biblioteca: marcas inline,
LaTeX→OMML, imagens, ordenação por banca, paginação do sumário) e o
`builder_rp.py`, que monta o documento na máscara RP.

## Deploy no Easypanel (panel.constroi.net.br, projeto n8n2)
1. Serviço **App > Dockerfile**, build a partir desta pasta (`rp/`).
2. Porta interna **8000**. Host interno esperado pelos workflows:
   `http://n8n_renderdocx_rp:8000`.
3. 1 vCPU / 1 GB RAM (LibreOffice roda por chamada).

## Endpoints
- `GET /health`
- `POST /extract` (multipart `file`: docx ou pptx) -> `{"text": ..., "images": [...]}`
- `POST /render` (JSON) -> docx binário. Headers `X-Toc-Paginated` e `X-Toc-Error`.

## Schema
Idêntico ao do renderizador Águia (ver README da raiz): `titulo`,
`apresentacao[]`, `capitulos[].blocos[]`, `formulas[]`, `questoes[]`,
`comentarios[]`, `gabarito[]`, `imagens{}`.

## Como cada bloco sai na máscara RP
| Bloco do JSON | Saída |
|---|---|
| `titulo` | capa "ASSUNTO:" + página do professor + sumário |
| `apresentacao` | seção INTRODUÇÃO (1ª linha "Fala, ..." em magenta) |
| capítulo | faixa roxa (estilo `Título`), entra no sumário (nível 1) |
| `subtitulo` | faixa magenta (estilo `Subtítulo`), sumário nível 2 |
| `paragrafo` | Calibri 12, justificado, entrelinha 1,5 |
| `dialogo` | caixa **INTERAÇÃO** (retângulo lilás com avatares; altura calculada) |
| `dica` / `mnemonico` / `aprofundando` / `lei` / `jurisprudencia` / `divergencia` / `revisao` | rótulo magenta em negrito + texto |
| `tabela` | tabela com cabeçalho roxo, bordas lilás, largura da página |
| `imagem` | imagem centralizada + legenda |
| `formulas` | seção RESUMO DAS FÓRMULAS APRESENTADAS (tabela com equações do Word) |
| `questoes` | LISTA DE QUESTÕES: numeração automática, banca em negrito no padrão `(CEBRASPE / TCE-AC – 2024)`, alternativas `a) b) c)` |
| `comentarios` | LISTA DE QUESTÕES COMENTADAS: questão + `GABARITO: ERRADO` + `COMENTÁRIOS:` em magenta |

Diferenças propositais em relação à máscara Águia: sem banners de
"QUESTÕES PARA PRATICAR"/"GABARITO FINAL" (o padrão RP usa as duas listas),
sem prefixo "CAPÍTULO N —" nos títulos e sem as caixas coloridas (a máscara RP
não tem esse elemento).

## Template
`template.docx` é o "Novo Padrão 2.0 - RP" com a página de apresentação do
professor embutida como `word/media/professor.png` (relationship `rIdProf`).
Para trocar a foto/arte do professor, substitua esse arquivo dentro do zip
mantendo o nome. Estilos usados: `Ttulo`, `Subttulo`, `Normal`,
`PargrafodaLista`, `GabaritoeComentrio`, `SemEspaamento` (Interação),
`Sumrio1/2`, `CabealhodoSumrio`.

## Fontes
A máscara usa Calibri; a imagem instala `fonts-crosextra-carlito` (métricas
compatíveis) para a medição do sumário e a conversão em PDF baterem com o Word.
