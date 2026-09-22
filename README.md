# renderdocxs

Microserviço em **Python (FastAPI) + LibreOffice** que transforma conteúdo estruturado em JSON em apostilas `.docx` diagramadas na identidade visual do cliente, prontas para conversão em PDF. Também extrai texto e marcações de arquivos `.docx`, `.pptx` e `.pdf` (com OCR) para alimentar pipelines de IA.

Em produção desde 2026 como etapa de renderização de um pipeline n8n que gera materiais didáticos de cursos preparatórios: o fluxo monitora uma pasta no OneDrive, extrai o material bruto do professor, estrutura teoria e questões com um LLM e chama este serviço para montar o documento final. O que antes levava dias de diagramação manual passou a sair em minutos.

## O que ele faz

- **Renderização fiel ao template**: capa, sumário real paginado, banners de capítulo, subtítulos com moldura, tabelas, caixas (dica, lei, jurisprudência, aprofundando, mnemônico, diálogo, divergência), lista de questões, questões comentadas e gabarito, tudo lendo os fragmentos XML do `template.docx` oficial.
- **Sumário com paginação correta**: o serviço converte o documento com LibreOffice, mede em que página cada título caiu e grava os números no campo TOC. No Word, F9 também atualiza.
- **Marcações inline**: `**negrito**`, `__sublinhado__`, `%%vermelho%%` e `[IMAGEM n]` (imagens embutidas em base64).
- **Matemática**: LaTeX inline (`$...$`) convertido para OMML nativo do Word, incluindo acentos (`\bar`, `\hat`, `\vec`) e proteção contra falsos positivos como `R$`.
- **Fontes da marca embutidas** em cada documento gerado, garantindo tipografia idêntica em qualquer máquina e na conversão para PDF.
- **Extração robusta**: `.docx` e `.pptx` (texto, sublinhados, vermelhos, equações OMML linearizadas, imagens) e `.pdf` (PyMuPDF; detecção de sublinhado por caractere, descarte de cabeçalho/rodapé repetido, OCR Tesseract pt-BR para PDF escaneado).
- **Multi-marca**: a pasta `rp/` contém uma segunda máscara (outra identidade visual) reaproveitando o mesmo builder e o mesmo contrato de API.
- **Regras determinísticas**: o builder impõe as regras de diagramação do cliente (caixa cinza só no cabeçalho da questão, Certo/Errado sem alternativas, rótulos em negrito preto, títulos justificados, capa até 6 linhas a 22pt) independentemente do que a IA mandar.

## Arquitetura

```
n8n (orquestração) ──► POST /extract ──► texto + marcas + imagens
        │
        ▼ LLM estrutura teoria e questões
        │
        └──► POST /render ──► builder.py monta o DOCX a partir do template
                              └► LibreOffice mede a paginação do sumário
                                   └► DOCX final ──► Stirling PDF ──► PDF
```

| Arquivo | Função |
|---|---|
| `app.py` | API FastAPI: `/health`, `/extract`, `/render` |
| `builder.py` | Monta o documento: colhe fragmentos XML do template, aplica marcas, imagens, OMML, tabelas, questões e pagina o sumário |
| `template.docx` | Máscara oficial (estilos `TtuloCapa`, `Ttulo1`, `Ttulo2`, `Ttulo4`, `Ttulo5`, `Sumrio1/2`) |
| `fonts/` | Fontes da marca (colocar antes do build) |
| `rp/` | Segunda marca: `builder_rp.py`, `app.py`, `template.docx`, `Dockerfile` |
| `Dockerfile` | `python:3.12-slim` + LibreOffice Writer/Math + Tesseract pt-BR |

## Endpoints

| Método | Rota | Entrada | Saída |
|---|---|---|---|
| `GET` | `/health` | — | `{"ok": true}` |
| `POST` | `/extract` | multipart `file` (.docx, .pptx, .pdf) | `{text, chars, images, kind, pages?, ocr?}` |
| `POST` | `/render` | JSON (schema abaixo) | binário `.docx`; headers `X-Toc-Paginated` e `X-Toc-Error` |

Flag opcional no `/render`: `"update_fields": true` grava `<w:updateFields/>` para o Word oferecer atualizar o sumário ao abrir (ligada automaticamente se a paginação falhar).

## Schema do `/render` (resumo)

```json
{
  "filename": "Aula 2 - Poder Disciplinar.docx",
  "data": {
    "titulo": "PODER DISCIPLINAR E PODER HIERÁRQUICO",
    "apresentacao": ["parágrafo 1", "parágrafo 2"],
    "capitulos": [
      {"titulo": "CAPÍTULO 1 — ...", "blocos": [
        {"tipo": "paragrafo", "texto": "texto com **negrito** e $\\bar{x}=\\frac{a+b}{2}$"},
        {"tipo": "subtitulo", "texto": "..."},
        {"tipo": "tabela", "colunas": ["A","B"], "linhas": [["1","2"]], "legenda": "opcional"},
        {"tipo": "dica", "texto": "..."},
        {"tipo": "lei", "fonte": "Lei 8.112/1990, art. 126", "texto": "..."},
        {"tipo": "jurisprudencia", "tribunal": "STF", "referencia": "Súmula 18", "texto": "..."},
        {"tipo": "mnemonico", "titulo": "...", "texto": "..."},
        {"tipo": "aprofundando", "titulo": "...", "texto": "..."},
        {"tipo": "dialogo", "falas": [{"quem": "Aluno", "texto": "..."}]},
        {"tipo": "divergencia", "pergunta": "...", "posicoes": [{"rotulo": "Majoritária", "texto": "..."}]}
      ]}
    ],
    "questoes": [
      {"cabecalho": "01. (CEBRASPE / PF / 2024) Julgue o item.", "corpo": ["afirmativa"], "certo_errado": true},
      {"cabecalho": "02. (FGV / ... )", "corpo": ["A) ...", "B) ...", "C) ...", "D) ...", "E) ..."]}
    ],
    "comentarios": [
      {"cabecalho": "01. (CEBRASPE / ...)", "corpo": ["afirmativa"], "gabarito": "Errado", "comentario": "..."}
    ],
    "gabarito": [{"n": 1, "g": "Errado"}, {"n": 2, "g": "B"}],
    "formulas": [{"nome": "Média", "latex": "\\bar{x}=\\frac{\\sum x_i}{n}"}]
  }
}
```

## Rodando

### Docker (recomendado)

```bash
# coloque as fontes da marca em fonts/ antes do build
docker build -t renderdocxs .
docker run -p 8000:8000 renderdocxs
curl http://localhost:8000/health
```

Para a segunda marca: `docker build -f rp/Dockerfile -t renderdocxs-rp rp/`.

### Easypanel

Serviço **App > Dockerfile** apontando para a raiz (ou para `rp/`), porta interna **8000**, 1 vCPU / 1 GB RAM (o LibreOffice roda por chamada). No n8n, chame pela rede interna: `http://<projeto>_renderdocx:8000`.

### Local

```bash
pip install -r requirements.txt
sudo apt install libreoffice-writer libreoffice-math tesseract-ocr tesseract-ocr-por
uvicorn app:app --port 8000
```

## Testes rápidos

```bash
# extração
curl -F "file=@aula.pdf" http://localhost:8000/extract | jq '.kind, .pages, .chars'

# renderização
curl -X POST http://localhost:8000/render -H "Content-Type: application/json" \
  -d @exemplo.json -o saida.docx -D -
```

Se `X-Toc-Paginated: false`, o sumário foi gerado sem números; abra no Word e pressione F9.

## Decisões técnicas que valem registrar

- **Fragmentos XML do template em vez de python-docx puro**: garante que banners, cores e espaçamentos saiam idênticos ao modelo aprovado pelo cliente.
- **Fontes embutidas e retagueadas**: faces com nome composto (ex.: "Winner Sans Wide Bold") têm a name table reescrita para registrar a família exata usada no documento, evitando que o LibreOffice caia em pesos errados.
- **Não instalar as fontes da marca no container de conversão PDF**: fonte instalada tem prioridade sobre a embutida; uma família parcial força pesos errados.
- **Detecção de sublinhado no PDF por caractere**, não por span: o PyMuPDF funde runs sublinhados com vizinhos.
- **`$` precedido de letra não abre fórmula**: evita que `R$ 100` e `R$ 200` na mesma frase virem uma equação.

## Limitações conhecidas

- PDFs com duas colunas ou tabelas complexas saem em ordem de leitura aproximada; `.docx` continua sendo a entrada preferida.
- OCR não recupera vermelho nem sublinhado.
- Se o mesmo conteúdo chegar em `.docx` e `.pdf`, os dois são processados.

## Histórico

Ver [CHANGELOG.md](CHANGELOG.md) para o detalhamento de cada versão (v4 a v7.4).
