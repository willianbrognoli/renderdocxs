# Renderizador Águia - Carreiras Policiais

Microserviço que recebe o conteúdo estruturado (JSON) e devolve o .docx
diagramado exatamente na máscara do modelo (capa, sumário paginado, banners,
boxes, tabelas, questões, comentadas e gabarito), com highlight amarelo em FGV.

## Deploy no easypanel
1. Crie um serviço do tipo **App > Dockerfile** apontando para este diretório
   (suba o zip no GitHub ou use o build por upload).
2. Antes do build, coloque as fontes reais em `fonts/`.
3. Porta interna: **8000**. Ative HTTPS/domínio.
4. Recursos sugeridos: 1 vCPU / 1 GB RAM (LibreOffice roda por chamada).

## Endpoints
- `GET /health`
- `POST /extract` (multipart `file`) -> `{"text": "..."}`
- `POST /render` (JSON) -> docx binário

## Schema do /render
```json
{
  "filename": "Aula 2 - Poder Disciplinar.docx",
  "data": {
    "titulo": "PODER DISCIPLINAR E PODER HIERÁRQUICO",
    "apresentacao": ["parágrafo 1", "parágrafo 2"],
    "capitulos": [
      {"titulo": "CAPÍTULO 1 — ...", "blocos": [
        {"tipo": "paragrafo", "texto": "texto com **negrito**"},
        {"tipo": "subtitulo", "texto": "..."},
        {"tipo": "tabela", "colunas": ["A","B","C"], "linhas": [["1","2","3"]], "legenda": "opcional"},
        {"tipo": "mnemonico", "titulo": "...", "texto": "..."},
        {"tipo": "dica", "texto": "..."},
        {"tipo": "lei", "fonte": "Lei 8.112/1990, art. 126", "texto": "..."},
        {"tipo": "jurisprudencia", "tribunal": "STF", "referencia": "Súmula 18", "texto": "...", "observacao": "opcional"},
        {"tipo": "aprofundando", "titulo": "...", "texto": "..."},
        {"tipo": "dialogo", "falas": [{"quem": "Aluno", "texto": "..."}, {"quem": "Professor", "texto": "..."}]},
        {"tipo": "divergencia", "pergunta": "...", "posicoes": [{"rotulo": "Majoritária", "texto": "..."}]},
        {"tipo": "revisao", "titulo": "O QUE ESTUDEI", "linhas": [{"tema": "...", "itens": ["...", "..."]}]}
      ]}
    ],
    "questoes": [
      {"cabecalho": "01. (BANCA / Órgão / Ano) Enunciado...", "corpo": ["afirmativa"], "certo_errado": true},
      {"cabecalho": "02. (FGV / ... ) ...", "corpo": ["A) ...", "B) ...", "C) ...", "D) ...", "E) ..."]}
    ],
    "comentarios": [
      {"cabecalho": "01. (BANCA / ...)", "corpo": ["afirmativa"], "gabarito": "Errado", "comentario": "..."}
    ],
    "gabarito": [{"n": 1, "g": "Errado"}, {"n": 2, "g": "B"}]
  }
}
```
"Errado (   )" após o enunciado, banners, cores e espaçamentos saem idênticos ao modelo.
O sumário é um campo TOC real: o serviço converte com LibreOffice, mede a página
de cada título e grava os números; no Word, F9 também atualiza.
