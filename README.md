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
sem linhas de marcação Certo/Errado (padrão v4), banners, cores e espaçamentos saem idênticos ao modelo.
O sumário é um campo TOC real: o serviço converte com LibreOffice, mede a página
de cada título e grava os números; no Word, F9 também atualiza.


## v4 (padrão do PDF de referência)
- Marcas inline: **negrito**, __negrito sublinhado__, %%negrito vermelho%%. O /extract anota sublinhados e vermelhos do docx fonte com __ e %%.
- Revisão: moldura dourada, título centrado, 2 colunas (chip escuro com tema em CAIXA ALTA centrado + itens com seta dourada em fonte de corpo, recuo deslocado).
- Questões sem linhas "Certo (   )/Errado (   )". Banner final: GABARITO FINAL.
- Títulos de mnemônico e pergunta/rótulos de divergência em preto.
- Fontes da marca: o builder embute as fontes da pasta `fonts/` do serviço (ou `/usr/share/fonts/truetype/aguia`) dentro de cada docx gerado, garantindo a tipografia correta em qualquer máquina e na conversão para PDF.
- v6: faces de nome mesclado (ex.: Winner Sans Wide Bold) são retagueadas internamente (name table, OS/2, head) para se registrarem com a família exata usada no documento, garantindo que LibreOffice/Stirling usem a fonte embutida.
- IMPORTANTE (Stirling): não instale as fontes da marca no container do Stirling. Fonte instalada tem prioridade sobre a embutida, e uma família parcial (sem Regular/Bold) força pesos errados (Thin/ExtraBold). Com o container limpo, o PDF usa as fontes embutidas do docx.
- v7: revisão no padrão final: moldura esticada até a largura do texto (sectPr), bloco interno centralizado com respiro simétrico, temas centralizados nos chips, itens à esquerda com recuo deslocado, setas pretas, título sempre "O QUE ESTUDEI" (forçado pelo builder) e sem highlight amarelo de banca nas seções de questões.
