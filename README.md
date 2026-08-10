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
- v7.1: fim das páginas em branco. A quebra de capítulo/seção agora usa pageBreakBefore no parágrafo de respiro (no-op quando a página anterior termina cheia) e o builder remove respiros órfãos antes de cada banner. Validado com sweep de 25 comprimentos de conteúdo.
- v7.1: subtítulo anti-órfão. O subtítulo (tabela de 1 célula no template, sem suporte a keepNext) virou parágrafo com pBdr idêntico + keepNext encadeado: se não couber junto com o conteúdo no fim da página, descem juntos. Validado com sweep de 23 comprimentos.
- v7.1: revisão polida: linha-respiro invisível garante vão inferior simétrico ao topo (medido 21px/22px), chips sem bordas soltas no primeiro e no último bloco, e box de dica sem vermelho (%%alerta%% rebaixa para negrito preto na dica).
- v7.1: respiro do subtítulo assimétrico (360 antes, 120 depois): separa do assunto anterior e gruda no próprio conteúdo.
- v7.1: imagens dentro de questões. Linhas de corpo com o marcador [IMAGEM N] são substituídas pela imagem real (questões com texto de apoio, charge, campanha), tanto em Questões para Praticar quanto nas Comentadas.
- v7.1: suporte a RLM/matemática no /extract. Equações do editor do Word (OMML) são linearizadas ((a)/(b), a^b, a_b, √(x)) e símbolos de fonte Symbol (w:sym) viram Unicode (∧ ∨ ¬ → ⇒ ∀ ∃ ∈ ∪ ∩ ≤ ≥ ≠ √ ...). Símbolos digitados como texto já passavam normalmente.
