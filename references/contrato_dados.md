# Contrato de dados — DocMatrix 1.0.0

## Fonte da verdade

`extracao_base.json` contém dados determinísticos (documentos, blocos, tabelas, registros e menções a datas), `validacoes` estruturais, inconsistências e comparativo conservador. `docmatrix_resultado.json` é a junção final, sem alteração de blocos originais.

Campos do objeto raiz:

- `schema_version`: `1.0.0`.
- `skill`: `docmatrix`.
- `execucao`: modo, foco, estratégia híbrida, modos efetivos e situação da etapa de IA.
- `documentos[]`: id, nome, formato, SHA-256, tamanho, páginas, status, avisos, blocos, tabelas, registros, eventos.
- `validacoes[]` / `inconsistencias[]`: checagens e achados estruturais.
- `comparativo[]`: correspondência entre dois registros com origem.
- `anotacoes`: dados adicionados pela IA, diferenciados dos resultados determinísticos.
- `analise`: síntese interpretativa ou observação de extração puramente determinística.

Cada bloco possui `id`, `pagina` (número ou `null`), `localizacao`, `tipo`, `texto`.

Cada registro possui `id`, `documento_id`, `campo`, `valor_original`, `valor_normalizado`, `tipo`, `pagina`, `localizacao`, `bloco_id`, `origem`.

Cada tabela possui `id`, `pagina`, `localizacao`, `cabecalho_identificado`, `colunas`, `linhas`.

## Exemplo de `anotacoes.json` (opcional)

```json
{
  "registros": [
    {
      "id": "ia_r001",
      "documento_id": "doc_001",
      "campo": "objeto_principal",
      "valor_original": "Trecho textual exato identificado",
      "valor_normalizado": "Descrição curta do objeto",
      "tipo": "texto",
      "pagina": 2,
      "bloco_id": "doc_001:b000002",
      "localizacao": "página 2",
      "evidencias": ["doc_001:b000002"]
    }
  ],
  "eventos": [],
  "validacoes": [],
  "inconsistencias": [],
  "comparativo": [],
  "analise": {
    "resumo": "Síntese contextual validada pela IA",
    "conclusoes": ["Conclusão 1"],
    "limitacoes": ["Exemplo de ressalva"]
  }
}
```

> Exemplo fictício: jamais associe uma evidência ao trecho errado. `analise` pode conter listas, que são serializadas em texto na aba Síntese.

## Regras obrigatórias

- Identificadores podem conter zeros à esquerda; não coerça em número.
- Números acima da precisão exata suportada pelo Excel devem permanecer texto.
- Decimais ambíguos, datas incompletas, fuso e moeda sem unidade devem permanecer no formato original.
- Tabelas não devem ser agregadas sem conhecer as chaves e medidas.
- Inconsistências/validações adicionadas pela IA devem ter método e referências quando aplicáveis.
- Acrescente novos campos sem mudar o significado dos existentes; mudanças incompatíveis exigem nova versão principal do schema.
