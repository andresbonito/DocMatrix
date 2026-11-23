# Validação e rastreabilidade

Critérios mínimos por documento: status da leitura, contagem de blocos e tabelas, largura das linhas tabulares e avisos de cobertura. Com múltiplos documentos, campos únicos com nomes literalmente iguais podem ser comparados sem inferir equivalência entre métricas diferentes.

Uma validação interpretativa só deve dizer `aprovado` se o critério estiver identificado e plenamente testado. Use `atencao`, `erro` ou `inconclusivo` nos demais casos, explicitando o método. Não confunda ausência de evidência com prova de inexistência.

Formato recomendado:

```json
{
  "id": "ia_v001",
  "documento_id": "doc_001",
  "tipo": "integridade",
  "status": "atencao",
  "criterio": "O campo informado deve estar preenchido",
  "metodo": "conferência das páginas identificadas",
  "descricao": "Não foi localizado o campo nas páginas verificadas",
  "evidencias": ["doc_001:b000003"],
  "origem": "interpretacao_ia"
}
```

Se houver cálculos: mantenha valor original, valor recalculado, fórmula/premissa, unidade, arredondamento e diferença. Registre comparação apenas entre grandezas efetivamente equivalentes.
