# Histórico de versões — DocMatrix

## 1.1.0 — 23/11/2026

- Documentação completa de instalação e uso para Claude Code e Codex (Windows, macOS, Linux).
- Instalador Python portátil `install.py`, com escopo global ou por projeto e atualização com `--force`.
- Dependências públicas declaradas em `requirements.txt`.
- Exportação Excel com seleção de backend: `auto`, `artifact_tool`, `xlsxwriter`.
- Backend XlsxWriter com `write_row`, `constant_memory=True`, cabeçalhos, filtros, congelamento e neutralização de fórmulas externas.
- CLI aceita `--excel-backend` nas operações `run` e `finalize`.
- Relatório `verificacao.json` passa a indicar o backend Excel escolhido.
- Novos testes de instalação, segurança do Excel e exportação de 2.500 registros.
- Mantidos os 5 modos e o contrato JSON 1.0.0, sem alterações incompatíveis nos dados.

## 1.0.0 — 23/11/2026

- Primeiro MVP: extração documental, comparação, validação, cronologia, síntese, geração Excel + JSON e rastreabilidade.
