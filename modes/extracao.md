# Modo `docmatrix:extracao`

1. Procure campos explícitos, tabelas, parágrafos, identificadores, datas e relações, sem especializar por domínio.
2. Mantenha o texto integral nos blocos do JSON, preservando referência para página/localização.
3. Extraia valores tabulares deterministicamente. Se a associação campo-valor exigir julgamento, registre em `anotacoes.json` com evidência.
4. Aponte ausência de texto digital, leitura truncada e tabelas ambíguas. Nunca preencha lacunas por suposição.
5. Organize a análise por prioridade do pedido; mantenha os demais dados estruturáveis nos arquivos.
