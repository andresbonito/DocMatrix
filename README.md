# DocMatrix v1.1.0

**Repositório oficial:** [andresbonito/DocMatrix](https://github.com/andresbonito/DocMatrix)

**Skill universal de análise documental com exportação de Excel e JSON**, preparada para Claude Code, Codex e agentes compatíveis com o padrão Agent Skills. O DocMatrix separa extração determinística em Python da interpretação semântica feita pelo agente. Não há especialização obrigatória por área.

> **MVP:** a Skill instrui o agente a analisar documentos e entregar uma resposta interpretativa mais dois arquivos: `docmatrix_resultado.xlsx` e `docmatrix_resultado.json`. O uso implícito depende das regras de descoberta e permissão do Claude Code/Codex. A CLI Python não chama um modelo de IA sozinha.

## Instalação rápida pelo GitHub

O modo recomendado é instalar diretamente do **[repositório DocMatrix](https://github.com/andresbonito/DocMatrix)** com a CLI [Skills](https://github.com/vercel-labs/skills), sem baixar ZIP ou copiar pastas manualmente. É necessário **Node.js 22.20+** (com `npx`) para este método; as rotinas de análise também exigem **Python 3.10+**.

**Claude Code — instalação global:**

```bash
npx skills add https://github.com/andresbonito/DocMatrix --skill docmatrix -g -a claude-code -y
```

**Codex — instalação global:**

```bash
npx skills add https://github.com/andresbonito/DocMatrix --skill docmatrix -g -a codex -y
```

**Claude Code e Codex — instalar nos dois simultaneamente:**

```bash
npx skills add https://github.com/andresbonito/DocMatrix --skill docmatrix -g -a claude-code -a codex -y
```

`-g` disponibiliza a Skill globalmente; **remova `-g`** para instalar apenas no projeto atual. O parâmetro `-y` dispensa perguntas de confirmação. No Windows PowerShell, se `npx` estiver bloqueado pela política de execução, experimente `npx.cmd` nos mesmos comandos. Para conferir se a Skill foi detectada antes de instalar:

```bash
npx skills add https://github.com/andresbonito/DocMatrix --list
```

**Importante:** a CLI instala os arquivos da Skill, **não as dependências Python**. Após instalar, siga a [seção 2](#2-pré-requisitos-e-instalação-de-bibliotecas) no diretório instalado; depois, reinicie o agente e faça o teste das seções 3 ou 4. Para revisar instalações e atualizações:

```bash
npx skills ls -g
npx skills check
npx skills update docmatrix -g
```

A atualização com `npx skills update` exige uma instalação gerenciada pela CLI Skills. Cópias feitas anteriormente pelo `install.py` não são rastreadas por esse comando. O instalador local continua disponível como alternativa sem Node.js.

## Índice

- [Instalação rápida pelo GitHub](#instalação-rápida-pelo-github)
1. [O que faz](#1-o-que-faz)
2. [Pré-requisitos e instalação de bibliotecas](#2-pré-requisitos-e-instalação-de-bibliotecas)
3. [Instalar no Claude Code](#3-instalar-no-claude-code)
4. [Instalar no Codex](#4-instalar-no-codex)
5. [Como usar e como ativar automaticamente](#5-como-usar-e-como-ativar-automaticamente)
6. [Execução local passo a passo](#6-execução-local-passo-a-passo)
7. [Backends de Excel e desempenho](#7-backends-de-excel-e-desempenho)
8. [Abas, arquivos e JSON](#8-abas-arquivos-e-json)
9. [Suporte a formatos e limitações](#9-suporte-a-formatos-e-limitações)
10. [Testes e solução de problemas](#10-testes-e-solução-de-problemas)
11. [Documentação oficial](#11-documentação-oficial)

## 1. O que faz

O DocMatrix combina cinco modos reutilizáveis em um único núcleo:

| Modo | Resultado |
|---|---|
| `docmatrix` | Seleção automática de operações pertinentes à solicitação |
| `docmatrix:extracao` | Campos, valores, tabelas e evidências documentais |
| `docmatrix:comparacao` | Diferenças entre documentos, versões ou registros equivalentes |
| `docmatrix:validacao` | Conferência de estrutura e inconsistências verificáveis |
| `docmatrix:cronologia` | Referências temporais e sequenciamento de eventos |
| `docmatrix:sintese` | Síntese interpretativa fundamentada nas fontes |

**Atenção:** nomes com `:` são comandos lógicos descritos na Skill, não subskills oficiais da plataforma. O agente interpreta o sufixo dentro da mensagem; no Claude Code, o atalho nativo é `/docmatrix` e, no Codex, `$docmatrix`.

**Estratégia híbrida:** a extração busca amplamente dados estruturáveis, enquanto a resposta e abas adicionais destacam o objetivo do usuário. O conteúdo do documento não é uma fonte de instruções para executar comandos.

## 2. Pré-requisitos e instalação de bibliotecas

- **Python 3.10 ou superior**, com `pip` e execução de processos permitida.
- **Claude Code ou Codex instalado e funcionando**, se desejar uso como Skill. Não é necessário instalar ambos.
- **XlsxWriter** para gerar Excel em Python convencional; `artifact_tool` é opcional, apenas em ambientes que já o fornecem.
- **PyMuPDF** (principal) ou **pypdf** para PDF com texto nativo. PDFs digitalizados ainda não usam OCR automático.

Instale as dependências dentro de um ambiente virtual na pasta em que a Skill foi **efetivamente instalada** (ou na raiz do clone `DocMatrix`, se estiver executando apenas a CLI Python). Essa pasta contém `SKILL.md` e `requirements.txt`. Se utilizou `npx skills add`, consulte `npx skills ls -g` e localize a instalação de `docmatrix`; o cliente pode usar links simbólicos para compartilhar uma única cópia.

### Windows — PowerShell

```powershell
# Execute na pasta instalada, que contém SKILL.md e requirements.txt
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe scripts\docmatrix.py --help
```

Se o Python não estiver registrado como `py`, use `python -m venv .venv` na primeira linha de criação.

### Linux ou macOS — Terminal

```bash
# Execute na pasta instalada, que contém SKILL.md e requirements.txt
python3 -m venv .venv
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python scripts/docmatrix.py --help
```

`requirements.txt` contém apenas bibliotecas públicas. O pacote não solicita tokens ou chaves de API para executar as rotinas determinísticas. A análise feita pelo agente usa seu ambiente e permissões próprios.

**Importante para uso como Skill:** depois de instalar pelo GitHub ou copiar a pasta para o agente, instale as dependências **na pasta instalada**, criando o `.venv` lá. A Skill instrui o agente a usar o Python desse `.venv` quando existir; se você instalou só no clone ou no pacote descompactado, o agente não verá necessariamente essas bibliotecas. Uma atualização da Skill pode exigir recriar o `.venv`, caso o instalador substitua os arquivos.

## 3. Instalar no Claude Code

Os locais documentados pelo Claude Code são:

| Escopo | Pasta final (com `SKILL.md` dentro) |
|---|---|
| Global, para seus projetos | `~/.claude/skills/docmatrix/` |
| Apenas um projeto | `<projeto>/.claude/skills/docmatrix/` |

### Instalação recomendada — GitHub + Skills CLI

```bash
npx skills add https://github.com/andresbonito/DocMatrix --skill docmatrix -g -a claude-code -y
```

Para **instalar somente no projeto atual**, execute a partir da raiz do projeto e remova `-g`. Em seguida, instale as dependências Python conforme a seção 2.

### Alternativa — Git clone + instalador Python

```bash
git clone https://github.com/andresbonito/DocMatrix.git
cd DocMatrix
python install.py --host claude --scope user
```

No Windows, pode substituir `python` por `py`. Para adicionar somente a um projeto:

```bash
python install.py --host claude --scope project --project "/caminho/do/projeto"
```

O parâmetro `--force` atualiza uma instalação DocMatrix já existente. Ele substitui a pasta antiga; faça backup de alterações pessoais antes de utilizá-lo. O instalador **não** instala pacotes Python nem altera seu `CLAUDE.md`.

### Instalação manual no Windows (PowerShell)

Alternativa para quem utiliza o ZIP da Skill, no diretório que contém a pasta `docmatrix`:

```powershell
New-Item -ItemType Directory -Force "$HOME\.claude\skills" | Out-Null
Copy-Item -Path ".\docmatrix" -Destination "$HOME\.claude\skills" -Recurse -Force
```

Para instalar apenas no projeto atual, substitua o destino por `".\.claude\skills"`.

### Instalação manual no macOS/Linux

```bash
mkdir -p "$HOME/.claude/skills"
cp -R ./docmatrix "$HOME/.claude/skills/"
```

Para um projeto, utilize `mkdir -p .claude/skills && cp -R ./docmatrix .claude/skills/` na raiz dele.

### Testar a instalação

1. Abra ou reinicie o Claude Code no projeto.
2. Experimente `/docmatrix` ou digite `/` e procure por `docmatrix`.
3. Faça um pedido com arquivos, por exemplo: `/docmatrix Compare os documentos e gere Excel + JSON.`
4. Se aparecer erro de biblioteca, instale dependências na pasta instalada (seção 2).

O Claude Code também pode invocar a Skill de modo implícito quando a descrição corresponder à tarefa. O `SKILL.md` carrega seu corpo apenas quando necessário.

## 4. Instalar no Codex

Locais documentados para Skill por pasta:

| Escopo | Pasta final (com `SKILL.md` dentro) |
|---|---|
| Global, para seus projetos | `~/.agents/skills/docmatrix/` |
| Apenas um projeto | `<projeto>/.agents/skills/docmatrix/` |

### Instalação recomendada — GitHub + Skills CLI

```bash
npx skills add https://github.com/andresbonito/DocMatrix --skill docmatrix -g -a codex -y
```

Para **instalar somente no projeto atual**, execute a partir da raiz do projeto e remova `-g`. Em seguida, instale as dependências Python conforme a seção 2.

### Alternativa — Git clone + instalador Python

```bash
git clone https://github.com/andresbonito/DocMatrix.git
cd DocMatrix
python install.py --host codex --scope user
```

Ou apenas no projeto:

```bash
python install.py --host codex --scope project --project "/caminho/do/projeto"
```

### Instalação manual no Windows (PowerShell)

```powershell
New-Item -ItemType Directory -Force "$HOME\.agents\skills" | Out-Null
Copy-Item -Path ".\docmatrix" -Destination "$HOME\.agents\skills" -Recurse -Force
```

Para o projeto, substitua o destino por `".\.agents\skills"`.

### Instalação manual no macOS/Linux

```bash
mkdir -p "$HOME/.agents/skills"
cp -R ./docmatrix "$HOME/.agents/skills/"
```

### Testar no Codex

1. Abra ou reinicie o Codex no projeto.
2. Digite `$docmatrix` para referenciar explicitamente a Skill (ou selecione-a no menu de Skills, se disponível).
3. Exemplo de pedido: `$docmatrix Analise estes relatórios e entregue Excel + JSON com evidências.`
4. Instale as dependências no diretório da Skill caso não estejam disponíveis no Python utilizado.

**Nota de compatibilidade:** versões e instaladores antigos do Codex podem utilizar `~/.codex/skills`. Esta versão segue os caminhos descritos na documentação atual (`~/.agents/skills`). A CLI `npx skills add` seleciona os diretórios compatíveis com o agente e pode utilizar links simbólicos. Em caso de discrepância, confira a listagem de Skills do seu cliente e a documentação da versão instalada.

## 5. Como usar e como ativar automaticamente

Você pode usar uma invocação explícita ou formular um pedido normal de análise documental. Exemplos:

```text
Analise o contrato anexo, destaque obrigações e gere o Excel e JSON com as evidências.
```

```text
/docmatrix Compare as duas versões do relatório. Use o modo comparacao.
```

```text
$docmatrix Compare e valide os documentos anexados. Use o modo comparacao.
```

### Como aumentar a chance de acionamento automático

A descrição de `SKILL.md` foi escrita para ser selecionada quando você pedir leitura analítica, conferência, comparação ou síntese de documentos. Não há como a própria Skill impor execução em uma conversa na qual não foi carregada.

Se quiser reforçar o comportamento **em um projeto**, adicione uma regra curta ao `CLAUDE.md` (Claude Code) ou `AGENTS.md` (Codex) da raiz desse projeto:

```md
Quando a tarefa envolver análise de arquivos/documentos, utilize a Skill docmatrix
se ela estiver disponível. Entregue uma análise rastreável, um Excel e um JSON.
Não altere fontes originais e declare limitações de leitura.
```

Essas regras não substituem as permissões e as instruções de maior prioridade do agente. Um pedido explícito do usuário pode alterar o formato de saída.

## 6. Execução local passo a passo

### Caminho A — Extração determinística com Excel + JSON

Entre na pasta da Skill e rode:

```bash
python scripts/docmatrix.py run --input "examples/exemplo_a.txt" "examples/exemplo_b.txt" "examples/exemplo_tabela.csv" --output-dir "saida" --mode auto --focus "Comparar valores e prazos" --excel-backend xlsxwriter
```

Troque `python` pelo executável Python da sua instalação, por exemplo `./.venv/bin/python` (macOS/Linux) ou `.\.venv\Scripts\python.exe` (PowerShell).

Resultados em `saida/`:

```text
extracao_base.json           # dados extraídos pelo Python, antes da interpretação
contexto_ia.md               # contexto compacto para economizar tokens
verificacao.json             # inspeção técnica dos arquivos

docmatrix_resultado.json     # registros + evidências + validações
docmatrix_resultado.xlsx     # planilha Excel
```

Este caminho **não executa interpretação semântica de IA**. Ele é útil para extração e conferência estrutural.

### Caminho B — Extração + interpretação pelo agente

**1. Prepare:**

```bash
python scripts/docmatrix.py prepare --input "examples/exemplo_a.txt" "examples/exemplo_b.txt" --output-dir "saida" --mode comparacao --focus "Mudanças de prazo e valores"
```

**2. Agente:** leia `saida/contexto_ia.md`, amplie a consulta a `saida/extracao_base.json` quando necessário e produza `saida/anotacoes.json` conforme `references/contrato_dados.md`. Preserve `documento_id`, `bloco_id` e `evidencias` válidas.

**3. Finalize:**

```bash
python scripts/docmatrix.py finalize --base "saida/extracao_base.json" --annotations "saida/anotacoes.json" --output-dir "saida" --excel-backend auto
```

**4. Verifique:**

```bash
python scripts/docmatrix.py verify --json "saida/docmatrix_resultado.json" --excel "saida/docmatrix_resultado.xlsx"
```

Se não houver interpretação do agente, omita `--annotations` e o JSON marcará `interpretacao_ia: nao_executada`.

## 7. Backends de Excel e desempenho

O DocMatrix possui **dois backends de exportação**:

| Backend | Como escreve | Melhor aplicação |
|---|---|---|
| `artifact_tool` | Matrizes em intervalos, com lotes de até 1.500 linhas | Ambientes que oferecem essa biblioteca |
| `xlsxwriter` | Uma chamada `write_row` por linha, em ordem, com `constant_memory=True` | Claude Code, Codex e Python local |
| `auto` | Usa `artifact_tool` quando disponível; caso contrário `XlsxWriter` | Comportamento padrão |

O XlsxWriter **não oferece escrita matricial nativa equivalente ao `artifact_tool`**. Sua alternativa é `write_row`: evita o loop explícito de escrita célula a célula, usa streaming e não mantém toda a planilha em memória. Em contrapartida, `constant_memory` não suporta `add_table()`; utilizamos **autofiltro e cabeçalhos formatados** nas abas criadas por ele. A estrutura tabular das informações é preservada em ambos os casos.

Se a seleção automática não funcionar, force o backend:

```bash
python scripts/docmatrix.py run --input "examples/exemplo_a.txt" --output-dir "saida" --excel-backend xlsxwriter
```

A opção escolhida aparece no retorno da CLI, no `verificacao.json` e na aba **Visão Geral**. Não é necessário instalar `artifact_tool` manualmente em Claude Code e Codex.

**Economia de tokens:** a Skill produz o JSON completo diretamente por Python e envia inicialmente ao agente apenas `contexto_ia.md` (12.000 caracteres por padrão, configurável com `--context-chars`). Não reduz o JSON para caber no contexto. A interpretação semântica deve consultar os blocos completos relevantes antes de concluir.

## 8. Abas, arquivos e JSON

O Excel mantém quatro abas fixas: **Visão Geral**, **Dados Extraídos**, **Evidências** e **Validações**. As abas **Comparativo**, **Cronologia**, **Síntese**, **Índice Tabelas** e **Tabela_NNN** são criadas conforme a necessidade.

O JSON é a representação estruturada compartilhada usada pelo Excel e pela resposta analítica, com identificadores de documento e evidência, valores originais e normalizados separados, origem e status. O contrato detalhado está em `references/contrato_dados.md`.

O limite padrão do Excel é 1.048.576 linhas por aba (incluindo cabeçalho). O script fragmenta conjuntos acima desse tamanho. Os dados originais continuam no JSON, que é o artefato de referência para integrações.

## 9. Suporte a formatos e limitações

| Formato | Cobertura MVP | Limitação |
|---|---|---|
| PDF | Texto nativo e tabelas detectáveis | Sem OCR automático; páginas digitalizadas exigem outro fluxo |
| DOCX | Corpo textual e tabelas | Cabeçalhos, figuras e revisões podem não ser extraídos |
| XLSX | Abas, células e fórmulas **como dados** | Não calcula fórmulas, nem preserva toda a formatação original |
| CSV/TSV | Linhas e colunas | Detecção de delimitador por heurística |
| TXT/MD | Texto e padrões chave-valor | Contexto semântico profundo depende do agente |
| JSON/XML | Campos e estruturas textuais | Estruturas muito complexas têm limitações e salvaguardas |
| HTML | Texto visível | Scripts e interações dinâmicas não são executados |

O DocMatrix não navega em links contidos nos documentos nem executa macros. Dados sensíveis continuam nos arquivos gerados; revise a distribuição. Inferência do agente é sempre distinguida de dados originais.

## 10. Testes e solução de problemas

Execute os testes (a partir de `docmatrix/`):

```bash
python -m unittest discover -s tests -v
```

**Erro: `No module named 'xlsxwriter'`** — instale `requirements.txt` no Python realmente usado pela CLI ou ative seu `.venv`.

**Erro: `Nenhum backend Excel disponível`** — instale `XlsxWriter` ou use explicitamente `artifact_tool` num ambiente que o forneça.

**Skill não aparece no Claude Code/Codex** — confira a pasta e se existe `SKILL.md` diretamente em `docmatrix/`, reinicie a sessão e tente `/docmatrix` (Claude) ou `$docmatrix` (Codex).

**PDF sem conteúdo** — confirme se é pesquisável/tem texto selecionável. PDFs digitalizados não recebem OCR automático nesse MVP.

**Planilha abre sem tabelas nativas** — é esperado com o backend XlsxWriter em memória constante; o Excel continua tendo dados e autofiltros.

## 11. Documentação oficial

- [Repositório DocMatrix — Código e versões](https://github.com/andresbonito/DocMatrix)
- [Skills CLI — Instalação e atualização via GitHub](https://github.com/vercel-labs/skills)
- [Claude Code — Skills](https://code.claude.com/docs/en/skills)
- [OpenAI — Criar Skills para Codex e ChatGPT](https://learn.chatgpt.com/docs/build-skills)
- [Agent Skills — Especificação](https://agentskills.io/specification)
- [XlsxWriter — Memória e desempenho](https://xlsxwriter.readthedocs.io/working_with_memory.html)
- [XlsxWriter — API de planilhas](https://xlsxwriter.readthedocs.io/worksheet.html)

---

**Versão:** 1.1.0. **Projeto:** DocMatrix. Distribuído como pacote independente de agentes, sem afiliação oficial à OpenAI ou à Anthropic. Verifique licenças e termos das dependências utilizadas em redistribuições.
