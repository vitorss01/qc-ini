# Regras comuns a todos os agentes do QC_INI

Todo agente lê este arquivo antes de começar. Ele vale para o Usuário Simulado, o Especialista em CQ, o UX/UI, o
Tester/QA e o Orquestrador. Arquitetura e fluxo: `README.md` nesta pasta.

## 1. Onde está cada coisa

| O quê | Onde |
|---|---|
| Arquivos de produção (QC em uso) | raiz do repositório: `QC_Bioquimica.xlsm`, `QC_Hematologia.xlsm` (variável `QC_RAIZ` nas entregas) |
| Fontes VBA (UTF-8) | `_arquitetura/integracao_dbseac/src/{comum,bio,hema}` |
| Decisões (o que o sistema deve fazer) | `_arquitetura/integracao_dbseac/ADR-0xx.md`; ADR antigos em `_arquitetura/ARCHITECTURE_DECISIONS.md` |
| Manual do laboratório | `_arquitetura/integracao_dbseac/GUIA_DE_USO.md` |
| Suítes de teste e infraestrutura COM | `_arquitetura/integracao_dbseac/testes/` (`qa_final.QA` abre uma cópia e revela as abas); helper `_arquitetura/etapa1_multilote/xlh.py` |
| Metodologia de bias, ET, Sigma | ADR-035 (`ARCHITECTURE_DECISIONS.md`), ADR-064, ADR-067, ADR-068, ADR-069; `src/comum/mCEQ.bas` (`BiasEQ` = bias do ET/Sigma; `ViesEQ` = triagem e u(bias) da incerteza); `src/bio/mEstatPeriodo.bas`; `GUIA_DE_USO.md` §4–5 |
| Relatórios dos agentes | `_arquitetura/agentes/relatorios/<ciclo>/` |
| Registro de ocorrências | `_arquitetura/agentes/OCORRENCIAS.md` (só o Orquestrador escreve) |
| Ferramentas dos agentes (calculadora independente etc.) | `_arquitetura/agentes/ferramentas/` |

## 2. Segurança dos dados (proibições)

1. **Produção nunca é aberta para escrita.** Toda execução trabalha numa **cópia** em
   `%TEMP%\qcw\auditoria\<ciclo>\` (`shutil.copy2` a partir da raiz). Só o Orquestrador troca a produção, e só pelo
   `entregar.py` (backup e conferência de hash).
2. É proibido excluir ou alterar registros reais, mudar resultados históricos, criar duplicidade em base operacional,
   ou salvar qualquer coisa nos arquivos da raiz ou no DB_SEAC.
3. **Dados de teste** só entram na cópia. Eles são marcados para nunca se confundirem com dados reais: comentário ou
   operador `TESTE-AUDITORIA`, e valores obviamente sintéticos registrados no relatório. A cópia é descartada no fim.
4. **Excel:** só **um** agente usa Excel por vez. Nada de Excel durante uma entrega. Encerre **somente o PID que o
   seu script criou** (`xlh.Excel().pid`; `qa_final.QA` imprime `EXCEL_PID`). Uma tarefa agendada abre Excel a cada
   2 min para o DB_SEAC de produção, e matar Excel por nome ou horário pode corromper esse arquivo.
5. **Nunca capture a tela do usuário.** Os dados podem ter informações de pacientes. Imagem de tela só pela
   exportação do próprio Excel da cópia (`CopyPicture`/`Export`, como `visual_telas.py` faz).
6. **O repositório remoto é público.** Nada de senha, caminho UNC da rede do laboratório ou `.xlsm` em arquivo
   versionado (relatórios inclusive). Se um caminho de rede aparecer numa evidência, escreva `<UNC>`.
7. **Power BI / BI:** não mexer.

## 3. Regras de método

1. **Não altere fórmulas, regras de negócio, critérios estatísticos ou métodos de cálculo.** O que parecer errado vira
   ocorrência. A mudança só acontece depois do parecer do Especialista em CQ **e** da autorização explícita do usuário,
   dada ao Orquestrador.
2. **Nunca declare aprovado o que não foi executado.** Cada teste tem um destes estados: `APROVADO`, `REPROVADO`, `NAO_EXECUTADO`
   (com o motivo) ou `BLOQUEADO` (com o que falta).
3. Uma correção **não é validada** por quem a implementou. Valida o Tester/QA e o Usuário Simulado, que reexecuta o
   mesmo fluxo.
4. **A metodologia vem do projeto, não de livro.** As fórmulas de referência (bias %, ET = |bias| + 1,65·CV, Sigma =
   (ETp − |bias|)/CV) servem para conferir. A definição que vale é a documentada nos ADR. Se o QC_INI usa outra, ela é
   reportada como diferença de método e vai ao Especialista em CQ. **Não é** reprovada automaticamente.
5. **Precisão:** compare com a precisão total do valor armazenado na célula (`Value2`), não com o texto formatado.
   A tolerância padrão é de 1e-9 relativo para valores calculados pelo mesmo método. Qualquer outra tolerância é
   justificada no relatório.
6. **Evidência reproduzível:** arquivo de origem (nome da cópia + sha256 curto da produção copiada), aba e célula,
   IDs dos registros, valores de entrada, e o script que gerou o número.

## 4. Classificação das ocorrências

| Prioridade | Quando |
|---|---|
| **Crítico** | Erro de cálculo; corrupção de dados; associação errada entre lotes, analitos, equipamentos ou níveis; estatística incorreta; falha que tira a confiabilidade do sistema |
| **Alto** | Funcionalidade importante que não funciona corretamente, ou risco operacional relevante |
| **Médio** | Problema de usabilidade, inconsistência de visualização ou dificuldade operacional |
| **Baixo** | Melhoria de apresentação, conveniência ou acabamento visual |

## 5. Ambiente

- Windows, Python com `pywin32` e `openpyxl`. Use `PYTHONIOENCODING=utf-8` e grave os scripts em UTF-8.
- Rodadas longas de Excel: `python -u` com saída para arquivo de log, e um teto de tempo por macro
  (`xlh.Excel.run(..., teto=)`).
- Em bash, nada de barra invertida dentro de heredoc. Para caminhos Windows, use Python ou PowerShell.
- Responda e escreva os relatórios em **português**.
