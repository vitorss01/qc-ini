---
name: entrega-limpa-sozinha
description: Entregar sempre de cópia limpa do commit (QC_RAIZ) e sem nenhum Excel em paralelo — lições de 21/10/2026
metadata:
  node_type: memory
  type: feedback
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-11-14T00:50:55.623Z
---

Em 21/10/2026, a entrega da Bioquímica falhou 4 vezes por erro de processo meu, sem tocar a produção:
- o src/ tinha código de outro ADR em andamento e a entrega o instalou;
- testes rodando em paralelo derrubaram o qa_desempenho (P01 832 ms contra teto de 600);
- testes que dependiam dos dados (qa_blocos_lote BN, qa_varredura V02);
- um instalador antigo (adr070) conferia o texto da dica no formato velho.

O que funcionou:
1. `git archive HEAD _arquitetura` para `%TEMP%\qcw\clean\qc-ini`;
2. `$env:QC_RAIZ` apontando para o qc-ini real e Set-Location na cópia limpa (scripts `qcw\entregarNN.ps1`);
3. nada de Excel até terminar.

A entrega 26 passou nos dois produtos (20 suítes cada; ver MANIFESTO).

**Why:** cada entrega leva de 2 a 5 h de Excel no PC do usuário, e cada falha custa horas.
**How to apply:**
- antes de entregar, rodar numa cópia os instaladores que fazem fumaça de gráfico e dica (adr070, 078, 080, frm) depois de mudar formato;
- não rodar teste nenhum durante a entrega, nem trabalho pesado de CPU: agentes/workflows lendo .xlsm como zip, parsing de tabelas inteiras. Em 13/11/2026 isso derrubou o qa_desempenho P02 da Hema (mediana de 335 ms contra teto de 300; normal 200). Durante a entrega, só edição leve e lint;
- se uma entrega parar numa suíte, retomar na mesma cópia instalada com `testes/rodar_suites.py --a-partir <suite>` e trocar com `agentes/ferramentas/trocar_producao.py`, que repete as checagens do entregar.py;
- a retomada roda SEMPRE de uma cópia limpa do commit (`git archive` + `tar -xf`; o pipe do PS 5.1 quebra o tar). Nunca rodar do repositório. Em 13/11 eu editava os testes do ADR-094 no repo enquanto a retomada da 39 os lia: tive de parar e recomeçar;
- entregar os dois produtos juntos quando os dois mudaram.

Ver [[motor-dados-vba-adr079]], [[corrida-lote-adr080]], [[trabalhar-direto-na-main]].
