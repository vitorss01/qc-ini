---
name: adr085-painel-tela-unica
description: "ADR-085 (29/10/2026) — duplo clique nunca abre \"Formatar\"; ações NC/INATIVAR/COMENTAR; comentário em fonte única; Westgard no ponto; Registros assinados pela ID; H3 vazio adiado"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-29T16:53:50.687Z
---

Pedido de 29/10/2026: "reestruturação e otimização", 13 itens. Acréscimo do usuário no mesmo dia: o duplo clique no gráfico NUNCA pode abrir o eixo nem a barra "Formatar" do Excel.

Decisões:
- `clsCht`: `Cancel = True` é a 1ª instrução no duplo clique e no clique direito. Fora de um ponto aparece só a dica (`mUI.CliqueNoGrafico`). `GarantirHooks` religa os gráficos no vigia.
- Estados sem estado novo:
  - NC = X_VERMELHO;
  - INATIVAR = NAO_PLOTAR;
  - os dois saem da média, do DP e do Westgard (regra mantida).
- Comentário: uma linha por ID em `tblComentariosTecnicos` (`GravarComentarioTecnico`), gravada antes do estado, com reversão (`DesfazerComentario`). A ação "só comentário" aplica na Principal sem ATUALIZAR (`AplicarComentarioNaBase` + `mEtlCQ.ComentarioConsolidado`).
- Gancho do QA: `ACAO|<NC|INATIVAR|COMENTAR|REATIVAR|TIRAR_X>|texto`. O formato antigo `CONFIRMAR|texto|1/0` continua.
- H3 vazio = todos os lotes: NÃO entrou no ADR-085. Exige uma chave de corrida composta (o mesmo RUN pode existir em dois lotes no mesmo dia), alvo por ponto, Westgard por lote e limites em degrau. Ficou para o ADR-086.
- VBE: "msg =\"" sem espaço faz o `codigo_atual` acusar "difere da fonte". Usar `difmod.py` (scratchpad) para achar a linha.

**Why:** prazo 30/10. O motor multilote arrisca o que já está validado.
**How to apply:** suíte `testes/qa_adr085.py` (Z01–Z14) e instalador `instalar_adr085.py` (no `entregar.py`, depois do adr084). Ver [[adr084-usuarios-calibracao]], [[corrida-lote-adr080]], [[westgard-sequencia-adr083]].
