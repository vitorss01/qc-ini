---
name: corrida-lote-adr080
description: "ADR-080 (21/10/2026) — eixo X do LJ = corrida do LOTE por nível (1,2,3...), pedido do usuário; RUN do dia não muda"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-21T12:53:26.419Z
---

Decisão do usuário (21/10/2026):
- cada nível tem a sua sequência de corridas;
- cada passagem do controle é uma corrida, mesmo à noite no mesmo dia;
- cada lote recomeça do 1;
- nunca há dois lotes no mesmo nível.

Implementação:
- mEstatistica.AtualizarCalc numera os plotados (elegível + X vermelho) por (lote, equipamento, nível) em todo o histórico do lote e publica em Eng_Saida BK..BM.
- O Calc tem as colunas CR..CT (N1..N3), com as bordas em 184/185, e os nomes xrunmin<i>/xrunmax<i>.
- No gráfico do nível i, o X de todas as séries aponta para essa coluna.
- O RUN yymmddkk (corrida do dia) continua sendo a chave do Westgard entre níveis e aparece na dica do ponto.

Com filtro de período ou mais de 180 corridas, o gráfico mostra o número real da corrida no lote (ex.: 58..237).

Instalador `instalar_adr080.py`, suíte `testes/qa_corrida_lote.py` (L01–L05). O instalar_adr078 preserva o X CR..CT.

ADR-081 (21/10, mesmo dia):
- A corrida conta SÓ os ativos; inativar pela ID renumera, sem lacuna.
- O X vermelho fica entre as corridas (k + j/(m+1); Eng BN..BV, Calc CU..DC).
- Duplo clique ou clique direito no ponto abre o frmInativar (instalar_frm_inativar.py).
- AtualizarSoEntradas recalcula a incerteza só do analito.
- Testes: qa_inativar_ux U00–U15.
- Entregue nos dois produtos em 21/10 (commit ed2aaf4).

**Why:** buraco no N1 quando só passava N2, e o "1" do gráfico não era a 1ª corrida do lote.
**How to apply:** qualquer mudança no eixo do LJ mexe nas duas coisas (motor + Calc CR..CT + séries). Ver [[linhas-alvo-adr078]], [[motor-dados-vba-adr079]].
