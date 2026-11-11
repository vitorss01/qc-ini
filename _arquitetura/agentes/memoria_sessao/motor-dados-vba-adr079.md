---
name: motor-dados-vba-adr079
description: ADR-079 (20/10/2026) — ATUALIZAR em VBA (mEtlSync) em vez do Power Query; MOTOR_DADOS; Sync_Log; armadilhas medidas
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-20T18:21:49.246Z
---

ADR-079: com Cfg_Integracao MOTOR_DADOS = VBA, o ATUALIZAR roda mEtlReceb → mEtlCQ → mEtlQA orquestrados por mEtlSync (staging → validação → publicação só a partir da 1ª linha diferente → Sync_Log). As consultas PQ ficam como reserva (MOTOR_DADOS = PQ) e verificador. Instalador `instalar_adr079.py` (já no entregar.py), suíte `testes/qa_sync.py` (cenários em `estudos/teste_incremental.py`), reconciliação `estudos/teste_sync.py`.

Medido em 20/10: Bio de 118–218 s para ~13 s sem novidade e ~17 s com 1.240 novos; Hema ~19–25 s. Equivalência com o PQ: 0 diferenças (Bio 33.202×40, Hema 45.539×40).

Armadilhas:
- A Hema de produção tinha a tblCQ_Final sem MOTIVO_INATIVACAO (não era atualizada desde antes do ADR-070). O ConferirEsquemas refaz pelo PQ, no instalador.
- `Range.Resize(k)` pelo pywin32 vira Item(k): usar Range(r1, r2).
- Heredoc do bash estraga `\s`/`\n` (scripts de patch sempre pelo Write).
- Gravar texto numérico exige NumberFormat "@".

**Why:** o usuário pediu ATUALIZAR ≤ 15 s (máx 30), sem perda de dado e com governança (projeto prioritário).
**How to apply:** mudanças na regra de dados vão nos DOIS motores (pq/*.m e mEtl*), e equivalencia_*.py prova. Ver [[integracao-db-seac]], [[estado-entrega-0510]].
