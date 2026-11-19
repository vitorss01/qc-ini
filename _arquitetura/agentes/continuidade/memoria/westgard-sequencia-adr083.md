---
name: westgard-sequencia-adr083
description: ADR-083 (22/10/2026) Westgard do mesmo nível na sequência de ativos + aba técnica Westgard_Fotos (27/10) que audita todos os lotes
metadata:
  type: project
---

Decisão do usuário (22/10/2026): as regras do MESMO nível (2_2s entre corridas, SAME_LEVEL) andam pela sequência de ativos do nível (SeqDoNivel em AvaliarWestgard); as regras entre níveis continuam pelo RUN. O relatório de mudanças de veredicto do histórico inteiro é o CSV `<arquivo>_ADR083_mudancas_westgard.csv` do instalar_adr083.py, que roda PRIMEIRO no entregar.py (precisa do VBA antigo na cópia).

Aba Westgard_Fotos (27/10, pedido: "nova aba ... dado de vários lotes"): oculta e protegida, foto dos eventos OFICIAIS por lote × equipamento com impressão digital. AuditarFotosWestgard roda no ATUALIZAR e no AtualizarSoEntradas, refaz só os lotes mudados e audita WESTGARD_EVENTO_NOVO/REMOVIDO/CORRIDA_NOVA. Custa 0,3–0,45 s sem mudança e ~0,8 s a mais com 1 lote mudado. O teto do U13 da Bioquímica passou de 15 para 18 s.

O usuário mandou NÃO mexer em BI/Power BI por enquanto (27/10). Prazo: sexta 30/10/2026.

**Why:** o veredicto do mesmo nível não pode depender de corridas inativadas; e a auditoria tem de cobrir lotes fora do Painel.
**How to apply:** mudança no Westgard mexe no mEstatistica Bio e Hema + qa_westgard_sequencia (W01–W36). Ver [[corrida-lote-adr080]], [[manual-nao-substitui]], [[entrega-limpa-sozinha]].
