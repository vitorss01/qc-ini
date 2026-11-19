---
name: manual-nao-substitui
description: D09b (06/10) REVOGADA em 22/10/2026 pelo ADR-082 (manual substitui via SUBSTITUI_ID); D-01 decidida (ADR-072)
metadata:
  type: feedback
---

**REVOGADA em 22/10/2026 (ADR-082):** "o manual deve prevalecer em relação ao automático". O usuário escolhe explicitamente, pela coluna SUBSTITUI_ID da tblResultados_Manuais (MOTIVO obrigatório); o automático vira SUBSTITUIDO_MANUAL e some do gráfico; manual sem SUBSTITUI_ID no mesmo instante continua CONFLITO_MANUAL. Histórico abaixo.

Em 06/10/2026 o usuário decidiu (D09b): "não devemos sobrescrever resultados automáticos por manuais". Automático ruim se inativa; a aba Digitar Resultados é só para o que o interfaceamento não trouxe. Manual com a mesma chave/instante de um automático = CONFLITO_MANUAL + alerta A09 (comportamento que já existia; mensagem atualizada no commit 17a2447).

D-01 (viés do Sigma): o usuário trouxe uma revisão da literatura com estas conclusões. O bias se guarda COM sinal. No Sigma entra |viés sistemático estimado| (Westgard 2018: o bias é absoluto na fórmula). Mean(|bias|) não é viés sistemático e fica como indicador separado ("magnitude média do desvio EQA"). Antes de mudar, comparar nos dados reais Mean|bias| × |Mean bias| × RMS × regressão → classe Sigma → regra de Westgard. Coskun 2022 NÃO recomendou a média com sinal (o ADR-064 atribuiu errado).

Estudo do gestor feito em 06/10/2026 (commit e8e5ec6, `_arquitetura/integracao_dbseac/estudos/`). Recomendação provisória ao RT: estimador G (RMS das médias por rodada), sem relaxar plano de CQ automaticamente; B (média com sinal) rejeitado porque cancela rodadas de sinais opostos (MCH: +1,58/−0,52/−1,64%). Com só 3 rodadas, a alternativa conservadora é manter A até ≥ 6 rodadas. A decisão final é do RT.

**D-01 DECIDIDA em 13/10/2026 (ADR-072):** o estimador do bias NÃO muda (BiasEQ ABS atual; G não implementado). O usuário escolhe as rodadas em 3 opções: (1) rodada específica; (2) Ano = um ano + Rodada = TODAS → pooled só daquele ano (exato, sem fallback); (3) Ano = TODOS + Rodada = TODAS → todos os anos. Não reabrir essa discussão.

**Why:** o usuário quer a fonte automática preservada e decisões de Sigma baseadas em evidência medida.
**How to apply:** substituição só pela escolha explícita do ADR-082 (nunca automática por horário). Não mudar o viés do Sigma sem o estudo comparativo e a aprovação do responsável técnico. Ver [[incerteza-nordtest]], [[estado-entrega-0510]].
