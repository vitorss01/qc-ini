---
name: incerteza-nordtest
description: Decisão do usuário (05/10/2026): incerteza de medição só com CIQ + CEQ (Nordtest), sem certificado de calibrador
metadata:
  type: feedback
---

Em 05/10/2026 o usuário decidiu: "ESQUECER CERTIFICADO DE INCERTEZA DE CALIBRADOR — APENAS USAREMOS CONTROLE INTERNO E EXTERNO/EQA". Registrado no ADR-067: uc = raiz(u(Rw)² + u(bias)²), u(bias) = raiz(RMS² + u(Cref)²) do CEQ, com ≥ 6 amostras de ≥ 2 rodadas e u(Cref) obrigatório. Sem u(bias) o U não é calculado (u(Rw) sozinha não é a MU).

**Why:** a literatura dá suporte (Magnusson 2012, Cui 2017, Hann 2017, Michaelis 2026), não depende de documento do fabricante e pode ser automatizada com dados que o QC_INI já tem.
**How to apply:** nunca reintroduzir u(cal) nem cobrar certificados de calibrador; triagem do viés continua como monitoramento. Ver [[integracao-db-seac]], [[estado-entrega-0510]].
