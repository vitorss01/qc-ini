---
name: estado-entrega-0510
description: Estado do QC_INI após a sessão em casa (05/10/2026, commit 6275ca8) e passos para ligar a automação no laboratório
metadata:
  type: project
---

Em 05/10/2026 a sessão de casa entregou e enviou à main o commit 6275ca8 (incerteza de medição, auditoria final, gravação trancada, QA-ETL-001; ADR-066). Bateria completa: Hematologia 191/191, Bioquímica 190/190 (10 suítes; `entregar.py` só troca a produção se tudo passar). Este computador do laboratório já fez `git pull` (main = 6275ca8).

DB_SEAC de produção `\SERVIDOR\COMPARTILHAMENTO\INTERFACEAMENTO_DB -  TESTE\DB_SEAC.xlsm` (dois espaços antes de TESTE) levado à v2.4 em 05/10/2026 (backup sha256 em `_backup_pre_v24_20261005` na mesma pasta). Para ligar: `entregar.py ... --modo SEAC --caminho <esse caminho>`; em cada QC, aba Cfg_Integracao (ADM): CAMINHO_DB_SEAC = caminho do laboratório e MODO_FONTE = SEAC (veio HISTORICO); depois ATUALIZAR DADOS e conferir QA_INTEGRACAO (A07 some; E06/E07/A03 = problemas da origem). Bateria real: `python entregar.py Hematologia Bioquimica --modo SEAC`. Em máquina com erro de certificado: `git -c http.sslBackend=schannel pull/push`.

Pendências: RT (D-01 viés do Sigma; manual pode substituir automático excluído?), senha do projeto VBA. Ver [[integracao-db-seac]], [[trabalhar-direto-na-main]].

**Atualização 05/10/2026 ~19h:** os dois QCs foram ENTREGUES em produção em MODO_FONTE=SEAC com o DB_SEAC da rede (Hema sha 2af4fed5, 191/191; Bio sha db53726d, 190/190). Backups em `_backup_pre_integracao_2026-10-05_1705` e `_1851`. Código no commit 566d7ab e evidências no 1c6c219 (push feito). Os .xlsm de produção saíram do git em 06/10/2026 (gitignore + MANIFESTO_PRODUCAO.md com SHA-256; decisão do gestor). Em SEAC, cada ATUALIZAR DADOS leva ~150 s (leitura pela rede), e a entrega completa leva ~5,5 h: rodar destacada via .ps1 (Start-Process -File) para o caminho com dois espaços chegar intacto.

**Atualização 08/10/2026:** os dois QCs estão em produção com ADR-068 a 071 (13 suítes PASS cada): ETp CLIA/VB/FAB (069), Inativar só com ID+Analito+Registrar-LJ+Motivo e ID na dica do LJ (070), seleção de rodadas do CEQ + Sigma com CV pooled de 6 meses (071). Bio entregue em 07/10 22:32, Hema em 08/10 02:12; hashes no MANIFESTO_PRODUCAO.md. A entrega leva ~4 h por produto em SEAC: rodar destacada (.ps1 + Start-Process). Pendentes do RT: D-01, CV alto no N1 da Bio, senha VBA; a senha de proteção de abas está pública no repositório.

**Atualização 13/10/2026 (versão de implantação):** os dois QCs em produção com ADR-068 a 072 (Hema 16:41, Bio 18:55; 13 suítes PASS cada; hashes no MANIFESTO). Entregues também GUIA_DE_USO.md e estudos/CV_INVESTIGACAO_2026-10-13.md (253 resultados suspeitos na Bio para o RT revisar e inativar). Ficam com o laboratório: senha do projeto VBA, troca da senha de abas (pública no repo), revisão dos IDs suspeitos, manter CEQ com DP do grupo e nº de laboratórios.
