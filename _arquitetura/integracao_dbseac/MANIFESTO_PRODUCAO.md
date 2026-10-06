# Manifesto de produção — QC_INI

Os `.xlsm` de produção não são versionados (contêm dados do laboratório e configuração interna; o remoto é
público — decisão do gestor em `estudos/DECISAO_GESTOR_2026-10-06.md`). A versão em uso é identificada pelo SHA-256.

| Data da entrega | Arquivo | Bytes | SHA-256 | Código (commit) | Bateria |
|---|---|---|---|---|---|
| 05/10/2026 | QC_Hematologia.xlsm | 13167580 | `2af4fed57df9846dfe0cea92ebd71eba252d0f41439efbb8f834652c58e99fc9` | 566d7ab | 191/191 PASS, MODO_FONTE=SEAC |
| 05/10/2026 | QC_Bioquimica.xlsm | 15348763 | `db53726deb3fe515d981ba0b34e5faabc863f49ed66a914778ec4d8ea7a0a390` | 566d7ab | 190/190 PASS, MODO_FONTE=SEAC |
| 06/10/2026 | QC_Hematologia.xlsm | 13182562 | `ca43b0005e5cca1bafa88bdac0772b7d406b6ce49d3f8c6b104270e7f3eda576` | c657456 | 10 suítes PASS (incerteza 15/15), MODO_FONTE=SEAC — ADR-068 e A09 |

O hash muda a cada gravação no laboratório (uso normal). Para conferir a versão entregue, compare com o arquivo
`_entrega_<data>/<produto>/<arquivo>.xlsm` da pasta de entrega (não versionada).
