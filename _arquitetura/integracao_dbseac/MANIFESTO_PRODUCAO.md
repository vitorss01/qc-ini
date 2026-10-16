# Manifesto de produção — QC_INI

Os `.xlsm` de produção não são versionados (contêm dados do laboratório e configuração interna; o remoto é
público — decisão do gestor em `estudos/DECISAO_GESTOR_2026-10-06.md`). A versão em uso é identificada pelo SHA-256.

| Data da entrega | Arquivo | Bytes | SHA-256 | Código (commit) | Bateria |
|---|---|---|---|---|---|
| 05/10/2026 | QC_Hematologia.xlsm | 13167580 | `2af4fed57df9846dfe0cea92ebd71eba252d0f41439efbb8f834652c58e99fc9` | 566d7ab | 191/191 PASS, MODO_FONTE=SEAC |
| 05/10/2026 | QC_Bioquimica.xlsm | 15348763 | `db53726deb3fe515d981ba0b34e5faabc863f49ed66a914778ec4d8ea7a0a390` | 566d7ab | 190/190 PASS, MODO_FONTE=SEAC |
| 06/10/2026 | QC_Hematologia.xlsm | 13182562 | `ca43b0005e5cca1bafa88bdac0772b7d406b6ce49d3f8c6b104270e7f3eda576` | c657456 | 10 suítes PASS (incerteza 15/15), MODO_FONTE=SEAC — ADR-068 e A09 |
| 07/10/2026 | QC_Bioquimica.xlsm | 15473750 | `952dc3af7e009a180cd1cdfc648cc844c10ebc699eb4a05cf985c587df3699b9` | 94c4791 | 13 suítes PASS, MODO_FONTE=SEAC — ADR-069, 070, 071 |
| 08/10/2026 | QC_Hematologia.xlsm | 13271260 | `1505bd5380471ea845b322a81f08cae6cbf2421e30b47762fdad87192110ce7b` | eff6af7 | 13 suítes PASS, MODO_FONTE=SEAC — ADR-070, 071 (+ aquecimento da janela do Sigma) |
| 13/10/2026 | QC_Hematologia.xlsm | 13289510 | `a1dd5c25317f4ddcc6dc352ce61f397ef439a5ae6eabe31a7d5c31b388bb0c6b` | 5db9b6e | 13 suítes PASS, MODO_FONTE=SEAC — ADR-072 (D-01 em 3 opções) |
| 13/10/2026 | QC_Bioquimica.xlsm | 15498129 | `37d7d97c8b9010ede8a6ccf395ee8f4dccc2949090235cc2b1ddb96e11b11853` | e5f2575 | 13 suítes PASS, MODO_FONTE=SEAC — ADR-072 (D-01 em 3 opções) |
| 14/10/2026 | QC_Hematologia.xlsm | 13332621 | `11a722f77b422a033b70bb7dfb6487b2b97e51d2e04db6d96dac44cbaa253c60` | e299d3e | 14 suítes PASS, MODO_FONTE=SEAC — ADR-073 (limites por lote + rastreabilidade) |
| 14/10/2026 | QC_Bioquimica.xlsm | 16411797 | `fa313eef9ebc1a89a52082b156851a912290fbd6fa99c81373f0e11e50639bad` | 137864b | 15 suítes PASS, MODO_FONTE=SEAC — ADR-073 e ADR-074 (aba Analitos com um bloco por lote; motor na aba Especificações) |
| 15/10/2026 | QC_Hematologia.xlsm | 13764507 | `e2d8783ce728044b5f6a7a6825b9568828442397ea5448ca7463cbc3c6d9e8e1` | 031fa04 | 15 suítes PASS, MODO_FONTE=SEAC — ADR-075 (especificações no modelo da Bioquímica) e ADR-074 (blocos por lote, 3 níveis) |
| 16/10/2026 | QC_Bioquimica.xlsm | 16460640 | `afe6c99794d2b0801594c206bc3db68a5b84849224938235c129868f7e9b833c` | de6b36a | 17 suítes PASS, MODO_FONTE=SEAC — ADR-076/077 (valores de TESTE, Painel acompanha o lote) e ADR-078 (linhas do alvo sem corridas; troca de lote não grava) |
| 16/10/2026 | QC_Hematologia.xlsm | 13940293 | `e219d0ee2e3a1d1de72e1428ca81d287df50a12fcb2c2fe315b3dac2b7597619` | de6b36a | 17 suítes PASS, MODO_FONTE=SEAC — ADR-076/077/078 |

O hash muda a cada gravação no laboratório (uso normal). Para conferir a versão entregue, compare com o arquivo
`_entrega_<data>/<produto>/<arquivo>.xlsm` da pasta de entrega (não versionada).
