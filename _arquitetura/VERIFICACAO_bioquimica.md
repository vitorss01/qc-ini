# Verificacao do QC_INI - Bioquimica

**Executada em:** 04/08/2026 20:45:48
**Artefato:** `C:\Users\vitor\QCINI_build_hardening1_Bioquimica\QC_Bioquimica_verificacao.xlsm`

| Item | Verificacao | Resultado | Evidencia |
|---|---|---|---|
| 0.1 | build_all.ps1 concluiu | OK |  |
| 1.1 | mEstatistica identico a fonte | OK | fonte 5bcec6951643 / artefato 5bcec6951643 |
| 1.1 | mDados identico a fonte | OK | fonte 9e2366bd42aa / artefato 9e2366bd42aa |
| 1.1 | mAuditoria identico a fonte | OK | fonte a24fbb7c2ffe / artefato a24fbb7c2ffe |
| 1.4 | formulario frmResultadoNaoConforme presente | OK |  |
| 1.4 | formulario frmExcluirRegistroNC presente | OK |  |
| 1.5 | botao btnNaoConforme na aba Resultados | OK | OnAction = 'AbrirFormNaoConforme' |
| 1.5 | botao btnExcluirNC na aba Registros | OK | OnAction = 'AbrirFormExcluirRegistroNC' |
| 1.2 | nenhum identificador aS/As no projeto | OK | ocorrencias: 0 |
| 1.3 | aba Eng_Saida existe | OK |  |
| 1.3 | aba Corridas existe | OK |  |
| 1.3 | aba Audit_Log existe | OK |  |
| 2.1 | AtualizarCalc | OK | 0,03s |
| 2.1 | AtualizarPainelEng | OK | 0,03s |
| 2.1 | AtualizarEstatisticaAba | OK | 0,05s |
| 2.1 | RegistrarEventosWestgard | OK | 0,03s |
| 2.2 | Eng_Saida publica chave ANALITO|RUN | OK | AB3 = 'Glicose|1' |
| 3.1 | parecer curto recusado, parecer valido aceito | OK | 3 palavras=False / 7 palavras=True |
| 3.2 | log grava e cresce | OK | linhas: 3 -> 6 |
| 3.3 | cadeia integra apos gravacao | OK | retorno: OK|3 |
| 3.4 | ADULTERACAO DETECTADA na linha certa | OK | retorno: QUEBRADO|5|conteudo alterado |
| 3.5 | cadeia volta a fechar apos restaurar | OK | retorno: OK|3 |
| 3.6 | reenvio NAO ressuscita registro excluido | OK | status antes=Excluido depois=Excluido | UpsertResultados devolveu '0|0|1' (novos|atualizados|bloqueados) |
| 3.7 | tentativa de reenvio fica registrada no log | OK | acao registrada: 'REENVIO_BLOQUEADO' |
| 3.10 | exclusao grava tambem em LOG_Resultados | OK | linhas no log do banco: 0 -> 1 (excluidos: 1) |
| 3.11 | as duas camadas compartilham o ID_Auditoria | OK | banco='AUD-000005-20260804204208' evento='AUD-000005-20260804204208' |
| 3.12 | parecer curto impede marcar nao conforme | OK | retorno vazio, status permaneceu 'Ativo' |
| 3.13 | resultado sai dos calculos e o valor ORIGINAL fica | OK | status Ativo -> Invalido | valor preservado: 92.1383 |
| 3.14 | ocorrencia aparece na aba Registros | OK | linha 6: analito 'Glicose', RUN 7 |
| 3.15 | as tres camadas registram com o mesmo ID | OK | ID: 'AUD-000006-20260804204211' | log do banco: 1 -> 2 |
| 3.8 | alterar elegibilidade sobe a versao da configuracao | OK | versao 1 -> 2 |
| 3.9 | alteracao de elegibilidade vai para a auditoria | OK | acao registrada: 'CFG_ELEGIBILIDADE_ALTERADA' |
| 4.1 | AUSENTE so na lista fechada da Sprint NC | OK | 720 AUSENTE, 0 fora da lista fechada (regRep2/regRep3 do Calc) |
| 4.2 | nenhuma formula virou VALOR | OK | VALOR     0 |
| 4.3 | nenhuma formula de interface calcula sobre o banco | OK |  |
| 4b.1 | motor completo abaixo de 5s | OK | 0,29s para as 4 rotinas sobre 1.575 registros |
| 4b.2 | gravacao em camada dupla abaixo de 250ms por evento | OK | 107ms por evento (5,3s para 50 eventos nas duas camadas) |
| 4b.3 | verificacao da cadeia integra e rapida | OK | OK|59 em 0,25s |
| 2.2 | valor ORIGINAL recuperavel no Audit_Log | OK | OK|92,0028 |
| 1.6 | nenhum nome publico duplicado entre modulos | OK | 107 nomes publicos, nenhum repetido |
| 4.6 | nenhum erro de formula no artefato | OK | #N/A excluido de proposito (lacuna de serie do grafico); achados: nenhum |
| 4.7 | estado global restaurado apos o motor | OK | ScreenUpdating=True Calculation=xlCalculationAutomatic (esperado True / -4105) |
| 4.8 | idempotencia: 1x e 11x produzem saida identica | OK | 1x b590e10010fc / 11x b590e10010fc |
| 4.9 | persistencia: reabrir e recalcular nao muda a saida | OK | reaberto b590e10010fc / recalculado b590e10010fc |
| 5.1 | so Login e a trilha ficam visiveis no arquivo salvo | OK | veryHidden 19 de 22 (visiveis: Login, Audit_Log, Audit_Legenda) |
| 5.5 | trilha de auditoria legivel com macros desabilitadas | OK | Audit_Log nao esta oculta: e material de auditoria, feito para ser lido |
| 5.2 | estrutura da pasta protegida | OK |  |
| 5.3 | toda aba tem <sheetProtection> GRAVADA no arquivo | OK | 22 de 22 abas |
| 5.4 | projeto VBA com senha (item 3.4) | **FALHA** | PASSO MANUAL PENDENTE: VBE > Ferramentas > Propriedades do VBAProject > Protecao |

**48 de 49 verificacoes passaram.**

## Falhas
- **5.4** projeto VBA com senha (item 3.4) -- PASSO MANUAL PENDENTE: VBE > Ferramentas > Propriedades do VBAProject > Protecao
