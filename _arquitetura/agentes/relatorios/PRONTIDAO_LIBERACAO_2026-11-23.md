# Prontidão para liberar o QC_INI na rotina (Bioquímica e Hematologia), 23/11/2026

Análise só de leitura, sem Excel e sem COM. O critério é a **segurança do resultado do CQ**: estatística, gráfico,
Westgard, perda de dado e segurança real. Perfeição não entra no critério.

Fontes:
- `OCORRENCIAS.md`: 31 ocorrências com status diferente de VALIDADA ou ENCERRADA;
- `continuidade/README.md`, `CONTINUAR_EM_CASA.md` e `STATUS_ENTREGA_39.md`;
- ADR-093 a ADR-100;
- os relatórios de evidência citados: `parecer_cq`/`parecer_ux` do ciclo01, reteste03, reteste04, entrega39 e reteste08;
- o código do `aplicar_alvos_ciq.py`, do `mMaterialLote.bas` e do `Configuracao.cls`.

Legenda:
- **A**: bloqueia a liberação;
- **B**: libera com uma orientação ao usuário;
- **C**: depende do laboratório;
- **D**: já está resolvida; falta formalizar.

## 1. Veredito em uma linha

**Nenhuma ocorrência registrada bloqueia (0 em A).** Todas as Críticas e Altas do código estão VALIDADAS desde a
entrega 45. O que condiciona a liberação são três itens **fora do registro**:
1. o ADR-100 aplicado com revisão do laboratório;
2. a entrega do ADR-099 passar no gate na máquina do trabalho, ou ser adiada;
3. a regressão da Hematologia, que ficou incompleta: faltam 10 suítes.

## 2. Ocorrências abertas

| OC | Prior. | Resumo | Classe | O que falta / orientação / ação |
|---|---|---|---|---|
| 0007 | Alto | Lotes de urina/PCR ativos em analito sérico | **D** | Prova: Tester no reteste03 (`tester_validar.md` l.250, "0 resultados … participando em analito do catálogo") e ADR-091 RESOLVIDA no reteste04 (43/43). Marcar VALIDADA. |
| 0008 | Alto | Alvos do lote 522611 (Hema) incompatíveis: IG# a −30 DP, RDW-SD a +398 DP | **C → D** | O lote venceu em 31/03/2026, e a Estatística e o Sigma não usam o alvo (parecer G6). O erro está no LJ e no Westgard desse lote. O ADR-100 substitui esses alvos pela média do laboratório. Quando ele for aplicado, o critério G6-1 (nenhum par com \|z\| > 5) pode ser conferido no CSV. A conferência da bula fica opcional. |
| 0010 | Médio | Vários "Sigmas"; o bloco da aba Analitos usa o CV do cadastro; Sigma legado no Eng_Saida | **B** | "O Sigma oficial é o da Estatística (coluna L) e o do Painel (I7). Ignore o Sigma dos blocos da aba Analitos." Depois do ADR-100, esse bloco muda (CV dos 40 primeiros) e fica ainda mais parecido com um número oficial: rotular na próxima entrega. |
| 0011 | Médio | Biblioteca de especificações 2026 vazia na Hema; Lipase VB e PCR CLIA na Bio | **C** (urgente) | Lote com ano de referência 2026 perde o ETp, o Sigma e o plano. Na Hema são **todos** os analitos. Não mostra número errado: mostra "-". O laboratório preenche 2026 antes do 1º lote com esse ano. **Conferir agora** se algum lote já está com 2026. |
| 0014 | Baixo | 90 linhas "Controllab" da simulação na tabela do CAP | **B** | Uso = NAO, fora de qualquer conta. Orientação: "nunca troque o Uso dessas linhas para SIM". |
| 0015 | Médio | RDW: CEQ no RDW-CV e VB no RDW-SD; nenhum dos dois tem Sigma | **C** | Conferir na fonte (EFLM) a que medida se referem o CVi e o CVg e mover a especificação, se for o caso. Nenhum número errado. |
| 0016 | Médio | ETp "CLIA" de WBC 7 % e PLT 15 % (o CLIA vigente é 10 / 25) | **C** | Conferir no 42 CFR 493.941. Hoje o erro é conservador: Sigma subestimado e plano de CQ mais pesado, não mais frouxo. |
| 0017 | Baixo | "FAB CVTp" da Hema com valores que não são CV | **C** | Esvaziar X53:AA53 antes de escolher a fonte FAB. Nenhum analito usa FAB hoje. |
| 0020 | Médio | Manual **colado** com analito, lote ou equipamento fora do cadastro fica ATIVO e não aparece no Painel | **B** | "Digite ou escolha na seta analito, lote e equipamento; não cole. Depois do ATUALIZAR, confira se o ponto apareceu no LJ." É a lacuna de produto mais relevante entre as Médias: uma omissão silenciosa de uma corrida. Corrigir na próxima entrega (A11 + destaque), sem autorização. |
| 0022 | Baixo | O I01 lista analitos de urina sem dizer a matriz | **B** | "Os códigos NA/K/CL/CRE2/PHOS no I01 são da urina." |
| 0025 | Médio | Status REJEITADO = pior corrida do período, sem dizer isso | **B** | "O status é o pior do período escolhido. Para a corrida de hoje, olhe o último ponto do LJ ou escolha o período de hoje." O ADR-100 reduz os REJEITADO falsos. |
| 0026 | Baixo | MOTIVO "Outro" sem complemento | **B** | "Ao escolher 'Outro', registre a explicação no comentário do ponto ou na NC." |
| 0037 | Médio | Com o H3 vazio, a Estatística mostra o lote em uso e diz "refletido do Painel" | **B** | "Com o Painel em todos os lotes, a aba Estatística continua mostrando o lote em uso." |
| 0038 | Médio | O período não volta ao retornar de outro lote | **B** | "Ao voltar a um lote, confira e reescolha o período." |
| 0039 | Médio | C4/A39 mostram o lote em uso (vencido na Hema) | **C + B** | Laboratório: atualizar o lote em uso (Configuração C20) nos dois produtos. Em 31/10 a Hema estava no 522611, vencido. Usuário: "C4 é o lote da digitação, não o da tela". |
| 0045 | Médio | Lista do H3 com lotes de todos os analitos; escolher um troca o analito | **B** | "Escolha no H3 um lote do analito que está na tela; se o analito mudar sozinho, volte." |
| 0046 | Médio | "Últ. violação" e "Classificação" são do histórico do lote, não do período | **B** | "Essas duas colunas são do lote inteiro; as violações do período são as da esquerda." |
| 0047 | Baixo | Ponto de violação cortado no topo do LJ | **B** | Só visual; o ponto existe e o evento aparece na tabela de violações. |
| 0053 | Médio | Bias digitado sobre a fórmula (P) no EQA | **D** | Prova: Tester VALIDADA no reteste03 (`tester_validar.md` l.27) e Usuário Simulado RESOLVIDA de carona no reteste04 (`Bioquimica_F1.md` l.51, `Hematologia_F1.md` l.46). Marcar VALIDADA. |
| 0060 | Médio | Inativar leva de 20 a 27 s; teto provisório até 30/11 | **B** | "Inativar demora cerca de 30 s. Espere a janela fechar; não clique de novo nem feche o Excel." O UX aprovou como provisório até 30/11 (`entrega39/parecer_ux.md`). |
| 0061 | Baixo | Lote só de urina abre o Painel vazio, sem explicar | **B** | "Lotes de urina não têm gráfico no Painel (decisão do ADR-091)." |
| 0062 | Baixo | Teste com dia escolhido sobre cópia defasada | **B** (só desenvolvimento) | Não afeta o produto. |
| 0063 | Baixo | O "aguarde N s" é fixo e não bate com o tempo real | **B** | Mesma orientação da OC-0060. |
| 0064 | Baixo | O GUIA diz "etapa k de 7"; o código mostra 8 | **B** | Corrigir uma palavra no GUIA, sem impacto. |
| 0065 | Médio | A tarefa DB_SEAC abre um Excel a cada 2 min e deixa a tela mais lenta em 35 a 45 % do tempo | **B** | "Lentidão de 1 a 2 s em alguns momentos é a atualização automática; não é erro." É risco residual (§5). |
| 0067 | Médio | Senha das abas em texto puro no repositório público | **B** (ação de segurança fora do gate) | A barreira real é o login PBKDF2 (ADR-084) mais a permissão da pasta. Mitigar já: repositório privado e pasta de produção com escrita só para a equipe. Trocar a senha numa entrega própria, em até 2 semanas. Não bloqueia porque a proteção de aba do Excel já é quebrável sem a senha. |
| 0069 | Médio | Nome do Controllab mapeado com o provedor CAP diz "entra no bias" e não entra | **B** | Latente (mapas do Controllab vazios). "Um CEQ do Controllab só entra no bias se o provedor do analito for Controllab (Bio: Especificações AR; Hema: Estatística R4)." |
| 0070 | Médio | Lote FORA_DO_CQ/SO_ANALITOS mostra SEM CORRIDA e pede Média/DP | **B** | Na Bio não é latente (8794 e 8353 no H3). "Nesses lotes, não cadastre Média/DP; SEM CORRIDA é esperado." O número está certo. |
| 0077 | Baixo | Bloco por ano no Eng_Saida que ninguém lê | **B** (informativo) | Sem efeito no que é exibido. |
| 0085 | Baixo | Fixture do R06/R09 no qa_adr097 | **D** | Prova: a variante do próprio Tester deu 29/0 (`reteste08/tester_validar.md` l.113-118) e o Orquestrador rodou a suíte corrigida com 29/0. Formalizar quando o gate da entrega 46 passar com o R06 aprovado. |
| 0087 | Baixo | Falta o teste automático (R10) do retrato tirado na seleção | **D** (o produto está validado) | O Tester exercitou o comportamento à mão no reteste08, e a OC-0084 está VALIDADA. Falta só automatizar o R10. Não afeta a liberação. |

**Totais:**

| Classe | Qtde | Ocorrências |
|---|---|---|
| A | 0 | — |
| B | 21 | OC-0010, 0014, 0020, 0022, 0025, 0026, 0037, 0038, 0045, 0046, 0047, 0060–0065, 0067, 0069, 0070, 0077 |
| C | 5 | OC-0011, 0015, 0016, 0017 e 0039 (esta também com orientação) |
| D | 5 | OC-0007, 0053, 0085, 0087 e 0008 (esta depois do ADR-100) |

## 3. Itens fora do registro que condicionam a liberação

| Item | Por que pesa | Condição |
|---|---|---|
| **ADR-100: alvos pelos 40 primeiros** (ferramenta, ainda não aplicada) | É o item de maior ganho de segurança. Com alvos de bula ou TESTE deslocados, o Painel grita REJEITADO sem motivo, e na rotina o usuário aprende a ignorar o alarme. Isso é risco real de CQ. | Aplicar **depois** de o RT revisar o CSV do `--simular` (critérios no §4, passo 6). |
| **ADR-099: Material do lote + avisos do CEQ** (testado em casa, modo HISTÓRICO, sem DB_SEAC) | Mexe na entrada dos motores (`tblLotesMaterial`, `mIntegracao` e o espelho no ATUALIZAR). Não é requisito de segurança: os avisos do CEQ ajudam, mas nenhum cálculo muda. | Entra só com o gate completo verde no trabalho. Se não passar a tempo, **liberar na entrega 45 + ADR-100** e deixar o ADR-099 para a semana seguinte. |
| **Regressão da Hema: 10 suítes não rodadas** (`qa_westgard_sequencia`, `adr084`, `usuarios_adr084`, `adr085`, `adr086`, `adr088`, `adr089`, `adr091`, `adr094`; a sessão caiu) | O `qa_westgard_sequencia` é justamente o Westgard que o ADR-100 vai recalcular. | Rodar no mínimo `qa_westgard_sequencia`, `qa_adr091` e `qa_adr094` no gate da entrega 46. |

Críticas aos ADRs, para a revisão do laboratório:

**ADR-100:**
- **DP dos 40 primeiros.** Se as 40 corridas couberem em poucos dias, o DP sai subestimado e aparecem mais 1_2s/1_3s
  falsos. A CLSI C24 pede pelo menos 20 dias e depois o DP acumulado. Usar no CSV a razão entre o DP novo e o DP do
  período inteiro, e o intervalo de datas.
- **Sem exclusão de pontos.** Corridas REJEITADAS e não inativadas entram na média.
- **Alvo único para DIMENSION 1 e 2 na Bio.** Infla o DP e mascara o viés entre os aparelhos; o CSV marca quando a
  diferença passa de 1 DP.
- **Retroatividade.** Eventos de Westgard antigos somem ou mudam, auditados como RETROATIVO; as NCs já registradas
  ficam no histórico.
- **Sem ajuste por par.** A ferramenta é tudo ou nada por produto. Um par discordante se corrige depois, pela aba
  Analitos, com trilha.

**ADR-099:**
- a coluna Material altera o que entra na estatística, com prévia, confirmação e trilha, mas **sem checagem de papel**
  em `mMaterialLote`/`Configuracao.cls`: vale a proteção da aba Configuração. Confirmar que só SUPERVISOR/ADMIN
  editam F26:F125;
- LÍQUOR e SANGUE ficaram fora da lista da Bio (decisão do laboratório).

**Repositório:**
- o `.gitignore` local já cobre `_backup_pre_alvos_*/` e `alvos_ciq_*`, mas a mudança **não está commitada**. Como o
  backup é gravado ao lado da produção, na raiz do repositório público, isso precisa estar no GitHub antes de rodar a
  ferramenta no trabalho;
- o cabeçalho do script ainda diz que o `.gitignore` não cobre esses arquivos: o texto está desatualizado.

## 4. Checklist mínimo, em ordem

| # | Quem | Passo |
|---|---|---|
| 1 | Desenvolvimento (casa) | Commit e push do ADR-099, do ADR-100, dos testes novos (`qa_adr099`, `qa_fluxo_manual`, `qa_fluxo_ceq`, `qa_blocos_lote` BN/BO) e do `.gitignore`. Conferir o diff: sem senha, sem caminho do DB_SEAC, sem `.xlsm`, sem `~$*`. Rodar o lint VBA. |
| 2 | Desenvolvimento (trabalho) | `git pull`. Fechar outros Excel. Pausar a tarefa DB_SEAC durante o gate. Avisar que ninguém deve usar o QC. |
| 3 | Desenvolvimento | Entrega 46 (ADR-099) pelo `entregar.py`, Hema e depois Bio, com as suítes direcionadas mais `qa_adr099`, `qa_fluxo_manual`, `qa_fluxo_ceq`, `qa_blocos_lote`, `qa_adr097` (R06), `qa_westgard_sequencia`, `qa_adr091` e `qa_adr094`. Queda por FailFast do Power Query: rodar a suíte sozinha de novo, nunca trocar a produção com suíte pendente. **Se não ficar verde no dia, seguir sem o ADR-099.** |
| 4 | Desenvolvimento | `aplicar_atualizar.py` nos dois. O instalador do ADR-099 lista os lotes sem regra de material: guardar a lista para o passo 8. |
| 5 | Desenvolvimento | `aplicar_alvos_ciq.py <produto> --simular` nos dois. Se for aplicar sobre a entrega 45, sem o ADR-099, rodar antes `--so-copia` para provar que a ferramenta funciona nessa versão. |
| 6 | **Laboratório (RT / Especialista em CQ)** | Revisar o CSV e aprovar por escrito. Marcar para investigar ou corrigir depois: (a) deslocamento novo × bula acima de 2 DP da bula (pode ser viés real do método, não alvo errado); (b) DP novo / DP do período inteiro abaixo de ~0,7, ou 40 primeiros em menos de ~10 dias (DP curto); (c) marca D1×D2 > 1 DP na Bio; (d) % fora de 3 DP depois acima de ~1 %; (e) 522611 IG#/RDW-SD com \|z\| < 5. |
| 7 | Desenvolvimento | Aplicar sem flag nos dois. A ferramenta já faz A00–A11, o backup, o sha e o ATUALIZAR só-entradas. Guardar o CSV e o JSON fora do repositório. |
| 8 | Laboratório | (a) Lote em uso atualizado (Configuração C20) na Hema e na Bio. (b) Revisar o Material dos lotes listados no passo 4. (c) Conferir se algum lote tem ano de referência 2026; se tiver, preencher as especificações 2026 (OC-0011). (d) Ano EP (Estatística N4) = 2026 quando entrar o 1º CEQ de 2026. (e) Login: o primeiro acesso do ADMIN no cadastro central ainda estava pendente em 29/10. Confirmar que cada analista tem login e assinatura. |
| 9 | Tester (independente) | Validar o ADR-099 (M01–M09 e avisos do CEQ) e o ADR-100: o BO do `qa_blocos_lote` recalcula a média e o DP; o JSON dos eventos antes e depois bate. Na mesma rodada, formalizar a OC-0085. O Orquestrador marca VALIDADAS as OC-0007 e 0053 com as provas do §2. |
| 10 | Usuário (com o desenvolvimento) | Entregar à equipe uma folha de orientação de 1 página com as linhas "B" do §2, em especial as OC-0020, 0025, 0060, 0046 e 0070. |

## 5. Riscos residuais conhecidos

1. **Excel/Power Query FailFast (0x80131623).** É ambiental e intermitente. No uso: se o Excel cair no ATUALIZAR,
   reabrir sem salvar nada e rodar o ATUALIZAR de novo. Na entrega: rodar a suíte de novo, nunca trocar a produção com
   suíte pendente.
2. **Desempenho.** Inativar leva cerca de 20 a 30 s (OC-0060, teto provisório até 30/11), e a tarefa DB_SEAC deixa a
   máquina mais lenta em boa parte do tempo (OC-0065). Afeta a usabilidade, não o resultado.
3. **Regressão incompleta da Hema** se o passo 3 não incluir as suítes pendentes.
4. **ADR-100.** O DP de curto prazo, o alvo único D1+D2 e a mudança retroativa do Westgard (§3). Mitigação: a revisão
   do passo 6 e a correção por par na aba Analitos, com trilha.
5. **Omissão silenciosa de um manual colado fora do cadastro (OC-0020).** Mitigação só por orientação até a próxima
   entrega.
6. **Segurança.** A senha das abas é pública (OC-0067). A barreira real é o login, a permissão da pasta e o Audit_Log.
7. **Cadastro.** Especificações 2026, o lote em uso e as fontes CLIA/VB dependem do laboratório (OC-0011, 0015, 0016 e
   0039). Hoje faltam números (aparece "-"), mas nenhum número errado é mostrado, salvo o Sigma conservador da
   OC-0016.
8. **Liberação/assinatura.** O login e a assinatura do ADR-084 dependem do cadastro central estar completo (passo 8e).
   Sem login não há trilha de autoria confiável.
9. **Arquivo único.** O `.xlsm` fica na pasta compartilhada. Manter backup diário (versões do OneDrive mais os
   `_backup_*` das entregas) e orientar a equipe a ter um usuário editando por vez. Esse comportamento não foi
   verificado nesta análise.

## 6. Recomendação final

**Liberar: SIM, nesta semana, com quatro condições:**
1. o **ADR-100 aplicado depois da aprovação escrita do RT** sobre o CSV (passos 5 a 7);
2. o **lote em uso atualizado** e os **logins confirmados** (passo 8a e 8e);
3. a **folha de orientação** entregue (passo 10);
4. a entrega 46 (ADR-099) **só com o gate verde**, incluindo as suítes pendentes da Hema. Se o gate não fechar a
   tempo, liberar na **entrega 45, já validada, + ADR-100**: o ADR-099 é melhoria e não pré-requisito de segurança.

As OC classe C seguem em paralelo, com o laboratório. A OC-0011 é urgente se algum lote já estiver com 2026. As
melhorias B de produto entram nas próximas entregas, nesta ordem: OC-0020 (A11), OC-0070, OC-0025/0046 (rótulos) e
OC-0010 (rótulo do bloco). A OC-0067 (troca da senha) fica para até 2 semanas.
