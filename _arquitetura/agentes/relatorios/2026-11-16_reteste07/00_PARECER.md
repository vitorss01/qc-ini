# Parecer consolidado: ciclo 2026-11-16_reteste07

Agente: Usuário Simulado (parecer consolidado). Data: 17/11/2026. Modo: reteste, consolidação.

**De onde vêm os dados.**
- Neste passo o Excel não foi aberto.
- Tudo vem dos 2 relatórios das frentes desta pasta (`Bioquimica_F1.md`, `Hematologia_F1.md`) e do resumo estruturado
  que cada frente devolveu.
- Para conferir o Relatório B, li só para leitura os JSON das frentes na pasta do ciclo no `%TEMP%`
  (`Bioquimica_F1_nova.json`, `Hematologia_F1_nova.json`, blocos `antes` e `depois2`). Recalculei:
  - a diferença relativa oráculo × `engPainel` em todos os níveis e momentos;
  - à mão, as somas "antes − retirados − substituído + manuais ATIVOS + reais novos";
  - os z do Westgard;
  - a aritmética de bias, ET e Sigma das linhas do Painel (B.3). É conferência aritmética, **não** validação do
    método.

| Item | Valor |
|---|---|
| Produção auditada (entrega 44, ADR-097) | `QC_Bioquimica.xlsm` sha256 curto **fdc173906fff**; `QC_Hematologia.xlsm` sha256 curto **5b0e78ace7b9**. Nas duas frentes, o hash foi igual antes e depois de todas as execuções (`origem_intacta.igual = true`): a produção não foi aberta para escrita |
| Base de comparação (entrega 43) | Bio: `_backup_pre_integracao_2026-11-16_0729`, sha256 curto **bd5314426ab8**. Hema: `_backup_pre_integracao_2026-11-16_0959`, sha256 curto **49f68fba99d5**. Nas duas, hash igual antes e depois |
| Lista do reteste | 1 item nos dois produtos: **OC-0078** (Alto). Apagar ou recortar uma linha manual deixava USUARIO e REGISTRADO_EM, e o próximo manual digitado nela herdava o carimbo: era recusado com "DATA/HORA posterior ao lançamento" ou ficava no nome de quem apagou. Tratada no ADR-097 §1 |
| Frentes executadas | Só a F1, nos dois produtos, em modo reteste, com os fluxos vizinhos. **Não fizeram parte deste reteste:** a F2 (CEQ), a F3 (bias, ET e Sigma) e a F4 completa. As outras OC da entrega 44 em AGUARDA RETESTE (OC-0048, OC-0079, OC-0080, OC-0081, OC-0082, OC-0083) **não** foram retestadas neste ciclo |
| Caminho | O do usuário: login (`mSeguranca.DoLogin`); filtros do Painel escritos com os eventos ligados; botão Digitar Resultados (`mIntegracao.IrDigitarResultados`); digitação célula a célula (`FormulaLocal`) com `Application.Interactive = True`, para as caixas do `EntradaMudou` aparecerem (lidas pelo texto da janela do PID da frente e fechadas no OK); ATUALIZAR DADOS pelo núcleo do botão (`AtualizarDadosAutomatico`), duas vezes. Na Hema, o Ano De/Até da aba Estatística (B3/D3) foi posto em 2026 como o usuário faria |
| Fora do caminho do usuário (declarado) | Teclado e mouse físicos não foram usados. O Delete foi emulado com `ClearContents`, o recorte com `Range.Cut(destino)` e o Ctrl+D com `Range.FillDown()`. O cenário "dias depois" foi reproduzido com 2,5 min entre o carimbo antigo e a nova digitação (o mecanismo é o mesmo). O carimbo órfão legado foi gravado de propósito com os eventos desligados |
| Dados de teste | Bio: Glicose, DIMENSION 1, lote 8977, 16/11/2026, N1 e N2. Hema: WBC, XN1000, lote 625311, 17/11/2026, N1 a N3. Valores sintéticos de 3 decimais com \|z\| ≤ 0,3 pelo alvo do lote; o manual que substitui vale o real + 0,123 (Bio 77,123; Hema 3,273). MOTIVO "Outro". Usuários `TESTE_AUDITORIA_R7`/`_R7B` (Bio) e `_R7H`/`_R7HB`/`_R7HC`/`_R7HD` (Hema), ANALISTA, criados só nas cópias |
| Cópias | Em `%TEMP%\qcw\auditoria\2026-11-16_reteste07\`: Bio `nova` e `base`; Hema `nova`, `base` e `c1` (complemento). Nenhuma salva, todas apagadas no fim. Ficam os JSON e os logs. Uma primeira execução da Hema, em 16/11, parou no 1º ATUALIZAR sem erro gravado. Ela foi arquivada em `anterior_interrompida\` e refeita do zero numa cópia nova |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`. Bio: PIDs 30228 (nova) e 12976 (base). Hema: 3472 (nova), 5700 (base) e 21728 (c1). O PID da execução interrompida (27492) já não existia. Os outros Excel da máquina (o do usuário e o da tarefa agendada) não foram tocados. Nenhuma captura de tela |
| Ferramentas | Em `_arquitetura/agentes/ferramentas/`, documentadas no cabeçalho: `reteste07_usuario_f1.py <nova\|base>` (Bio), `reteste07_usuario_f1_hema.py <nova\|base>` e `reteste07_usuario_f1_hema_c1.py` (Hema) |
| tblCQ_Final | Bio 36.056 linhas no início: o 1º ATUALIZAR trouxe 2 resultados novos do DB_SEAC, de outros analitos (nenhum no filtro). Hema 48.484 linhas no início: o 1º ATUALIZAR trouxe 155 resultados novos (16/11 à tarde e 17/11 até 08:10:26), dos quais 5 no filtro WBC/XN1000/625311, todos contados na conta à mão |

**Resumo em números**

| | Bioquímica | Hematologia | Total |
|---|---|---|---|
| Item do reteste: RESOLVIDA / PERSISTE / EFEITO_COLATERAL | 1 / 0 / 0 | 1 / 0 / 0 | OC-0078 **RESOLVIDA** nos dois produtos, com uma ressalva no critério 3 (destino do recorte) |
| Testes (T01..T12, com T03a/T03b): APROVADO / REPROVADO / NAO_EXECUTADO | 11 / 1 (T03b) / 1 (T12) | 12 / 1 (T03b) / 0 | 23 / 2 / 1 |
| REPROVADOs que já existiam na entrega 43 (base) | 1 de 1 | 1 de 1 | Nenhum REPROVADO é efeito colateral da entrega 44 |
| Cálculos das frentes no Relatório B (aprovados / reprovados) | 4 / 0 | 9 + 1 corrida do lote + 1 Westgard (11 / 0) | 15 / 0 |
| Conferência aritmética de bias/ET/Sigma feita no consolidado (B.3) | 4 linhas do `engPainel` + 2 do Painel | 6 do `engPainel` + 6 do Painel | 18 conferidas, 0 divergentes (método não validado) |
| Maior diferença relativa (n, média, DP, CV) | 3,3e-13 (DP da aba Estatística); 1,9e-16 no `engPainel` | 3,1e-16 | Todas abaixo de 1e-9 |
| Ocorrências das frentes / deduplicadas (P-xx) | 1 / — | 1 / — | 2 / 2 (P-01 das duas frentes; P-02 de acabamento, aberta no consolidado) |
| Críticas / Altas / Médias / Baixas (deduplicadas) | — | — | 0 / 0 / 1 / 1 |
| Dúvidas de regra de negócio (deduplicadas) | 2 | 3 | 3 |

**Situação do item do reteste**

| Item | Bio | Hema | Consolidado |
|---|---|---|---|
| OC-0078: linha manual apagada ou movida mantinha o carimbo | **RESOLVIDA**. Os critérios 1, 2, 3 (origem), 4, 5 e 6 foram aprovados. O manual digitado na linha apagada por outro usuário (B = MAN_0007; E = MAN_0008; N = MAN_0009) ganha ID e carimbo novos, fica ATIVO e entra no Painel, na Estatística e no LJ (slots 134 e 135). Base: os três ficam MANUAL_INCOMPLETO "DATA/HORA posterior ao lançamento", e o E reaproveita o MAN_0002 | **RESOLVIDA**. Mesmos critérios aprovados. O critério 1 foi fechado no complemento c1 (MAN_0004 ATIVO, slot 73, Painel N1 n 21 = oráculo). No roteiro principal, o MAN_0007 ganhou carimbo novo, mas caiu na regra dos 30 min (CONFLITO_MANUAL a 29,6 min do HEM-379943). É regra correta, não é o defeito. Base: defeito reproduzido | **RESOLVIDA** nos dois produtos, sem efeito colateral, com prova pela base de que o roteiro pega o defeito. **Ressalva (T03b):** no recorte, o ID vai com a linha, mas o carimbo não. O destino recebe o carimbo de quem moveu, e o USUARIO_MANUAL do resultado já integrado muda sem registro no Audit_Log. É o limite declarado no ADR-097 §5, e a base faz o mesmo. Contradiz o texto do critério de aceite 3 da OC-0078 (P-01, D-01) |

---

## Relatório A: experiência do usuário

### A.1 Funcionalidades testadas

| Frente | Bioquímica | Hematologia |
|---|---|---|
| F1: apagar uma linha manual | Delete de ID…MOTIVO; Delete só de DATA…MOTIVO (o ID fora da seleção); Delete de 2 linhas de uma vez; Delete numa linha com SUBSTITUI_ID | Os mesmos, mais "apagar só o ID de uma linha com dado" (C04, no c1) |
| F1: reaproveitar a linha apagada | Outro usuário, 2,5 min depois, pelo botão Digitar Resultados, com DATA/HORA posterior ao carimbo antigo (o caso que era recusado) | Idem, com 3 níveis; complemento c1 com horário a mais de 30 min de qualquer real do dia |
| F1: mover e repetir | Recortar ID…MOTIVO de um manual já integrado (origem e destino); Ctrl+D de uma linha carimbada para 2 linhas | Idem |
| F1: corrigir e substituir | Editar o RESULTADO de um manual existente na própria linha; ADR-082 (SUBSTITUI_ID digitado só com o número) | Idem; a MATRIZ SANGUE TOTAL se preenche sozinha |
| F1: legado | Carimbo órfão gravado com os eventos desligados numa linha vazia, limpo pelo ATUALIZAR | Idem |
| F1: o que o usuário vê depois | Painel (`engPainel`, A7:T8), Eng_Saida (pontos do LJ, corrida, marcas), aba Estatística de novembro, tblCQ_Final, Audit_Log | Painel (A7:T9), Eng_Saida com a corrida do lote, aba Estatística por ano (B3/D3 = 2026) |
| Base (entrega 43) | Mesmo roteiro, para provar que o teste pega o defeito | Idem |

### A.2 Facilidade de uso

**Reaproveitar uma linha apagada agora funciona como o usuário espera.**
- O botão Digitar Resultados leva à linha apagada.
- O que se digita ali fica no nome e na hora de quem digitou, agora, e entra no Painel, na Estatística e no LJ.
- O usuário não precisa saber que existem colunas de carimbo. Na entrega 43, o mesmo gesto perdia o resultado com
  um motivo que ele não tinha como corrigir.

**Apagar ficou consistente.** De ID…MOTIVO, de DATA…MOTIVO, de duas linhas de uma vez ou numa linha com SUBSTITUI_ID, a
linha fica realmente vazia: sem ID solto, carimbo ou SUBSTITUI_ID esquecido. O ID que sobrava não é mais
reaproveitado pelo próximo manual.

**O Ctrl+D é seguro.** A linha de cima não muda de dono. As de baixo ganham ID novo, com uma caixa só, que lista os IDs
gerados e lembra de usar a hora real para não cair no CONFLITO_MANUAL.

**Corrigir na própria linha** mantém o ID e o dono, e fica no Audit_Log como RESULTADO_MANUAL_ALTERADO, como o GUIA
manda.

**SUBSTITUI_ID** continua prático: digitar só o número (379301 / 379002) vira BIO-/HEM-, a linha se preenche
sozinha, e o I04 explica a troca e a diferença (0,123).

**Dá para ver o que foi salvo.** O n, a média e o ponto no LJ mudam depois do ATUALIZAR. Os três manuais que a entrega
43 recusava aparecem na corrida certa.

### A.3 Dificuldades encontradas

**Recortar uma linha manual muda o dono dela em silêncio (P-01, Médio).**
- O ID vai com a linha, mas o carimbo do destino passa a ser o de quem moveu, na hora do recorte.
- Depois do ATUALIZAR, a tblCQ_Final atribui o lançamento a quem só moveu, e o Audit_Log não registra a troca.
- O usuário não tem como perceber: as colunas de carimbo são travadas e nada avisa.

**Texto do aviso do Ctrl+D (P-02, Baixo).** A caixa diz "As 2 linhas **coladas** traziam IDs…" também quando o gesto
foi o Ctrl+D (preencher para baixo). É compreensível, mas "copiadas" ou "repetidas" descreveria os dois gestos. As
frentes o trataram como acabamento; registro no consolidado para o Orquestrador decidir.

**Regra dos 30 min ao reaproveitar a linha (só registro, D-03).** Na Hema, um manual com a hora de "agora", digitado
logo depois de um automático do mesmo nível, caiu a 29,6 min dele e ficou CONFLITO_MANUAL. A regra está certa e
documentada (ADR-082, GUIA §3), e o DETALHE e o A01 explicam o que fazer. Na rotina da Hema, porém, esse caso tende a
ser comum.

**Apagar só o ID gera outro na hora, sem aviso (Hema, T09).** O dado não se perde nem duplica, mas quem olha a tela vê
o número mudar sem saber por quê. O GUIA §3 já manda apagar o ID nesse caso. Não é ocorrência.

### A.4 Navegação

- O botão Digitar Resultados leva à primeira linha livre, que agora pode ser uma linha apagada sem resíduo.
- O Painel e o Eng_Saida refletem o ATUALIZAR sem passo extra.
- Na Hema, a frente pôs o Ano De/Até da Estatística em 2026 à mão. O aviso novo do ADR-097 para o lote fora do
  intervalo (OC-0048) não foi exercitado por ela.

### A.5 Apresentação

- A trilha da retirada (RESULTADO_MANUAL_RETIRADO) fica no nome de quem rodou o ATUALIZAR, e não de quem apagou a linha
  (D-02).
- As corridas do dia foram renumeradas depois da retirada dos manuais apagados (Bio: MAN_0003 passou de 26111602 para
  26111601), como o ADR-080 prevê.

### A.6 Tempos observados

As frentes não mediram tempo neste reteste.

### A.7 Sugestões do usuário

Para o Orquestrador priorizar. Nenhuma foi implementada.

1. **P-01:** quando o destino recebe um ID MAN_ que já existe na tblCQ_Final e está sem carimbo, recuperar dali o
   USUARIO_MANUAL e o instante do lançamento. Alternativas:
   - o motor registrar no Audit_Log a troca de USUARIO_MANUAL de um manual existente;
   - mínima: o GUIA §3 dizer "não recorte linhas manuais; para corrigir, edite na própria linha".

   Em qualquer caso, alinhar o critério de aceite 3 da OC-0078 ao ADR-097 §5 (D-01).
2. **P-02:** no aviso do `EntradaMudou`, trocar "linhas coladas" por "linhas copiadas ou repetidas".
3. **D-02:** se a trilha precisar dizer quem retirou, registrar o Delete no momento em que acontece.

---

## Relatório B: validação matemática

**Escopo deste ciclo.** Só a F1. O que se validou:
1. n, média, DP e CV das séries que recebem os manuais digitados, apagados, movidos, editados e substituídos, no Painel
   (`engPainel` e A7:T8/T9) e na aba Estatística, antes e depois do 2º ATUALIZAR;
2. que os valores que **não** podem entrar não entraram;
3. os pontos novos no Eng_Saida (fonte do LJ), a corrida do lote (ADR-080) e a ausência de marcas de Westgard
   (nenhuma era esperada).

Bias (do CEQ), ET e Sigma **não** foram recalculados de forma independente a partir dos dados brutos (não houve F3 nem
F2). O consolidado só conferiu a aritmética das colunas do Painel (B.3).

**Oráculo das frentes:** Python puro (`statistics.fmean`, `statistics.stdev`, CV = 100·DP/média) sobre o `Value2` da
tblCQ_Final da cópia. O filtro é PARTICIPA_ESTATISTICA = "SIM", analito, equipamento, lote, nível e a janela de
DATA_HORA. Tolerância: a padrão, 1e-9 relativo (REGRAS §3.5).

### B.1 Tabela (n, média, DP, CV)

Bio: cópia `Bioquimica_F1_nova.xlsm` (produção fdc173906fff). Painel: L4 DIMENSION 1 · C3 Glicose · H3 lote 8977 ·
K3 2026 · K4 Nov. Estatística: janela 01/11/2026–30/11/2026.
Hema: cópias `Hematologia_F1_nova.xlsm` e `Hematologia_F1_c1.xlsm` (produção 5b0e78ace7b9). Painel: C3 WBC · H3 lote
625311 · K3 2026 · K4 Nov (G3/G4 01/11/2026–30/11/2026). Estatística: B3 = D3 = 2026, B4 = 625311.

Formato: n · média · DP · CV %. Fórmula encontrada em todas as linhas: média = Σx/n; DP amostral (n − 1);
CV = 100·DP/média.

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI (Value2) | Diferença (relativa) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026, antes | 12 reais, Σ 930,0 | média = Σx/n; DP amostral; CV = 100·DP/média | 12 · 77,5 · 0,7977240352 · 1,0293213358 | `engPainel` e Estatística: idem | 0 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026, antes | 12 reais, Σ 3272,0 | idem | 12 · 272,6666666667 · 6,3149653901 · 2,3160019768 | `engPainel` e Estatística: idem | DP 1,4e-16 (`engPainel`), 1,4e-13 (Estatística) | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N1 | nov/2026, 2º ATUALIZAR | 930 − BIO-379301 77,0 (substituído) + MAN_0004 78,851 (editado) + MAN_0005 79,545 + MAN_0010 79,7 + MAN_0011 78,928 (Ctrl+D) + MAN_0012 77,123 (substitui) + MAN_0007 79,083 (o que era recusado). MAN_0001 retirado | 1326,230/17 | 17 · **78,0135294118** · 1,0547901520 · 1,3520605463 | `engPainel` e Painel A7:E7: idem; Estatística DP 1,054790151975805 | média 0; DP 0 (`engPainel`), 3,3e-13 (Estatística); CV 1,6e-16 | APROVADO |
| Glicose | DIMENSION 1 | 8977 | N2 | nov/2026, 2º ATUALIZAR | 3272 + MAN_0003 275,781 (movido) + MAN_0006 275,027 + MAN_0008 275,379 (o que era recusado) + MAN_0009 273,67 (digitado na origem do recorte). MAN_0002 retirado; BIO-379339 continua | 4371,857/16 | 16 · **273,2410625** · 5,5197913785 · 2,0201178139 | `engPainel` e Painel A8:E8: idem; Estatística DP 5,519791378530835 | média 0; DP 0 (`engPainel`), 5,5e-14 (Estatística) | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026, antes | 18 reais, Σ 56,44 | idem | 18 · 3,1355555556 · 0,0700606647 · 2,2343939826 | `engPainel`: idem | CV 2,0e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, antes | 24 reais, Σ 175,81 | idem | 24 · 7,3254166667 · 0,1043813729 · 1,4249206247 | `engPainel`: idem | média 2,4e-16; CV 3,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026, antes | 13 reais, Σ 227,34 | idem | 13 · 17,4876923077 · 0,2380179911 · 1,3610600355 | `engPainel`: idem | 0 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026, roteiro, 2º ATUALIZAR | 56,44 − HEM-379002 3,15 (substituído) + MAN_0012 3,273 + MAN_0004 3,156 (editado) + MAN_0005 3,192 + MAN_0010 3,2 + MAN_0011 3,16 + reais HEM-379713 3,09 e HEM-379943 3,13. MAN_0001 retirado; MAN_0007 (CONFLITO 30 min) fora | 75,491/24 | 24 · **3,1454583333** · 0,0692299126 · 2,2009483278 | `engPainel` e Painel A7:E7: idem | 0; CV 2,0e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, roteiro, 2º ATUALIZAR | 175,81 + MAN_0006 7,337 + MAN_0008 7,346 (o que era recusado) + reais HEM-379548 7,4 e HEM-379788 7,3. MAN_0002 retirado; HEM-379077 continua | 205,193/28 | 28 · **7,3283214286** · 0,0975868288 · 1,3316395813 | `engPainel` e Painel: idem | média 1,2e-16; CV 1,7e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026, roteiro, 2º ATUALIZAR | 227,34 + MAN_0003 17,491 (movido) + MAN_0009 17,384 (digitado na origem do recorte) + real HEM-379863 17,67 | 279,885/16 | 16 · **17,4928125** · 0,2196072915 · 1,2554144255 | `engPainel` e Painel: idem | média 2,0e-16; CV 1,8e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1 | nov/2026, c1, 2º ATUALIZAR | 56,44 + reais 3,09 e 3,13 + **MAN_0004 3,168** (o manual do critério 1). MAN_0001 retirado | 65,828/21 | 21 · **3,1346666667** · 0,0657976697 · 2,0990324223 | `engPainel`: idem | 2,1e-16 | APROVADO |
| WBC | XN1000 | 625311 | N2 | nov/2026, c1, 2º ATUALIZAR | 175,81 + reais 7,4 e 7,3 + MAN_0003 7,314 (C04: o mesmo dado com ID novo, contado uma vez). MAN_0002 retirado | 197,824/27 | 27 · **7,3268148148** · 0,0993985904 · 1,3566412269 | `engPainel`: idem | 2,4e-16 | APROVADO |
| WBC | XN1000 | 625311 | N3 | nov/2026, c1, 2º ATUALIZAR | 227,34 + real 17,67 | 245,01/14 | 14 · 17,5007142857 · 0,2338133472 · 1,3360217381 | `engPainel`: idem | 2,0e-16 | APROVADO |
| WBC | XN1000 | 625311 | N1/N2/N3 | ano 2026 (Estatística), antes / roteiro / c1 | PARTICIPA = SIM do lote em 2026 | idem | antes 39 · 3,1274358974 · 0,0616320981 / 63 · 7,3326984127 · 0,1094591741 / 34 · 17,4232352941 · 0,2065990200; roteiro 45 · 3,1338 · 0,0632310338 / 67 · 7,3334776119 · 0,1065004422 / 37 · 17,4306756757 · 0,2023143942; c1 42 · 3,1275714286 · 0,0599585804 / 66 · 7,3329393939 · 0,1073302422 / 35 · 17,4302857143 · 0,2077680510 | Estatística: iguais | ≤ 2,4e-16 | APROVADO |
| WBC (corrida do lote, ADR-080) | XN1000 | 625311 | N1/N2/N3 | roteiro, 2º ATUALIZAR | Eng_Saida, última corrida do lote por nível | corrida = posição na sequência de ativos do lote e nível | n de 2026 = 45 / 67 / 37 | Eng_Saida: corrida final 45 / 67 / 37 | 0 | APROVADO |
| Glicose e WBC (Westgard) | DIMENSION 1 / XN1000 | 8977 / 625311 | todos | pontos de teste | Alvo do lote: Bio N1 79,3140496 ± 1,5427784, N2 274,7758621 ± 5,0269933; Hema N1 3,18 ± 0,07935, N2 7,33 ± 0,1368, N3 17,44 ± 0,25395 | z = (x − média do alvo)/DP do alvo; regra só dispara com \|z\| ≥ 2 | maior \|z\|: Bio MAN_0012 (77,123 − 79,3140)/1,5428 = −1,42; Hema MAN_0012 (3,273 − 3,18)/0,07935 = +1,17 → nenhuma marca nova | Eng_Saida: nenhuma marca nos slots de teste; status do Painel OK | nenhuma | APROVADO |

**Valores que não podem entrar e não entraram** (nos dois produtos):
- os manuais apagados MAN_0001 e MAN_0002;
- o resultado substituído (BIO-379301; HEM-379002): SUBSTITUIDO_MANUAL, PARTICIPA NÃO, fora do Eng_Saida;
- as linhas apagadas MAN_0013/0014/0015;
- o valor antigo do MAN_0004 (Bio 79,005; Hema 3,164);
- na Hema, o MAN_0007 (CONFLITO pela regra dos 30 min) e, no c1, o MAN_0002 antigo do C04.

O n bate exatamente com a contagem dos ATIVOS.

**Base (entrega 43):** os três manuais B, E e N ficam MANUAL_INCOMPLETO e fora da estatística. O defeito da OC-0078
era de entrada, não de cálculo. A base é o arquivo de 15/11, com outros dados reais, por isso os n não se comparam com
os da nova.

### B.2 LJ (Eng_Saida, a fonte do gráfico) depois do 2º ATUALIZAR

| Produto | Slot | RUN | Corrida do lote (Hema N1/N2/N3) | N1 (ID · valor) | N2 (ID · valor) | N3 (ID · valor) | Marcas |
|---|---|---|---|---|---|---|---|
| Bio | 126 | 26111601 | — | MAN_0004 · 78,851 | MAN_0003 · 275,781 | — | nenhuma |
| Bio | 127 | 26111602 | — | MAN_0005 · 79,545 | MAN_0006 · 275,027 | — | nenhuma |
| Bio | 128 | 26111603 | — | MAN_0010 · 79,7 | — | — | nenhuma |
| Bio | 129 | 26111604 | — | MAN_0011 · 78,928 | — | — | nenhuma |
| Bio | 131 | 26111606 | — | MAN_0012 · 77,123 (substitui BIO-379301) | — | — | nenhuma |
| Bio | 133 | 26111608 | — | — | BIO-379339 · 271,0 | — | nenhuma |
| Bio | 134 | 26111609 | — | **MAN_0007 · 79,083** | **MAN_0008 · 275,379** | — | nenhuma |
| Bio | 135 | 26111610 | — | — | **MAN_0009 · 273,67** | — | nenhuma |
| Hema | 68 | 26111601 | 39 / 63 / 34 | MAN_0012 · 3,273 (substitui HEM-379002) | HEM-379077 · 7,36 | HEM-379157 · 17,78 | OK |
| Hema | 70 | 26111701 | 40 / 65 / 35 | MAN_0004 · 3,156 | MAN_0006 · 7,337 | MAN_0003 · 17,491 | OK |
| Hema | 71 | 26111702 | 41 / — / — | MAN_0005 · 3,192 | — | — | OK |
| Hema | 72 | 26111703 | 42 / — / — | MAN_0010 · 3,2 | — | — | OK |
| Hema | 73 | 26111704 | 43 / — / — | MAN_0011 · 3,16 | — | — | OK |
| Hema | 76 | 26111707 | — / 67 / 37 | — | **MAN_0008 · 7,346** | **MAN_0009 · 17,384** | OK |

No c1, o MAN_0004 (3,168) do critério 1 está no slot 73, corrida do lote N1 42, sem marca.

Os pontos da tabela batem com os valores da tblCQ_Final do Relatório B.1: não há divergência entre tabela e gráfico.

### B.3 Conferência aritmética de bias, ET e Sigma (feita no consolidado, sem Excel)

Fonte: `Value2` das linhas do Painel e do `engPainel` gravadas nos JSON das frentes (blocos `antes` e `depois2`). Não
recalculei o bias do CEQ (`mCEQ.BiasEQ`) nem o CV pooled do Sigma a partir dos dados brutos. Por isso esta seção
**não** valida o método do ADR-035/064, só a coerência aritmética das colunas.

**1. Painel (A7:T8 na Bio, A7:T9 na Hema), colunas G:I, depois do 2º ATUALIZAR, e Hema antes:**
- O ETp é o do cadastro: Glicose 10 %, WBC 7 %.
- O bias exibido é o mesmo nos níveis e **não muda** quando se incluem ou corrigem manuais do CIQ: Bio 2,7332944433 %,
  Hema 1,8648146012 %. É o esperado para um bias do CEQ.
- O ET exibido = |bias| + 1,65·CV do período, com diferença 0 nas 6 linhas conferidas (Bio N1 4,9641943447; N2
  6,0664888363. Hema antes 5,5515646725 / 4,2159336319 / 4,1105636598; depois 5,4963793420 / 4,0620199104 /
  3,9362484032).
- O Sigma exibido **não** usa o CV do período. O CV implícito, (ETp − |bias|)/Sigma, é:
  - Bio: N1 2,1063534818; N2 2,2144333724;
  - Hema antes: 2,0779586158 / 1,3966311964 / 0,9310091387;
  - Hema depois: 2,0854095598 / 1,3854306412 / 0,9293557839.

  Ele muda pouco com os manuais do lote 625311, o que é compatível com o CV pooled do CIQ (lotes com n ≥ 20, janela de
  6 meses; ADR-064). **Não conferido** contra os dados brutos.

**2. `engPainel`, colunas 6 a 9 (o "bias/Sigma oculto" da OC-0010):**
- O bias é 100·(média do período − média do alvo do lote)/média do alvo do lote.
- O ET = |bias| + 1,65·CV do período, e o Sigma = (ETp − |bias|)/CV do período.
- Nas 10 linhas (Bio N1/N2 e Hema N1 a N3, antes e depois), a diferença relativa foi de no máximo 5,3e-13 no bias, 1,0e-14
  no ET e 3,8e-15 no Sigma. Exemplo: Bio N1 depois, (78,0135294118 − 79,3140495868)/79,3140495868 = −1,6397097132 %.
- É outro bias (CIQ contra o alvo), diferente do exibido no Painel, como já registrado na OC-0010 (aberta). Não é
  ocorrência nova.

**Observação de método (para o Especialista em CQ, não reprovada).** No Painel, o ET usa o CV do período e o Sigma usa
outro CV (o implícito acima). Se o ADR-064 manda o CV pooled só no Sigma, está coerente. Se manda também no ET, é
diferença de método. A F3 deste ciclo não foi executada para decidir.

### B.4 Etapas para reproduzir e o que o consolidado conferiu

1. Copiar a produção atual e o backup da entrega 43 para `%TEMP%\qcw\auditoria\<ciclo>\` e conferir o sha256: Bio
   fdc173906fff e bd5314426ab8; Hema 5b0e78ace7b9 e 49f68fba99d5.
2. Rodar, uma cópia por vez: `reteste07_usuario_f1.py nova`, `... base` (Bio); `reteste07_usuario_f1_hema.py nova`,
   `... base` e `reteste07_usuario_f1_hema_c1.py` (Hema).
3. Oráculo: filtrar a tblCQ_Final por PARTICIPA_ESTATISTICA = "SIM", ANALITO, EQUIPAMENTO, LOTE, NIVEL e DATA_HORA
   (nov/2026 no Painel; nov/2026 na Estatística da Bio; ano 2026 na Estatística da Hema). Depois aplicar
   `statistics.fmean`, `statistics.stdev` e CV = 100·DP/média.
4. Conta à mão: n₁ = n₀ − retirados − substituídos + manuais ATIVOS + reais novos; Σ₁ idem.

**Conferido de novo neste consolidado (sem Excel):**
- Somas e médias refeitas à mão, todas iguais às publicadas:
  - Bio: 1326,23/17 = 78,0135294; 4371,857/16 = 273,2410625;
  - Hema roteiro: 75,491/24 = 3,1454583; 205,193/28 = 7,3283214; 279,885/16 = 17,4928125;
  - Hema c1: 65,828/21 = 3,1346667; 197,824/27 = 7,3268148; 245,01/14 = 17,5007143.
- Diferença relativa oráculo × `engPainel`, lida dos JSON (`antes`, `depois2`, todos os níveis): n idêntico; média,
  DP e CV com no máximo 1,9e-16 na Bio e 3,1e-16 na Hema.
- z do Westgard: −1,4201 (Bio) e +1,1720 (Hema).
- A aritmética de B.3.
- Os números da aba Estatística foram tomados dos relatórios das frentes; não os reli nos JSON.

### B.5 Caminho do dado (conferido neste ciclo)

- **Manual até o Painel e o LJ:**
  1. Digitar Resultados → `Worksheet_Change`/`EntradaMudou`: limpa o carimbo da linha vazia (`LimparCarimboManual`),
     tira o ID e o SUBSTITUI_ID que sobram, tira o carimbo do lançamento novo (`LancamentoNovo`), gera ID novo para ID
     repetido e carimba (`Carimbar`).
  2. ATUALIZAR (`PrepararEntradas` limpa o órfão legado; `mEtlCQ`: chave, status, RUN; regra dos 30 min; DATA/HORA >
     REGISTRADO_EM → recusa).
  3. tblCQ_Final (PARTICIPA = SIM).
  4. Série equipamento × analito × lote × nível × período.
  5. Eng_Saida (pontos, alvo, corrida do lote, marcas).
  6. `engPainel`, Painel e aba Estatística.
- **Automático (parcial):** DB_SEAC → ATUALIZAR → tblCQ_Final com ID HEM-<id>. Na Hema, os 5 reais novos do filtro
  (HEM-379548, 379713, 379788, 379863 e 379943) entraram no nível, no lote e na data certos e foram contados uma vez.
  Na Bio, os 2 novos eram de outros analitos e ficaram fora do filtro.
- **Trilha:** Audit_Log RESULTADO_MANUAL_INCLUIDO / ALTERADO / RETIRADO / RECUSADO e RESULTADO_SUBSTITUIDO_POR_MANUAL;
  QA_INTEGRACAO I04, A01 e E04. Lacuna: a troca de USUARIO_MANUAL por recorte não gera linha no Audit_Log (P-01).
- **CEQ → bias → ET → Sigma:** não exercitado. Só se viu que o bias exibido não muda com o CIQ e que o ET reage ao
  CV do período (B.3).

---

## Relatório C: ocorrências

**Deduplicação.**
- As frentes trouxeram 2 ocorrências, a R7-01 da Bio e a mesma R7-01 da Hema: ficam como **P-01** (Ambos).
- O consolidado abriu a **P-02** (Baixo), a partir do texto do aviso que as duas frentes descreveram nas impressões
  como acabamento.

Os IDs P-xx são provisórios; o ID definitivo (OC-xxxx) é dado pelo Orquestrador. Nenhuma é Crítica nem Alta. **Nenhuma
é efeito colateral da entrega 44:** a P-01 é idêntica na base (entrega 43) nos dois produtos. A P-02 é texto do aviso de ID
repetido; o texto da base não foi comparado (na Hema, a frente só anotou que a base não tem a frase do CONFLITO_MANUAL).

### C.1 Tabela

| ID | Produto | Prioridade | Tipo | Tela/módulo | Resumo | Origem | Relação | Base (entrega 43) |
|---|---|---|---|---|---|---|---|---|
| P-01 | Ambos | Médio | `funcional` | Digitar Resultados (`mIntegracao.EntradaMudou`/`PrepararLinha`/`Carimbar`) e motor (USUARIO_MANUAL da tblCQ_Final, Audit_Log) | Recortar e colar um manual já integrado troca o dono (USUARIO_MANUAL) para quem moveu, sem registro no Audit_Log | Bio R7-01 (T03b); Hema R7-01 (T03b) | Critério de aceite 3 da OC-0078 × ADR-097 §5 (limite conhecido) | Igual nos dois |
| P-02 | Ambos | Baixo | `usabilidade` | Aviso do `EntradaMudou` (ID repetido, ADR-096) | O aviso diz "As N linhas coladas…" também no Ctrl+D (preencher para baixo) | Impressões da Bio e da Hema (T04) | — | Texto não comparado |

### C.2 Detalhe

**P-01 (Ambos, Médio, funcional): recortar e colar um manual já integrado troca o dono, sem registro no Audit_Log**
- **Tela/módulo:** aba Digitar Resultados (`mIntegracao.EntradaMudou`/`PrepararLinha`/`Carimbar`) e motor
  (USUARIO_MANUAL da tblCQ_Final, Audit_Log).
- **Ação:**
  1. U1 digita o manual (Bio MAN_0003, N2 275,781, carimbo R7 10:11:57; Hema MAN_0003, N3 17,491, carimbo R7H
     08:39:12) e roda o ATUALIZAR. O USUARIO_MANUAL fica U1.
  2. U2 seleciona ID…MOTIVO, recorta e cola 9 linhas abaixo.
  3. U2 roda o ATUALIZAR.
- **Esperado:** pelo critério de aceite 3 da OC-0078 ("o ID e o carimbo vão com ela"), o destino mantém o carimbo
  original, e o resultado continua atribuído a quem o lançou.
- **Observado:**
  - O ID vai (MAN_0003). O carimbo do destino passa a ser o de quem moveu, na hora do recorte (Bio R7B 10:14:29; Hema
    R7HB 08:41:43).
  - Depois do ATUALIZAR, a tblCQ_Final mostra USUARIO_MANUAL = U2, e o Audit_Log não tem nenhuma linha do MAN_0003:
    nem ALTERADO, nem "movido".
  - O valor, a data/hora, o status e o ponto do LJ não mudam.
  - É idêntico na base (entrega 43).
  - O ADR-097 §5 declara o carimbo de quem moveu como limite conhecido. Não menciona a troca do USUARIO_MANUAL de um
    resultado já integrado nem a falta de rastro no Audit_Log.
- **Dificuldade:** nenhuma para o usuário, que nem percebe: as colunas de carimbo são travadas e nada avisa.
- **Consequência:** a trilha do resultado de CQ fica errada sem deixar rastro. O sistema passa a dizer que quem lançou
  foi quem apenas moveu, e o lançamento original deixa de constar. Nenhum número muda. Mover uma linha manual é raro.
- **Sugestão (para decisão do Orquestrador):**
  - quando o destino recebe um ID MAN_ que já existe na tblCQ_Final e está sem carimbo, recuperar dali o
    USUARIO_MANUAL e o instante;
  - ou o motor registrar no Audit_Log a troca de USUARIO_MANUAL de um manual existente;
  - ou, no mínimo, o GUIA §3 orientar "não recorte linhas manuais; corrija na própria linha";
  - em qualquer caso, alinhar o critério de aceite 3 da OC-0078 ao ADR-097 §5 (D-01).
- **Evidência:** `%TEMP%\qcw\auditoria\2026-11-16_reteste07\Bioquimica_F1_nova.json` e `Hematologia_F1_nova.json`:
  - `passos.F2c_recorte` (`M_antes` × `destino_logo_apos_recorte`);
  - `atualizar1.tblCQ_Final_meus` × `atualizar2.tblCQ_Final_meus` (MAN_0003);
  - `atualizar2.auditoria`, sem linha do MAN_0003;
  - nas bases (`*_F1_base.json`), os mesmos campos.

  Scripts `ferramentas/reteste07_usuario_f1.py` e `reteste07_usuario_f1_hema.py`. Produções fdc173906fff e
  5b0e78ace7b9; bases bd5314426ab8 e 49f68fba99d5.
- **Não executado:** o Ctrl+X/Ctrl+V físico (emulado com `Range.Cut(destino)`).

**P-02 (Ambos, Baixo, usabilidade): o aviso do ID repetido fala em "linhas coladas" também no Ctrl+D**
- **Tela/módulo:** aviso do `EntradaMudou` para ID MAN_ repetido (ADR-096).
- **Ação:** selecionar ID…MOTIVO de 3 linhas, com a de cima carimbada, e preencher para baixo (Ctrl+D).
- **Esperado:** um aviso que descreva o gesto feito.
- **Observado:** a caixa única diz "As 2 linhas coladas traziam IDs que já existem em outras linhas e receberam IDs
  novos (MAN_0010 a MAN_0011)…". O conteúdo está certo: IDs novos e o lembrete de usar a hora real.
- **Dificuldade:** pequena: quem usou o Ctrl+D pode estranhar "coladas".
- **Consequência:** nenhuma sobre os dados.
- **Sugestão:** "As N linhas copiadas ou repetidas…".
- **Evidência:** T04 dos dois relatórios das frentes (texto da caixa lido da janela do PID); `*_F1_nova.json`, passo
  do Ctrl+D.

### C.3 Dúvidas de regra de negócio (para o Orquestrador encaminhar)

| ID | Produto | Para | Dúvida |
|---|---|---|---|
| D-01 | Ambos | Orquestrador / Especialista em CQ | (P-01) O carimbo de uma linha manual movida deve ser o do lançamento original, como diz o critério de aceite 3 da OC-0078, ou o de quem moveu, como diz o ADR-097 §5? Hoje é o de quem moveu, e o USUARIO_MANUAL do resultado já integrado muda em silêncio |
| D-02 | Ambos | Orquestrador / Especialista em CQ | O Audit_Log grava RESULTADO_MANUAL_RETIRADO no nome de quem rodou o ATUALIZAR, e não de quem apagou a linha (Bio: apagou R7, gravou R7B; Hema: apagou R7H/R7HC, gravou R7HB/R7HD). É o esperado? Se a trilha precisa dizer quem retirou, o Delete teria de ser registrado no momento em que acontece |
| D-03 | Hematologia | Especialista em CQ (só registro) | Um manual com a hora de "agora" pode cair a menos de 30 min de um automático do mesmo nível e ficar CONFLITO_MANUAL (MAN_0007 a 29,6 min do HEM-379943). A regra está certa e documentada; na rotina da Hema, esse caso tende a ser comum. Sem ocorrência |
| D-04 | Ambos | Especialista em CQ | (B.3) No Painel, o ET usa o CV do período e o Sigma usa outro CV (compatível com o pooled do ADR-064). É essa a definição do projeto para os dois? |

---

## Relatório D: parecer final

Critério: "validado" só onde houve teste executado e aprovado neste ciclo. O resto é "não validado" (teste reprovado)
ou "não executado" (com o motivo). Este ciclo foi o reteste de 1 item (OC-0078), só com a F1 e os fluxos vizinhos.

### D.1 A rotina funciona?

**Sim, no que foi refeito.** A OC-0078 está RESOLVIDA nos dois produtos, sem efeito colateral, com uma ressalva no
destino do recorte (P-01, Médio).

**Validado (Bio e Hema):**
- apagar uma linha manual de qualquer forma (ID…MOTIVO, DATA…MOTIVO, 2 linhas, com SUBSTITUI_ID) deixa a linha
  realmente vazia;
- o manual digitado depois na linha apagada, por outro usuário, ganha ID e carimbo novos, fica ATIVO e entra no
  Painel, na Estatística e no LJ;
- a origem de um recorte fica limpa, e o manual digitado nela fica ATIVO;
- o Ctrl+D não muda o dono da origem;
- editar na própria linha mantém o ID e o dono;
- o ADR-082 (SUBSTITUI_ID) continua funcionando;
- o carimbo órfão legado é limpo no ATUALIZAR;
- na Hema, apagar só o ID gera outro sem perda nem duplicidade;
- a base (entrega 43) reproduz o defeito com o mesmo roteiro, o que prova que o teste o pega.

**Não validado:** T03b, o carimbo e o USUARIO_MANUAL de uma linha manual recortada (P-01), nos dois produtos.

**Não executado:**
- teclado e mouse físicos;
- o intervalo de dias (reproduzido com 2,5 min);
- na Bio, apagar só o ID de uma linha com dado (T12, coberto pelo C04 do `qa_adr097`, do Tester).

### D.2 Bias, ET e Sigma estão validados?

**Não executado neste ciclo** (não houve F3 nem F2). O bias do CEQ e o CV pooled do Sigma não foram recalculados dos
dados brutos.

O consolidado conferiu só a aritmética das colunas do Painel (B.3):
- o bias exibido não muda com o CIQ;
- o ET = |bias| + 1,65·CV do período, com diferença 0;
- o Sigma usa outro CV (D-04);
- o bias/Sigma do `engPainel` (CIQ contra o alvo, OC-0010) é coerente com a própria fórmula.

Isso **não** é validação do método. A última validação independente de bias, ET e Sigma é a do reteste 04
(entrega 41).

### D.3 A integração CIQ + CEQ está validada?

**Não executado neste ciclo.** Nenhum CEQ foi incluído nem corrigido. Só se observou que incluir e corrigir CIQ não
mexe no bias exibido e que o ET e o Sigma reagem ao CV (B.3).

### D.4 Os filtros de período e de lote estão corretos?

**Validado, no escopo executado:**
- Painel com lote no H3 e mês nov/2026: as séries Glicose DIMENSION 1 lote 8977 (N1, N2) e WBC XN1000 lote 625311 (N1 a
  N3) recebem só os ATIVOS do lote e do mês, com n, média, DP e CV iguais ao oráculo;
- aba Estatística da Bio (janela 01–30/11/2026) e da Hema (ano 2026, lote 625311): iguais ao oráculo;
- corrida do lote no LJ da Hema = n de 2026 (45/67/37).

**Não executado:**
- mês a mês além de novembro, trimestres, troca de lote, H3 vazio, outros equipamentos;
- a Estatística da Hema com o Ano De/Até do arquivo e o aviso novo do ADR-097 (OC-0048).

A F4 não fez parte deste reteste.

### D.5 A entrada manual e a automática estão validadas?

**Entrada manual: validada nos cenários da OC-0078,** nos dois produtos: apagar, reaproveitar, recortar (origem),
Ctrl+D, editar, SUBSTITUI_ID, várias linhas e órfão. Status, RUN, tblCQ_Final, Audit_Log, I04, A01 e E04 foram
conferidos, e os números são iguais ao oráculo.

**Não validado:** o dono do manual recortado (P-01).

**Entrada automática:** validada só no escopo do filtro.
- Na Hema, os 5 resultados reais novos do DB_SEAC (HEM-<id>) entraram no lote, no nível e na data certos, contados
  uma vez, e a conta à mão fecha.
- Na Bio, os 2 novos ficaram fora do filtro.
- **Não executados:** o fluxo completo do recebimento (Sync_Log, o que não entrou, duplicidade fora do filtro) e a
  equivalência manual × automático com o mesmo valor.

### D.6 Há divergência matemática?

**Não**, nos cálculos executados.
- Relatório B.1: 15 linhas das frentes, 15 aprovadas (13 de n/média/DP/CV, 1 da corrida do lote, 1 de Westgard).
- Maior diferença relativa: 3,3e-13 (DP da Estatística da Bio). No `engPainel`, no máximo 1,9e-16 na Bio e 3,1e-16 na
  Hema. Todas abaixo de 1e-9.
- Valores que não podem entrar não entraram.
- B.3: 18 conferências aritméticas de bias/ET/Sigma sem divergência (máximo 5,3e-13). O método não foi validado
  (D.2), e há uma dúvida de método (D-04).

### D.7 Há problema operacional?

**Nada Alto nem Crítico.** O problema Alto do ciclo anterior (OC-0078) foi resolvido.
- **P-01 (Médio):** recortar uma linha manual atribui o lançamento a quem moveu, sem rastro no Audit_Log. É raro, não
  muda número e é anterior à entrega 44.
- **P-02 (Baixo):** o texto "linhas coladas" no Ctrl+D.
- **D-03:** a regra dos 30 min pode recusar, na Hema, o manual digitado logo depois de um automático do mesmo nível.
  A regra está correta.

As produções ficaram intactas: mesmo hash antes e depois de todas as execuções.

### D.8 O que fica pendente (o que falta)

1. **OC-0078:** o Usuário Simulado a dá como RESOLVIDA nos dois produtos. Falta a **validação do Tester** (VALIDAR),
   que não consta desta pasta, para passar a VALIDADA no registro.
2. **P-01 (Médio):** decidir a D-01 e então:
   - corrigir (recuperar o dono ou registrar a troca no Audit_Log);
   - ou documentar no GUIA §3 e aceitar o limite do ADR-097 §5.

   Em qualquer caso, alinhar o critério 3. **P-02 (Baixo):** decidir o texto (só Tester, pela cadência de 15/11).
3. **Dúvidas:** encaminhar a D-01 e a D-02 (trilha de quem moveu e de quem apagou), a D-03 (só registro) e a D-04
   (qual CV entra no ET e qual no Sigma) ao Especialista em CQ.
4. **OC da entrega 44 em AGUARDA RETESTE, não retestadas neste ciclo:**
   - OC-0048 (Ano De/Até da Hema), OC-0079 (GUIA: copiar ID…MOTIVO), OC-0080 (DETALHE do ID fora do padrão) e OC-0081
     (aviso e A01 do repetido por chave), todas Médio;
   - OC-0082 e OC-0083 (filtros do Registros), Baixo.

   Também segue em AGUARDA RETESTE a OC-0053 (|Bias| digitado sobre a fórmula, entrega 37).
5. **Seguem abertas, de ciclos anteriores:**
   - aguardando o laboratório: OC-0007, OC-0008, OC-0011, OC-0015, OC-0016, OC-0017;
   - do Orquestrador: OC-0010, OC-0014, OC-0020, OC-0022, OC-0025, OC-0026, OC-0037, OC-0038, OC-0039, OC-0045, OC-0046,
     OC-0047, OC-0061, OC-0062, OC-0063, OC-0064, OC-0065, OC-0069, OC-0070, OC-0077;
   - OC-0060: aprovado provisório até 30/11;
   - OC-0067: troca da senha que está no repositório público, com o usuário.
6. **Não executados, a cobrir num próximo ciclo:**
   - F3 sobre a entrega 44 (bias do CEQ, CV pooled, ET, Sigma; ETp ausente "-"; ETp VB/FAB);
   - F2 (CEQ);
   - o recebimento automático completo e a equivalência manual × automático;
   - F4 completa (mês a mês, trimestres, troca de lote, H3 vazio, LJ e Westgard com violação);
   - teclado e mouse físicos, que pedem uma sessão interativa.

Dados pessoais: nenhum login de pessoa, máquina ou caminho de rede consta deste parecer. Os usuários citados
(`TESTE_AUDITORIA_R7*`) são de teste, criados só nas cópias.
