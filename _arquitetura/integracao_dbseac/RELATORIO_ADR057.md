# QC_INI — Nova arquitetura de dados (ADR-057) · Relatório de entrega

**Data:** 02/10/2026 · **Produtos:** Hematologia e Bioquímica · **Branch:** `integracao-db-seac`
**Decisão:** [ADR-057](ADR-057.md) · **Fontes:** `_arquitetura/integracao_dbseac/`

---

## A. Auditoria inicial

**DB_SEAC.xlsm** (`INTERFACEAMENTO_DB_COPIA\DB_SEAC.xlsm`) — Power Query sobre os CSV do SIPEC
(relat207 Hematologia, relat209 Bioquímica), janela móvel de 12 meses, atualizado pela macro `ATUALIZAR`.

| Tabela | Colunas | Linhas | Observações |
|---|---|---|---|
| `tbHematologia` (DADOS_HEMATOLOGIA) | EQUIPAMENTO, LOTE, NIVEL, NIVEL_DESC, DATA, DATA_HORA, ANALITO, VALOR, FLAG + **ID_ORIGEM, ID_AMOSTRA, ITEM_ID, UNIDADE** | 44.392 | 1 equipamento (XN1000), 9 lotes (522611 … 625311), 31 analitos, 93 sem valor (flag `----`), 589 retransmissões |
| `tbBioquimica` (DADOS_BIOQUIMICA) | EQUIPAMENTO, MATRIZ, BLOCO, LOTE, NIVEL, NIVEL_DESC, DATA, DATA_HORA, ANALITO, PRIMEIRO_DO_DIA, VALOR, FLAG + as 4 novas | 32.033 | DIMENSION 1 e 2, matrizes SORO/URINA, 7 blocos, 29 lotes, réplicas reais (ITEM_ID diferente) |

- **Chave de origem:** a coluna `id` do CSV do SIPEC é única por resultado e sequencial. Não existia no
  DB_SEAC; a cópia foi levada à **v2.4** (`alterar_db_seac.py`): as 4 colunas novas são geradas pelo próprio
  M de `DB_HEMATOLOGIA`/`DB_BIOQUIMICA`, então **sobrevivem a qualquer `ATUALIZAR`** (o worker agendado só
  chama essa macro). Assinaturas de integridade (CONTROLE), histórico de versão (VALIDACAO) e LOG_AUDITORIA
  atualizados; integridade CONFORME.
- **Duplicidade real:** na Hematologia, 589 linhas são retransmissão (mesma amostra, item, instante e valor,
  `id` diferente); na Bioquímica há réplicas legítimas (outro `ITEM_ID`). Data + analito não é chave.

**QC_INI antes** (Hematologia e Bioquímica): DB_Resultados gravada por frmCorrida / frmMassa (Hema) / aba
Importar (Bio) e exclusão lógica por frmExcluir; elegibilidade em Cfg_Status; repetições Rep 1-3 em Registros
(X roxo); aba Resultados (visão); EQC_Dados marcada "LEGADO (ADR-034)". Mapa de dependências completo
(VBA × fórmulas × gráficos × nomes) levantado antes de qualquer remoção.

## B. Arquitetura nova

```
                     ENTRADAS
             ┌──────────┴──────────┐
             ▼                     ▼
   INTERFACEAMENTO (DB_SEAC)   Digitar Resultados (MAN_0001…)
             ▼                     │
            DB  ──► DB_ORGANIZADO  │   (visão horizontal, só conferência)
             │                     │
             └──────────┬──────────┘
                        ▼
        PROCESSAMENTO (Power Query) ◄── Inativar (+ COMENTARIOS_TECNICOS)
                        ▼
             Principal - Resultados   (tblCQ_Final · DB_CQ_FINAL)
                        │
            ┌───────────┴───────────┐
            ▼                       ▼
     LEVEY-JENNINGS             ESTATÍSTICA
  normal · X vermelho ·     cálculos + bloco de
      não plotar             repetições por analito
```

Abas e botões com o **mesmo nome**: **DB** · **Inativar** · **Digitar Resultados** · **Principal - Resultados**
(+ DB_ORGANIZADO, COMENTARIOS_TECNICOS, QA_INTEGRACAO, Cfg_Integracao). Botão **⟳ ATUALIZAR DADOS** no Painel e
em todas as abas de dados.

## C. Power Query

| Consulta | Tipo | O que faz |
|---|---|---|
| `CFG_INTEGRACAO` | conexão | lê `tblConfigIntegracao` (caminho, prefixo, GAP, tolerância) |
| `SEAC_ORIGEM` | conexão | lê a tabela do setor no DB_SEAC (nomes e ordem do DB_SEAC), tipa, cria `ID_REGISTRO = HEM-/BIO-<id>` |
| `DB_RECEBIMENTO` → aba **DB** | tabela | **acumula**: lê a própria tabela e acrescenta só IDs novos (anti-junção sobre o já recebido); nunca reescreve nem apaga |
| `DB_ORGANIZADO` | tabela | **Table.Pivot** (analito → colunas, média do instante), contagem de agregados |
| `DB_CQ_FINAL` → aba **Principal - Resultados** | tabela | ver D |
| `QA_INTEGRACAO` | tabela | E01–E05, A01–A06, I01–I03 |

Operações em `DB_CQ_FINAL`: de/para de analito (junção plana com `tblDeParaAnalitos`); detecção de retransmissão
(agrupamento nativo por 10 colunas + junção plana só das chaves repetidas); staging dos manuais (validação,
normalização `MAN_nnnn`, ID repetido → sufixo `~L<n>`); conflito manual × interfaceamento (±30 min, mesma
chave) e manual × manual; **append** (UNION ALL) interface + manuais; **LEFT OUTER** com a inativação (nunca
anti-join) e com os comentários; classificação; corrida/RUN numa tabela estreita (índice + coluna da linha
anterior + `Table.FillDown`) devolvida por junção no ID. **Sem unpivot:** o recebimento já é vertical (uma
linha por resultado); pivotar para depois despivotar seria trabalho sem ganho — a grade horizontal é só a
visão DB_ORGANIZADO.

## D. Principal - Resultados (DB_CQ_FINAL)

- **Granularidade:** uma linha = um resultado (ID_REGISTRO único). Hematologia: **44.392 linhas × 39 colunas**.
- Campos: ID_REGISTRO, ORIGEM_RESULTADO, SETOR, EQUIPAMENTO, MATRIZ, BLOCO, DATA, HORA, DATA_HORA, CORRIDA_NO_DIA,
  RUN, NIVEL, NIVEL_DESC, LOTE, ANALITO, ANALITO_ORIGEM, ANALITO_CADASTRADO, RESULTADO, UNIDADE, FLAG,
  **STATUS_ANALITICO, PARTICIPA_ESTATISTICA, REGISTRAR_RESULTADO_NO_LJ, TIPO_PLOTAGEM_LJ**, INATIVACAO_REGISTRADA,
  COMENTARIO_TECNICO, TEM_JUSTIFICATIVA, GOVERNANCA, DATA_INATIVACAO, USUARIO_INATIVACAO,
  MOTIVO_EXCLUSAO_AUTOMATICA, ID_RELACIONADO, ID_ORIGEM, ID_AMOSTRA, ITEM_ID, PRIMEIRO_DO_DIA, MOTIVO_MANUAL,
  USUARIO_MANUAL, RECEBIDO_EM — responde às 14 perguntas do pedido (ID, origem, automático/manual, data,
  corrida, nível, lote, analito, resultado, ativo, participa, aparece no LJ, como plota, justificativa).
- **Regras de status** (precedência): MANUAL_INCOMPLETO › DUPLICIDADE_ORIGEM › CONFLITO_MANUAL › SEM_VALOR ›
  INATIVADO › ATIVO. Só ATIVO participa. Estados de plotagem:

| Estado | PARTICIPA_ESTATISTICA | REGISTRAR_RESULTADO_NO_LJ | TIPO_PLOTAGEM_LJ |
|---|---|---|---|
| Normal | SIM | SIM | NORMAL |
| Inativado + REGISTRAR - LJ marcado (padrão) | NÃO | SIM | X_VERMELHO |
| Inativado + desmarcado | NÃO | NÃO | NAO_PLOTAR |

- **RUN** = yymmdd·100 + corrida do dia (blocos por equipamento/matriz/lote/analito/dia com intervalo ≤ 10 min;
  o 2º resultado do mesmo nível abre posição nova). Garante ≤ 1 participante por (equipamento, lote, nível,
  analito, RUN) — conferido: 0 violações. Inativar/reativar não muda RUN.

## E. Inativação (aba Inativar)

Soft-delete: digitar/colar o **ID_REGISTRO** (o número puro também vale — o evento normaliza para `HEM-`/`BIO-`).
O resultado continua na Principal - Resultados com PARTICIPA_ESTATISTICA = NÃO. **REGISTRAR - LJ** é checkbox
nativo, nasce **marcado** (SIM → X vermelho); desmarcado → não plota. DATA_INATIVACAO e USUARIO são carimbados
pelo evento. Colunas de conferência (analito, nível, lote, data/hora, RUN, resultado, plotagem, justificativa)
vêm da tabela final por XLOOKUP. **Reabilitar = apagar a linha** (mesmo ID, sem redigitar). Justificativa
obrigatória em COMENTARIOS_TECNICOS: sem ela o QA acusa **E01 ERRO DE GOVERNANÇA** (ID, analito, nível,
data). Cada inativação, reativação e troca de plotagem vai ao Audit_Log (cadeia de hash).

## F. Resultados manuais (aba Digitar Resultados)

Mesmo schema lógico do interfaceamento (DATA, HORA, EQUIPAMENTO, MATRIZ, LOTE, NÍVEL, ANALITO, RESULTADO,
UNIDADE, MOTIVO + RUN/STATUS/DETALHE de conferência). O ID **MAN_0001, MAN_0002…** é gerado sozinho (maior + 1,
nunca reaproveitado) — formato que não colide com `HEM-`/`BIO-`. Entra pelo **append** e segue as mesmas regras
(RUN, inativação). ORIGEM_RESULTADO = MANUAL. Se o mesmo resultado chegar depois pelo interfaceamento (±30 min,
mesma chave), o manual vira **CONFLITO_MANUAL** (guardado, fora do cálculo, ID_RELACIONADO aponta o do
interfaceamento) e o QA emite **A01**; fora da tolerância no mesmo dia, **A06** — nunca duplica em silêncio.

## G. Levey-Jennings

O motor lê **só a Principal - Resultados** (`mDados.CarregarDB`). Pontos = PARTICIPA = SIM; X = TIPO_PLOTAGEM_LJ
= X_VERMELHO, publicados em bloco próprio (`Eng_Saida!AD:AL`, nomes `engXN1..3`) — fora das colunas que o
Painel soma. O eixo do Westgard é só o das corridas com resultado elegível (idêntico ao de antes); a corrida que
só tem X ganha posição no gráfico, sem veredicto. As séries "Repetição" (X roxo) viraram **"Não conforme (X)"**,
marcador X **vermelho**, mesmo tamanho do ponto normal, **sem linha**. Comprovação: testes T09–T14 (seção J).

## H. Estatística

Hematologia: `EstatBasica` sobre o índice do motor (só "Ativo" = PARTICIPA SIM). Bioquímica: `EstatPeriodo`
reapontada para a mesma fonte, com filtro de equipamento e cache invalidado a cada atualização (a inativação
muda PARTICIPA sem mudar o nº de linhas). Bloco **RESULTADOS NÃO CONFORMES / REPETIÇÕES** (Estatística!A142):
uma fórmula de matriz dinâmica — analitos do cadastro × contagem de STATUS_ANALITICO = INATIVADO no lote,
período (e equipamento) da aba. Rótulo: resultados de controle retirados da população principal; não é erro
laboratorial nem resultado de paciente. Comprovação: T17–T18.

## I. Legado

| Removido | Por quê | Dependências resolvidas |
|---|---|---|
| Aba **Importar** (Bio) + `mImportar` | porta de entrada concorrente | IrLancar → Digitar Resultados |
| Aba **Resultados** + `mOperacao` (view) | substituída por Principal - Resultados | IrResultados removido |
| Aba **EQC_Dados** ("LEGADO ADR-034") | não alimentava nada | nenhuma fórmula/VBA lia (provado) |
| Aba **DB_Resultados** + `mBanco` + nomes r*, capBanco | fonte antiga | Início reapontado para tblCQ_Final; AnoDosDados idem |
| Aba **Cfg_Status** | elegibilidade agora no PQ | EhElegivel removido; Audit_Legenda atualizado |
| `frmCorrida`, `frmMassa`, `frmExcluir`, `frmConfigEstatistica` | lançamento/exclusão antigos; form órfão | AbrirForm* removidos |
| **Rep 1, Rep 2, Rep 3** (Registros F:H) + regRep1-3 | substituídas pelo X da inativação | Calc reapontado; RegistrosStore alinhada (12→9 colunas) |

**Preservados:** EQA.CAP_Dados, EQA.Controllab_Dados, EQA_Base e módulos mEQA/mCEQ (controle externo vigente).
A remoção é condicionada: o instalador procura referência às abas legadas em todas as fórmulas e nomes e
**aborta sem remover** se achar alguma.

## J. QA

Bateria `testes/qa_final.py`, rodada sobre **cópias byte a byte dos arquivos entregues** (SHA-256 conferido).
O teste usa o caminho real: escreve nas abas Inativar / Digitar Resultados / COMENTARIOS_TECNICOS com os eventos
ligados, roda o mesmo núcleo do botão ATUALIZAR DADOS e confere a Principal - Resultados, o que o Levey-Jennings
plota (Eng_Saida/Calc), a Estatística, o QA e o Audit_Log. Evidência completa em `testes/resultados/qa_*.json`;
imagens do gráfico com o X em `testes/resultados/LJ_*_X_vermelho.png`.

**Resultado: Hematologia 27 PASS / 0 FAIL · Bioquímica 27 PASS / 0 FAIL.**

Defeitos achados e corrigidos durante o QA (nenhum ficou aberto):
- erro de compilação 450 (variável `linhasX` escondendo a função `LinhasX` — o VBA não diferencia maiúsculas):
  função renomeada `LinhasXDe`; o `patch_vba.py` passou a recusar sombreamento e o instalador compila o VBA com
  cão de guarda;
- coluna espúria "DadosExternos_1: Obtendo dados…" na aba DB (texto provisório da 1ª carga lido pela própria
  consulta acumulativa): nome da tabela só depois da 1ª carga + filtro de colunas na consulta;
- bloco de repetições perdia o 1º dia do período quando a data inicial tinha hora: `INT()` nas duas pontas;
- erro de automação abria a janela de depuração do VBA: entrada `AtualizarDadosAutomatico` (sem diálogo).

### Hematologia: 27 PASS / 0 FAIL

| Teste | Resultado | Evidência |
|---|---|---|
| T01b Botões: cada aba de dados tem os botões DB · Inativar · Digitar Resultados · Principal - Resultados (mesmo nome das abas) | **PASS** | {"DB": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Inativar": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Digitar Resultados": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Principal - Re… |
| T01 Estrutura: abas novas presentes; LEGADO/IMPORTAR/Resultados/DB_Resultados/Cfg_Status removidas; Rep 1-3 removidas; módulos legados removidos | **PASS** | {"abas_novas": ["DB", "DB_ORGANIZADO", "Digitar Resultados", "Inativar", "COMENTARIOS_TECNICOS", "Principal - Resultados", "QA_INTEGRACAO", "Cfg_Integracao"], "legado_restante": [], "modulos_legados_restantes": [], "cabecalho_Registros": ["Nº", "Data", "Analit… |
| T02 DB_RECEBIMENTO recebe o DB_SEAC com as colunas do DB_SEAC + ID_REGISTRO + RECEBIDO_EM | **PASS** | {"linhas": 44392, "colunas": ["ID_REGISTRO", "EQUIPAMENTO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA", "ANALITO", "VALOR", "FLAG", "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE", "RECEBIDO_EM"], "ids_unicos": 44392} |
| T03 DB_ORGANIZADO horizontal (uma coluna por analito) | **PASS** | {"linhas": 1413, "colunas": ["DATA", "DATA_HORA", "EQUIPAMENTO", "LOTE", "NIVEL", "NIVEL_DESC", "WBC", "RBC", "HGB", "HCT", "MCV", "MCH", "..."], "n_colunas": 39} |
| T04 DB_CQ_FINAL: uma linha por resultado, ID único, campos obrigatórios, <=1 participante por corrida/nível | **PASS** | {"linhas": 44392, "colunas": 39, "ids_unicos": 44392, "status": {"ATIVO": 43710, "DUPLICIDADE_ORIGEM": 589, "SEM_VALOR": 93}, "plotagem": {"NORMAL": 43710, "NAO_PLOTAR": 682}, "tempo_atualizacao_s": 85.6, "resumo": "Atualização concluída.\r\n\r\nResultados rec… |
| T05 Resultado real escolhido no LJ em tela (analito/lote/RUN) | **PASS** | {"analito": "BASO#", "lote": "522611", "equip": null, "slot": [55, 25103103, 0.15], "candidato": {"ID_REGISTRO": "HEM-213735", "ORIGEM_RESULTADO": "INTERFACEAMENTO", "SETOR": "HEMATOLOGIA", "EQUIPAMENTO": "XN1000", "MATRIZ": "SANGUE TOTAL", "BLOCO": null, "DAT… |
| T06 ID estável: mesmos IDs e mesmo RUN após 3 atualizações | **PASS** | {"ids": 44392, "ids_base": 44392, "exemplo": ["HEM-213735", 25103103.0]} |
| T07 Inativação por ID: soft-delete (origem intacta, registro continua, PARTICIPA=NÃO), checkbox padrão SIM, carimbo automático | **PASS** | {"linha_inativacao": {"ID_REGISTRO": "HEM-213735", "REGISTRAR - LJ": true, "DATA_INATIVACAO": "2026-10-02 15:24:46", "USUARIO": "QCINI", "ANALITO": "=IF(tblInativacao_NaoConformes[[#This Row],[ID_REGISTRO]]=\"\",\"\",IFERROR(XLOOKUP(tblInativacao_NaoConformes[… |
| T08 Governança: inativado SEM comentário = QA ERRO E01 (FAIL de governança detectado) | **PASS** | {"achado": {"SEVERIDADE": "ERRO", "CODIGO": "E01", "TESTE": "Resultado inativado sem justificativa tecnica", "ID_REGISTRO": "HEM-213735", "ANALITO": "BASO#", "NIVEL": 1.0, "DATA_HORA": "2025-10-31 14:12:24", "DETALHE": "ERRO DE GOVERNANÇA: HEM-213735 ¦ BASO# ¦… |
| T09 LJ: inativado com REGISTRAR-LJ=SIM aparece como X na posição da corrida, fora da série normal e do Westgard | **PASS** | {"slot": 55, "RUN": 25103103, "engXN1": 0.15, "engValN1": null, "Calc_X": 0.15, "Calc_valor": "", "veredicto_Westgard_N1": null} |
| T09b Principal - Resultados -> Levey-Jennings: todo resultado que participa está no ponto da sua corrida/nível, nenhum ponto sem origem, todo inativado com REGISTRAR-LJ aparece como X | **PASS** | {"pontos_conferidos": 151, "divergentes": [], "pontos_sem_origem": 0, "X_ausentes": [], "janela_corridas": 56} |
| T09c Imagem do gráfico com o X exportada para conferência visual | **PASS** | {"arquivo": "C:\\Users\\VITOR~1.SAN\\AppData\\Local\\Temp\\claude\\C--Users-vitor-santos-OneDrive---MSFT-INI-DB-AJUSTES-OneDrive---MSFT-Desktop-QC-INI-PRONTO\\cc5cf662-a93c-4a60-a8b5-23064798503c\\scratchpad\\bench\\final\\LJ_Hematologia_X_vermelho.png"} |
| T10 X sem impacto estatístico: n/média/DP do Painel = população SIM (recalculada à parte); n caiu 1 | **PASS** | {"antes": [1.0, 50.0, 0.1454, 0.0057888, 3.981292904], "depois": [1.0, 49.0, 0.145306122, 0.005810207, 3.998597397], "independente": [49, 0.14530612244897959, 0.005810206829482101]} |
| T11 Gráfico: série "Não conforme (X)" com marcador X vermelho e SEM linha ligando à série | **PASS** | {"series": ["Não conforme (X)", "Não conforme (X)", "Não conforme (X)"], "marcador": [-4168, -4168, -4168], "cor_BGR": 255, "tamanho": [6, 6, 6]} |
| T12 Comentário técnico: com justificativa o QA passa (E01 some), TEM_JUSTIFICATIVA=SIM | **PASS** | {"COMENTARIO_TECNICO": "QA automatizado: resultado fora do esperado, repetido.", "GOVERNANCA": "OK"} |
| T13 REGISTRAR-LJ desmarcado: não plota, não participa, continua auditável; estatística igual ao estado X | **PASS** | {"TIPO_PLOTAGEM_LJ": "NAO_PLOTAR", "slot_no_LJ": 55, "painel_X": [1.0, 49.0, 0.145306122, 0.005810207, 3.998597397], "painel_nao_plotar": [1.0, 49.0, 0.145306122, 0.005810207, 3.998597397]} |
| T14 Reabilitação: apagar o ID volta ATIVO/NORMAL, mesmo ID, sem duplicar, estatística volta ao valor inicial | **PASS** | {"final": {"ID_REGISTRO": "HEM-213735", "STATUS_ANALITICO": "ATIVO", "PARTICIPA_ESTATISTICA": "SIM", "TIPO_PLOTAGEM_LJ": "NORMAL", "RUN": 25103103.0}, "linhas": 44392, "painel_inicial": [1.0, 50.0, 0.1454, 0.0057888, 3.981292904], "painel_reabilitado": [1.0, 5… |
| T15 Resultado manual MAN_0001: ORIGEM=MANUAL, PARTICIPA=SIM, RUN pela mesma regra, aparece no LJ | **PASS** | {"linha_manual": {"ID_REGISTRO": "MAN_0001", "DATA": "2025-10-31 03:00:00", "HORA": 0.9993055555555556, "EQUIPAMENTO": "XN1000", "MATRIZ": "SANGUE TOTAL", "LOTE": "522611", "NIVEL": 1.0, "ANALITO": "BASO#", "RESULTADO": 0.15, "UNIDADE": null, "MOTIVO": "Falha … |
| T16 Duplicidade manual x interfaceamento: chave (equip, matriz, lote, nível, analito, dia, ±30 min) -> CONFLITO_MANUAL; interfaceamento prevalece; QA A01; nada duplica em silêncio | **PASS** | {"manual": {"ID_REGISTRO": "MAN_0002", "STATUS_ANALITICO": "CONFLITO_MANUAL", "ID_RELACIONADO": "HEM-212368", "MOTIVO_EXCLUSAO_AUTOMATICA": "coincide com resultado do interfaceamento (4 min)"}, "interface": {"ID_REGISTRO": "HEM-212368", "STATUS_ANALITICO": "AT… |
| T17 Estatística: bloco RESULTADOS NÃO CONFORMES / REPETIÇÕES conta {'WBC': 3, 'RBC': 2, 'HGB': 1} | **PASS** | {"bloco": {"WBC": 3.0, "RBC": 2.0, "HGB": 1.0}, "outros_nao_zero": {}, "ids": {"WBC": ["HEM-202969", "HEM-203266", "HEM-203401"], "RBC": ["HEM-202970", "HEM-203267"], "HGB": ["HEM-202971"]}, "filtro": {"lote": "522611", "ini": 45658.0, "fim": 46022.0, "equip":… |
| T18 Estatística usa só PARTICIPA=SIM da DB_CQ_FINAL: n caiu exatamente o número de inativados; n/média batem com o cálculo independente | **PASS** | {"WBC": {"antes": [50.0, 3.044599999999999, 0.054928896525138604], "depois": [47.0, 3.0451063829787235, 0.05520209276221688], "independente": [47, 3.0451063829787235, 0.05520209276221688]}, "RBC": {"antes": [50.0, 2.4022, 0.030256724669982107], "depois": [48.0… |
| T19 Refresh idempotente (4x): mesma contagem, mesmos IDs/RUN, inativações, comentários e manuais mantidos | **PASS** | {"linhas": [44394, 44394], "tempos_s": [111.2, 135.0, 119.1, 146.2], "manuais": ["MAN_0002", "MAN_0001"]} |
| T20 Refresh síncrono: falha de camada PARA o processo, não diz "concluída", não atualiza gráfico/estatística | **PASS** | {"erro": "Etapa: atualizar tblDB_Recebimento\r\nErro 1004: [DataSource.NotFound] File or Folder: Não foi possível localizar o arquivo 'C:\\Users\\vitor.santos\\OneDrive - MSFT\\INI_DB_AJUSTES\\OneDrive - MSFT\\Desktop\\QC_INI_PRONTO\\qc-ini\\INTERFACEAMENTO_DB… |
| T21 Audit_Log: inativação, reativação, plotagem e manual registrados (ISO 15189 8.4) | **PASS** | {"RESULTADO_INATIVADO": 7, "PLOTAGEM_LJ_ALTERADA": 1, "RESULTADO_REATIVADO": 1, "RESULTADO_MANUAL_INCLUIDO": 2, "ATUALIZACAO_FALHOU": 1} |
| T22 Fonte única: nenhum módulo analítico nem fórmula de LJ/Estatística lê DB_SEAC, DB_RECEBIMENTO, manual, inativação, Resultados, Importar ou DB_Resultados | **PASS** | [] |
| T23 VBA compila (Depurar > Compilar) e os formulários que ficaram abrem | **PASS** | {"formularios": "frmAssinar;frmDev;"} |
| T24 Fechar, reabrir e atualizar: camadas, motor, LJ e Estatística funcionando | **PASS** | {"linhas_final": 44392, "corridas_no_LJ": 56, "tempo_1a_s": 104.7, "tempo_apos_reabrir_s": 191.1, "resumo": "Atualização concluída.\r\n\r\nResultados recebidos: 44.392 (novos nesta atualização: 0)\r\nPrincipal - Resultados: 44.392 resultados\r\nQA da integraçã… |

### Bioquimica: 27 PASS / 0 FAIL

| Teste | Resultado | Evidência |
|---|---|---|
| T01b Botões: cada aba de dados tem os botões DB · Inativar · Digitar Resultados · Principal - Resultados (mesmo nome das abas) | **PASS** | {"DB": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Inativar": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Digitar Resultados": ["DB", "Inativar", "Digitar Resultados", "Principal - Resultados"], "Principal - Re… |
| T01 Estrutura: abas novas presentes; LEGADO/IMPORTAR/Resultados/DB_Resultados/Cfg_Status removidas; Rep 1-3 removidas; módulos legados removidos | **PASS** | {"abas_novas": ["DB", "DB_ORGANIZADO", "Digitar Resultados", "Inativar", "COMENTARIOS_TECNICOS", "Principal - Resultados", "QA_INTEGRACAO", "Cfg_Integracao"], "legado_restante": [], "modulos_legados_restantes": [], "cabecalho_Registros": ["Nº", "Data", "Analit… |
| T02 DB_RECEBIMENTO recebe o DB_SEAC com as colunas do DB_SEAC + ID_REGISTRO + RECEBIDO_EM | **PASS** | {"linhas": 32033, "colunas": ["ID_REGISTRO", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC", "DATA", "DATA_HORA", "ANALITO", "PRIMEIRO_DO_DIA", "VALOR", "FLAG", "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "UNIDADE", "RECEBIDO_EM"], "ids_unicos": 320… |
| T03 DB_ORGANIZADO horizontal (uma coluna por analito) | **PASS** | {"linhas": 4594, "colunas": ["DATA", "DATA_HORA", "EQUIPAMENTO", "MATRIZ", "BLOCO", "LOTE", "NIVEL", "NIVEL_DESC", "LA", "URCA", "ALB", "DBI", "..."], "n_colunas": 45} |
| T04 DB_CQ_FINAL: uma linha por resultado, ID único, campos obrigatórios, <=1 participante por corrida/nível | **PASS** | {"linhas": 32033, "colunas": 39, "ids_unicos": 32033, "status": {"ATIVO": 32033}, "plotagem": {"NORMAL": 32033}, "tempo_atualizacao_s": 85.1, "resumo": "Atualização concluída.\r\n\r\nResultados recebidos: 32.033 (novos nesta atualização: 0)\r\nPrincipal - Resu… |
| T05 Resultado real escolhido no LJ em tela (analito/lote/RUN) | **PASS** | {"analito": "Cloro", "lote": "8974", "equip": "DIMENSION 1", "slot": [179, 26070601, 106.0], "candidato": {"ID_REGISTRO": "BIO-314216", "ORIGEM_RESULTADO": "INTERFACEAMENTO", "SETOR": "BIOQUIMICA", "EQUIPAMENTO": "DIMENSION 1", "MATRIZ": "SORO", "BLOCO": "PAIN… |
| T06 ID estável: mesmos IDs e mesmo RUN após 3 atualizações | **PASS** | {"ids": 32033, "ids_base": 32033, "exemplo": ["BIO-314216", 26070601.0]} |
| T07 Inativação por ID: soft-delete (origem intacta, registro continua, PARTICIPA=NÃO), checkbox padrão SIM, carimbo automático | **PASS** | {"linha_inativacao": {"ID_REGISTRO": "BIO-314216", "REGISTRAR - LJ": true, "DATA_INATIVACAO": "2026-10-02 15:59:42", "USUARIO": "QCINI", "ANALITO": "=IF(tblInativacao_NaoConformes[[#This Row],[ID_REGISTRO]]=\"\",\"\",IFERROR(XLOOKUP(tblInativacao_NaoConformes[… |
| T08 Governança: inativado SEM comentário = QA ERRO E01 (FAIL de governança detectado) | **PASS** | {"achado": {"SEVERIDADE": "ERRO", "CODIGO": "E01", "TESTE": "Resultado inativado sem justificativa tecnica", "ID_REGISTRO": "BIO-314216", "ANALITO": "Cloro", "NIVEL": 1.0, "DATA_HORA": "2026-07-06 08:09:54", "DETALHE": "ERRO DE GOVERNANÇA: BIO-314216 ¦ Cloro ¦… |
| T09 LJ: inativado com REGISTRAR-LJ=SIM aparece como X na posição da corrida, fora da série normal e do Westgard | **PASS** | {"slot": 179, "RUN": 26070601, "engXN1": 106.0, "engValN1": null, "Calc_X": 106.0, "Calc_valor": "", "veredicto_Westgard_N1": null} |
| T09b Principal - Resultados -> Levey-Jennings: todo resultado que participa está no ponto da sua corrida/nível, nenhum ponto sem origem, todo inativado com REGISTRAR-LJ aparece como X | **PASS** | {"pontos_conferidos": 282, "divergentes": [], "pontos_sem_origem": 0, "X_ausentes": [], "janela_corridas": 180} |
| T09c Imagem do gráfico com o X exportada para conferência visual | **PASS** | {"arquivo": "C:\\Users\\VITOR~1.SAN\\AppData\\Local\\Temp\\claude\\C--Users-vitor-santos-OneDrive---MSFT-INI-DB-AJUSTES-OneDrive---MSFT-Desktop-QC-INI-PRONTO\\cc5cf662-a93c-4a60-a8b5-23064798503c\\scratchpad\\bench\\final\\LJ_Bioquimica_X_vermelho.png"} |
| T10 X sem impacto estatístico: n/média/DP do Painel = população SIM (recalculada à parte); n caiu 1 | **PASS** | {"antes": [1.0, 148.0, 113.324324324, 6.670287673, 5.886015833], "depois": [1.0, 147.0, 113.37414966, 6.66539919, 5.87911725], "independente": [147, 113.37414965986395, 6.665399189921798]} |
| T11 Gráfico: série "Não conforme (X)" com marcador X vermelho e SEM linha ligando à série | **PASS** | {"series": ["Não conforme (X)", "Não conforme (X)", "Não conforme (X)"], "marcador": [-4168, -4168, -4168], "cor_BGR": 255, "tamanho": [6, 6, 6]} |
| T12 Comentário técnico: com justificativa o QA passa (E01 some), TEM_JUSTIFICATIVA=SIM | **PASS** | {"COMENTARIO_TECNICO": "QA automatizado: resultado fora do esperado, repetido.", "GOVERNANCA": "OK"} |
| T13 REGISTRAR-LJ desmarcado: não plota, não participa, continua auditável; estatística igual ao estado X | **PASS** | {"TIPO_PLOTAGEM_LJ": "NAO_PLOTAR", "slot_no_LJ": 179, "painel_X": [1.0, 147.0, 113.37414966, 6.66539919, 5.87911725], "painel_nao_plotar": [1.0, 147.0, 113.37414966, 6.66539919, 5.87911725]} |
| T14 Reabilitação: apagar o ID volta ATIVO/NORMAL, mesmo ID, sem duplicar, estatística volta ao valor inicial | **PASS** | {"final": {"ID_REGISTRO": "BIO-314216", "STATUS_ANALITICO": "ATIVO", "PARTICIPA_ESTATISTICA": "SIM", "TIPO_PLOTAGEM_LJ": "NORMAL", "RUN": 26070601.0}, "linhas": 32033, "painel_inicial": [1.0, 148.0, 113.324324324, 6.670287673, 5.886015833], "painel_reabilitado… |
| T15 Resultado manual MAN_0001: ORIGEM=MANUAL, PARTICIPA=SIM, RUN pela mesma regra, aparece no LJ | **PASS** | {"linha_manual": {"ID_REGISTRO": "MAN_0001", "DATA": "2026-07-06 03:00:00", "HORA": 0.9993055555555556, "EQUIPAMENTO": "DIMENSION 1", "MATRIZ": "SORO", "LOTE": "8974", "NIVEL": 1.0, "ANALITO": "Cloro", "RESULTADO": 107.06, "UNIDADE": null, "MOTIVO": "Falha do … |
| T16 Duplicidade manual x interfaceamento: chave (equip, matriz, lote, nível, analito, dia, ±30 min) -> CONFLITO_MANUAL; interfaceamento prevalece; QA A01; nada duplica em silêncio | **PASS** | {"manual": {"ID_REGISTRO": "MAN_0002", "STATUS_ANALITICO": "CONFLITO_MANUAL", "ID_RELACIONADO": "BIO-310844", "MOTIVO_EXCLUSAO_AUTOMATICA": "coincide com resultado do interfaceamento (5 min)"}, "interface": {"ID_REGISTRO": "BIO-310844", "STATUS_ANALITICO": "AT… |
| T17 Estatística: bloco RESULTADOS NÃO CONFORMES / REPETIÇÕES conta {'Glicose': 3, 'Ureia': 2, 'Creatinina': 1} | **PASS** | {"bloco": {"Glicose": 3.0, "Ureia": 2.0, "Creatinina": 1.0}, "outros_nao_zero": {}, "ids": {"Glicose": ["BIO-203133", "BIO-203166", "BIO-203617"], "Ureia": ["BIO-203132", "BIO-203165"], "Creatinina": ["BIO-203136"]}, "filtro": {"lote": "8974", "ini": 45931.0, … |
| T18 Estatística usa só PARTICIPA=SIM da DB_CQ_FINAL: n caiu exatamente o número de inativados; n/média batem com o cálculo independente | **PASS** | {"Glicose": {"antes": [173.0, 92.36416184971098, 30.444961568441585], "depois": [170.0, 91.3, 26.741110340607047], "independente": [170, 91.3, 26.741110340607037]}, "Ureia": {"antes": [199.0, 34.28643216080402, 9.210411699315062], "depois": [197.0, 33.99492385… |
| T19 Refresh idempotente (4x): mesma contagem, mesmos IDs/RUN, inativações, comentários e manuais mantidos | **PASS** | {"linhas": [32035, 32035], "tempos_s": [62.7, 70.7, 64.4, 102.7], "manuais": ["MAN_0002", "MAN_0001"]} |
| T20 Refresh síncrono: falha de camada PARA o processo, não diz "concluída", não atualiza gráfico/estatística | **PASS** | {"erro": "Etapa: atualizar tblDB_Recebimento\r\nErro 1004: [DataSource.NotFound] File or Folder: Não foi possível localizar o arquivo 'C:\\Users\\vitor.santos\\OneDrive - MSFT\\INI_DB_AJUSTES\\OneDrive - MSFT\\Desktop\\QC_INI_PRONTO\\qc-ini\\INTERFACEAMENTO_DB… |
| T21 Audit_Log: inativação, reativação, plotagem e manual registrados (ISO 15189 8.4) | **PASS** | {"IMPORTACAO_ABA": 1, "RESULTADO_INATIVADO": 7, "PLOTAGEM_LJ_ALTERADA": 1, "RESULTADO_REATIVADO": 1, "RESULTADO_MANUAL_INCLUIDO": 2, "ATUALIZACAO_FALHOU": 1} |
| T22 Fonte única: nenhum módulo analítico nem fórmula de LJ/Estatística lê DB_SEAC, DB_RECEBIMENTO, manual, inativação, Resultados, Importar ou DB_Resultados | **PASS** | [] |
| T23 VBA compila (Depurar > Compilar) e os formulários que ficaram abrem | **PASS** | {"formularios": "frmAssinar;frmDev;"} |
| T24 Fechar, reabrir e atualizar: camadas, motor, LJ e Estatística funcionando | **PASS** | {"linhas_final": 32033, "corridas_no_LJ": 1, "tempo_1a_s": 66.2, "tempo_apos_reabrir_s": 101.2, "resumo": "Atualização concluída.\r\n\r\nResultados recebidos: 32.033 (novos nesta atualização: 0)\r\nPrincipal - Resultados: 32.033 resultados\r\nQA da integração:… |


## K. Performance

**Bloqueio resolvido — estouro de pilha.** A 1ª versão calculava a corrida com um contador em `List.Generate`;
campos de registro do M são preguiçosos e cada contador apontava para o anterior. Quando a ordenação seguinte
pedia o último valor primeiro, a avaliação aninhava ~44 mil níveis. Bisseção (bancada `testes/depurar_final.py`):

| Versão da DB_CQ_FINAL | 2 mil | 10 mil | 20 mil | 44.392 (Hema completa) |
|---|---|---|---|---|
| v1 — contador em List.Generate | 11,6 s | — | **estouro de pilha** (53 s) | **estouro** (passo `Ordem2`, 169 s) |
| v2 — segmentos + List.Combine | 9,2 s | 49 s | 107 s | ~240 s (projetado) |
| v3 — índice + FillDown, tabela estreita, 1 buffer | — | 7,7 s | 19,0 s | 76,9 s |
| v4 — colunas deslocadas, sem buffer largo | — | 13,2 s | — | 99,3 s |
| **v5 — junções planas + 1 buffer + colunas deslocadas (final)** | — | **8,1 s** | — | **55,7 s** |

Invariantes conferidas em todas as medidas: IDs únicos, 0 resultados sumidos, 0 chaves do motor com mais de um
participante, 589 retransmissões e 93 sem valor (Hema completa) — o resultado lógico não mudou entre versões.
`Table.Buffer` só onde há reuso medido (R1, U, e as três tabelas estreitas de corrida).

**Botão ATUALIZAR DADOS (refresh síncrono das 4 camadas + motor), Hematologia completa, instalação limpa:**
DB 4,9 s · DB_ORGANIZADO 6,9 s · Principal - Resultados 52,1 s · QA_INTEGRACAO 17,7 s · motor (LJ, Westgard,
Estatística) 6,5 s — **~88 s**. Volume: 44.392 resultados × 39 colunas; DB 15 colunas; DB_ORGANIZADO 1.413 × 39.
Máquina: 4 núcleos, 8 GB; a tarefa agendada do DB_SEAC de produção abre Excel a cada 2 min no mesmo
computador e explica a variação observada entre rodadas (52–103 s na camada final).

**Bioquímica (instalação em produção):** DB 7,7 s · DB_ORGANIZADO 6,7 s · Principal - Resultados 47,6 s ·
QA_INTEGRACAO 12,6 s · motor 3,5 s — **~78 s**. Volume: 32.033 resultados × 39 colunas; DB 18 colunas;
DB_ORGANIZADO 4.594 × 45. No QA final, cada ATUALIZAR DADOS completo levou 63–147 s (Hema) e 63–103 s (Bio),
conforme a carga da máquina.

## L. Git

- **Branch:** `main` (produção). A linha `integracao-db-seac` foi incorporada à `main` por fast-forward e encerrada;
  a partir de 02/10/2026 o trabalho segue direto na `main`.
- **Checkpoint antes de qualquer alteração:** tag `checkpoint-pre-integracao-dbseac` (`24f2aaa`).
- **Commit da entrega:** "DB_SEAC vira a porta de entrada e a Principal - Resultados a fonte única (ADR-057)" — o
  commit que contém este relatório (`git log -1 -- _arquitetura/integracao_dbseac/RELATORIO_ADR057.md`): 54 arquivos —
  `QC_Hematologia.xlsm`, `QC_Bioquimica.xlsm`, `.gitignore` e a pasta `_arquitetura/integracao_dbseac/` (consultas M,
  VBA gerado, instaladores, testes, evidências, ADR e este relatório).
- **Segredo fora do repositório:** a senha de administração do DB_SEAC não está em nenhum arquivo nem no histórico;
  `alterar_db_seac.py` lê `DB_SEAC_SENHA` ou pede no terminal.
- **Fora do Git (o remoto é público):** `INTERFACEAMENTO_DB_COPIA/` (cópia do DB_SEAC com dados do laboratório) e os
  backups `_backup_pre_integracao_2026-10-02_1504/` (arquivos de produção antes da instalação, SHA-256 conferido) —
  ambos no `.gitignore`.
- **Push:** `main` enviada ao `origin` após a aprovação do QA pelo usuário (02/10/2026).
- Reversão completa, se necessário: copiar os dois .xlsm de `_backup_pre_integracao_2026-10-02_1504/` de volta para a
  raiz, ou `git checkout checkpoint-pre-integracao-dbseac -- QC_Hematologia.xlsm QC_Bioquimica.xlsm`.

---

## Checklist de conclusão

| Item | | Item | |
|---|---|---|---|
| DB_SEAC auditado | ✅ | Estatística usa apenas DB_CQ_FINAL | ✅ |
| DB_RECEBIMENTO (aba DB) funcionando | ✅ | Estatística só com PARTICIPA_ESTATISTICA = SIM | ✅ |
| DB_ORGANIZADO funcionando | ✅ | Bloco de repetições por analito | ✅ |
| Tabela de inativação (aba Inativar) | ✅ | Levey-Jennings usa apenas DB_CQ_FINAL | ✅ |
| Soft-delete | ✅ | Westgard não recebe inativados | ✅ |
| ID estável | ✅ | Refresh aguarda conclusão real (síncrono) | ✅ |
| Checkbox nativo funcionando | ✅ | Nenhum refresh prematuro | ✅ |
| Default do checkbox = SIM | ✅ | Stack Overflow resolvido | ✅ |
| Comentário técnico obrigatório (E01) | ✅ | Teste de escala aprovado | ✅ |
| Reabilitação | ✅ | Refresh repetido aprovado | ✅ |
| Tabela manual (aba Digitar Resultados) | ✅ | Fechamento/reabertura aprovado | ✅ |
| IDs MAN_ | ✅ | Aba LEGADO (EQC_Dados) removida | ✅ |
| Origem automática/manual | ✅ | Aba Importar removida | ✅ |
| Duplicidades controladas | ✅ | Aba Resultados removida | ✅ |
| DB_CQ_FINAL (aba Principal - Resultados) | ✅ | Rep 1, Rep 2, Rep 3 removidas | ✅ |
| Uma linha = um resultado | ✅ | Nenhuma dependência residual | ✅ |
| Gerada pelo Power Query | ✅ | Regressão completa aprovada | ✅ |
| Conserva inativados | ✅ | Abas e botões com o mesmo nome (DB · Inativar · Digitar Resultados · Principal - Resultados) | ✅ |
| PARTICIPA / REGISTRAR / TIPO_PLOTAGEM | ✅ | X vermelho, não conectado, sem impacto estatístico | ✅ |

---

## Uso diário

1. **⟳ ATUALIZAR DADOS** (Painel ou qualquer aba de dados): traz o DB_SEAC para a aba **DB**, refaz
   **DB_ORGANIZADO**, **Principal - Resultados** e o **QA_INTEGRACAO**, nesta ordem e esperando cada uma terminar;
   só então atualiza Levey-Jennings, Westgard e Estatística. A mensagem final traz o QA (erros/alertas).
2. Resultado de controle que não deve entrar na estatística → aba **Inativar**: cole o ID_REGISTRO (coluna A da
   Principal - Resultados). **REGISTRAR - LJ** já vem marcado (aparece como X vermelho no gráfico); desmarque se
   não quiser vê-lo no gráfico. Escreva a justificativa em **COMENTARIOS_TECNICOS** (obrigatória). Clique
   ATUALIZAR DADOS. Para desfazer, apague a linha e atualize.
3. Interfaceamento caiu → aba **Digitar Resultados**: DATA, HORA, LOTE, NÍVEL, ANALITO, RESULTADO (o ID MAN_ nasce
   sozinho). ATUALIZAR DADOS. A coluna STATUS mostra se entrou (ATIVO) ou se o mesmo resultado já veio pelo
   interfaceamento (CONFLITO_MANUAL).
4. Bioquímica: escolha o equipamento no Painel (L4 — DIMENSION 1 / DIMENSION 2).

## Pendências que dependem do laboratório (não são defeito da integração)

- **Lotes recebidos sem cadastro** (QA I03): Hematologia 528211, 533811, 536411, 602911, 608511, 614011, 619611,
  625311; Bioquímica 8977 e os lotes de HbA1c/PCR/urina. Os resultados estão na Principal - Resultados; para vê-los
  no Painel, cadastre o lote (Configuração) e a média/DP (Média/DP).
- **Média/DP do lote 522611 (Hematologia)** gera violação em quase todas as corridas do período — revisar os
  parâmetros cadastrados.
- **DB_SEAC de produção** (pasta de rede, atualizado pela tarefa agendada) ainda é v2.3, sem `id` do SIPEC. Aplicar
  `alterar_db_seac.py` nele (já validado na cópia) antes de trocar `CAMINHO_DB_SEAC` na aba Cfg_Integracao.
  Enquanto isso, a integração lê a cópia `INTERFACEAMENTO_DB_COPIA\DB_SEAC.xlsm`.
- **Power BI:** `BI_Data` continua gerada por `mBI.AtualizarBIData` (agora sobre a mesma fonte). Recomendado
  apontar o Power BI direto para a tabela `tblCQ_Final`, que já tem o contrato completo.
