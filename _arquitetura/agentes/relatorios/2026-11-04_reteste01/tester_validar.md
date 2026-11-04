# Tester/QA — modo VALIDAR — ADR-088 (reteste01, 04/11/2026)

- **Quem corrigiu:** o agente programador (ADR-088, commits 759b758 + aec2c43, entrega 35). Não fui eu: VALIDAR permitido.
- **Base:** cópias da produção já entregue, nunca aberta para escrita. Bio `06bdd4ec53b7`, Hema `4c0b41f0c0a6`, iguais
  no fim de todas as rodadas. Cópias em `%TEMP%\qcw\auditoria\2026-11-04_reteste01\`, descartadas no fim.
- **Método:** um Excel por vez. Encerrei só os PIDs que eu criei: 3648, 17440, 27036, 17040, 17608, 28216 e os das
  suítes. Na regressão que eu interrompi, encerrei o 10436 (o `qa_incerteza` que meu script tinha aberto). Nenhuma
  captura de tela.
- **Dados de teste:** MOTIVO `Outro: TESTE-AUDITORIA`. Valores 244,222 / 17,433 / 77,111 / 3,111 / −1 / 0, nos IDs
  MAN_0001..0006 da cópia.
- **Scripts (meus):** `ferramentas/valida_adr088.py` (reaproveita `repro_comum.py` do ciclo 01) e
  `ferramentas/regressao_pos_atualizar.py`.
- **Logs:** `valida_adr088_<produto>.json` e `regressao_Bioquimica/*.log` na pasta da cópia.
- **ATUALIZAR:** `mIntegracao.AtualizarSoEntradas` (núcleo do ATUALIZAR, sem ler o DB_SEAC). Na cópia, ele aplica
  IMPLAUSIVEL e a data estrita, como o 1º ATUALIZAR depois da entrega faria.

## Vereditos

| OC | Produto | Veredito | Evidência |
|---|---|---|---|
| OC-0001 | Bio | **VALIDADA** | Ver abaixo |
| OC-0001 | Hema | **VALIDADA** | Ver abaixo |
| OC-0002 | Ambos | **VALIDADA** | Ver abaixo |
| OC-0003 | Hema | **VALIDADA** | Ver abaixo |
| OC-0003 | Bio | **VALIDADA** | Ver abaixo |
| OC-0004 | Bio | **VALIDADA** | Ver abaixo |
| OC-0004 | Hema | **VALIDADA** | Ver abaixo |
| OC-0005 | Ambos | **VALIDADA** | Ver abaixo |
| OC-0006 | Bio | **VALIDADA** | Ver abaixo |
| OC-0006 | Hema | **VALIDADA** | Ver abaixo |
| OC-0009 | Bio | **VALIDADA** | Ver abaixo |
| OC-0013 | Ambos | **VALIDADA** | Ver abaixo |

### OC-0001 — Bio

- "dimension 1" e "glicose" passam a "DIMENSION 1"/"Glicose" já ao digitar.
- Os manuais ganham corrida própria: RUN 26102807 e 26102705.
- A12 = ∅. O real continua no Eng_Saida: RUN 26102701 = 279 e 26102801 = 276 (IDs da data de hoje; o 371977 do
  ciclo 01 não está mais na janela), com Westgard OK antes e depois.
- Painel N2 n 20, média 267,7722 = Estatística = independente (sem caixa e "um por RUN"), com 1e-9.
- 0 chaves (equip, lote, nível, analito, RUN) repetidas em toda a tblCQ_Final.

### OC-0001 — Hema

- "xn1000" e "wbc" passam a "XN1000"/"WBC". Os manuais ganham RUN 26093003 e 26092907. A12 = ∅.
- O real continua no Eng_Saida: 26092901 = 16,94 e 26093001 = 16,84, com Westgard OK antes e depois.
- WBC 619611: Painel n 55/68/51 = Estatística (com B3:D3 = 2026) = independente. As médias são iguais (1e-9).
- 0 chaves repetidas.

### OC-0002 — Ambos

- H7:H8/H9 = `IF(OR($E="",NOT(ISNUMBER($G))),"",…)`.
- H fica "" com G = SEM EP em:
  - Bio: HbA1c 8590, PCR 8353, Proteína total com M4 = Controllab;
  - Hema: RET%, RDW-SD, e RBC/LYMPH# com Controllab.
- Com G numérico, H = 1,65·CV + |G| (1e-9). Exemplos: Proteína total com CAP 6,4486/7,0114; RBC 3,581/2,813/2,812.

### OC-0003 — Hema

Na troca 522611 → 619611 → 522611, B4 acompanha o lote, e C:F batem com o independente pelo B4 e pelo B3:D3 (1e-9)
em todos os passos:

- 522611: n 50/51/51, CV N1 1,2595 %, ET 3,6119 = Painel.
- Ano De/Até 2026: RBC N1 n 54, CV 1,134432 % (independente 1,134432 %).
- Ano 2025: n 0 = independente.

Nenhum número velho ficou na tela. Ver a observação O1.

### OC-0003 — Bio

- 8974 ↔ 8977: o Painel e a Estatística acompanham. Glicose 8977: n 76/76, CV 2,3415/16,7505 iguais (1e-12).
- 8974 com "Tudo": a Estatística dá n 174/161 = independente com o lote inteiro (1e-12). O Painel dá 140/131 por causa
  da janela de 180 corridas, que é a OC-0012, aberta e fora deste lote.

### OC-0004 — Bio

Antes do ATUALIZAR, a cópia tem 56 participantes ≤ 0 (43 < 0) e 0 IMPLAUSIVEL. Depois:

- IMPLAUSIVEL = 56. É o mesmo conjunto da varredura independente, e A11 = 56 com os mesmos IDs.
- Todos com PARTICIPA = NÃO, REGISTRAR_RESULTADO_NO_LJ = NÃO e TIPO_PLOTAGEM_LJ = NAO_PLOTAR.
- 0 participantes ≤ 0. 0 outros IDs mudaram PARTICIPA. 0 RUN mudados.
- Manuais −1 e 0 → IMPLAUSIVEL + A11.
- Glicose DIM1 8977 N2 out/2026: o CV vai de 34,27 % (n 20) para 24,22 % (n 19). Não volta a ~2 % porque o
  BIO-371184 vale **1** (> 0, nível trocado), fora da regra autorizada.
- Simulei a inativação pelo laboratório (parecer G4, ação 1) na cópia: n 18 e CV **2,0849481 %** = critério 2,0849 %,
  igual ao independente (1e-12).

### OC-0004 — Hema

0 IMPLAUSIVEL na produção, como diz o ADR. Os manuais WBC −1 e 0 → IMPLAUSIVEL + A11. WBC N1 fica com n 54 + 1
(só o 28/02 válido entra).

### OC-0005 — Ambos

- 31/02/2026 → MANUAL_INCOMPLETO, com DETALHE "…: DATA", E04, PARTICIPA = NÃO e RUN vazio.
- 28/02/2026 → ATIVO (RUN 26022801). Vale nos dois produtos.

### OC-0006 — Bio

Bilirrubina direta: G = 19,99420975515574 nos dois níveis, AC 13, AD 3. Antes da entrega era 30,7709. A bilirrubina
total ficou em 10,8223, com AC 15.

### OC-0006 — Hema

IG%, IG#, NRBC% e NRBC#, N1–N3: G = "SEM EP" nas 12 linhas.

### OC-0009 — Bio

- B103 = 4 = contagem independente (antes da entrega: 22).
- B98:B102 = 0/0/3/2/53, iguais aos de antes da entrega.
- F98:F100 foi de 46/2/10 para 47/1/10. Atribuí isso à OC-0006: P da Bilirrubina direta muda com o bias novo, e F é
  um COUNTIF de P.
- Critério 2 (Controllab = 62): `NAO_EXECUTADO`. Mudar o L4 da Estatística por COM não trocou o provedor (B103 e G
  ficaram iguais).

### OC-0013 — Ambos

- H3 vazio: G6 = "Bias % do método (CEQ)", I6 = "Sigma do método".
- C:E e H ficam vazios.
- G = Estatística G e I = Estatística L (1e-12). Exemplos:
  - Bio Glicose: 2,7333 / 3,190 e 0,550;
  - Hema WBC: 1,8648 / 2,272, 3,580 e 5,683 (= critério G12).

## Regressão

**Entrega 35 (logs do `entregar.py`, `_entrega_2026-11-03_2123`):** as 28 suítes passaram nos dois produtos. Elas
rodaram sobre o arquivo instalado **antes** do 1º ATUALIZAR.

**Minha regressão sobre a cópia Bio já ATUALIZADA** (56 IMPLAUSIVEL; cópia salva só em %TEMP%):

| Suíte | Resultado |
|---|---|
| `qa_westgard_sequencia` | 29 / 0 |
| `qa_adr088` | 16 / 0 |
| `qa_incerteza` | 15 / 0 |
| `qa_final` | 29 / 0 |
| `qa_rodadas_cv` | **11 / 1: o R00 reprova** |

Sobre o R00:

- O R00 compara com o arquivo de antes da entrega e exige AF..AU idênticos.
- Ele acusa 8 células em 2 linhas: a linha 14 em AF, AG, AK, AL, AM e AQ (CVp 8,893 → 5,979, n 242 → 241, u(Rw)
  9,63 → 7,03) e a linha 15 em AF e AG.
- É o efeito esperado da OC-0004: CVp e u(Rw) leem PARTICIPA. O próprio ADR-088 prevê isso em "Não rodadas".
- **Não é defeito do produto. É uma limitação do oráculo R00 depois do ATUALIZAR** (ver O3).

Sem o ATUALIZAR (primeira rodada, interrompida): `qa_adr088` 16 / 0 e `qa_westgard_sequencia` 29 / 0.

Hema: a regressão depois do ATUALIZAR não foi rodada (`NAO_EXECUTADO`). Na Hema, o ATUALIZAR não muda nada: 0
IMPLAUSIVEL, 0 PARTICIPA alterados e nenhuma linha da Estatística mudou. Valem as 28 suítes da entrega.

## Efeitos colaterais e observações

- **O1 (Hema, UX, Médio, proposta de OC nova).** A Estatística filtra pelo próprio Ano De/Até (B3:D3) e não acompanha o
  lote. Logo depois de trocar para o 619611 com B3:D3 = 2025, ela mostra n 0 enquanto o Painel mostra n 54. O número
  está coerente, e não velho (o defeito da OC-0003 acabou). Mas o usuário vê vazio até ajustar o Ano. Sugestão ao
  UX/CQ: aviso no B4, ou B3:D3 iniciados pelo período do lote.
- **O2 (Bio).** Depois do ATUALIZAR, o resumo Sigma B98:B102 continua 0/0/3/2/53, e só 2 linhas da Estatística mudam
  CVp/u(Rw). O ADR-088 diz que ficaria 6/1/3/7/41. Medido aqui, sobre `06bdd4ec53b7`, não se reproduz. O Orquestrador
  deve pedir ao programador a origem desse número: cópia ou estado diferentes. Não muda os vereditos.
- **O3 (teste).** O R00 do `qa_rodadas_cv` reprova sempre que roda sobre uma cópia já ATUALIZADA, porque compara com o
  arquivo de antes da entrega. Proposta: o R00 exclui as linhas cujo analito × nível tem IMPLAUSIVEL, ou roda só
  antes do ATUALIZAR. Isso fica registrado no próprio teste.
- **O4.** Glicose N2 só fica perto de 2 % depois de o laboratório inativar o BIO-371184 (valor 1). O resto da lista
  da F3 (nível trocado, |z| > 5) continua ativo. É a pendência OC-0004/laboratório, fora do ADR-088 por decisão do
  usuário.
- **Sem efeito colateral encontrado** em:
  - corrida e RUN (0 RUN mudados);
  - PARTICIPA dos não ≤ 0 (0 mudanças);
  - Westgard dos RUN reais (OK = OK);
  - chaves duplicadas (0);
  - Eng_Saida (os valores reais se mantêm);
  - tempo da troca de lote (1,9 a 3,7 s por troca).
