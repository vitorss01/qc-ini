# Reteste 05: Usuário Simulado, Bioquímica, frente F1

Data: 15/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (1 ocorrência: OC-0071) · Produção: entrega 42.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção atual (só copiada) | `QC_Bioquimica.xlsm` sha256 curto **b20fe00f7e51**, igual antes e depois da execução |
| Base de comparação (antes da entrega 42) | `_backup_pre_integracao_2026-11-15_0011/QC_Bioquimica.xlsm` sha256 curto **8ec815ab653d** (é a entrega 41, a mesma que o reteste 04 testou), igual antes e depois |
| Cópias (novas, nenhuma salva, apagadas no fim) | `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Bioquimica_F1_base41.xlsm` e `Bioquimica_F1_nova.xlsm` |
| Excel | um por vez, abertos e encerrados pelo `qa_final.QA`: PID 29768 (base), depois PID 30020 (nova). O Excel 13084, que já estava aberto antes (tarefa agendada), não foi tocado. Nenhuma captura de tela |
| Script | `_arquitetura/agentes/ferramentas/reteste05_usuario_f1.py` (documentado no cabeçalho). Saída: `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Bioquimica_F1.json/.log` |
| Dados | nenhum dado de teste gravado: só filtros do Painel. A tblCQ_Final é **a mesma nas duas cópias** (35.906 linhas, hash do Value2 `471c3b76f80f008c` nas duas) |

**Caminho usado.** É o do usuário. Primeiro, as células de filtro do Painel (L4, C3, H3, K3, K4, G3/G4), escritas com
os eventos ligados. Depois, a aba Estatística aberta (`Activate`). Em cada cenário li:
- da Estatística: `Formula`, `Text`, `Value2` e `NumberFormat` de A4:J5, o texto das linhas 1–12 (busca de "1900",
  "yyyy" e "Ano 0") e o `Value2` de A1:AU200 (a tabela inteira);
- do Painel: G3/G4/K3/K4/H3 e o `engPainel`.

A mesma sequência rodou na base (entrega 41) e na nova (entrega 42). A comparação foi feita célula a célula, na
precisão total do `Value2`, com tolerância **zero**: é o mesmo motor sobre os mesmos dados.

**Artefato da automação (não é defeito do QC_INI):** as datas que escrevi no G3/G4 pelo COM nos cenários S7 e S8
entraram com 03:00 (Value2 `…,125`), por causa do fuso. A entrada foi a mesma na base e na nova, então a comparação
continua válida. Um usuário que digita a data grava 00:00.

## Resultado por ocorrência

| OC | Estado | Evidência (`Bioquimica_F1.json`) |
|---|---|---|
| OC-0071 | **RESOLVIDA** | Detalhe nos três blocos abaixo |

**Fluxo original, "Tudo".** Cenário S1: Painel em DIMENSION 1 · Glicose · lote 8974 · K4 "Tudo (sem filtro)", com G3/G4
vazios e K3 = 2026.

| Célula | Base (entrega 41) | Nova (entrega 42) |
|---|---|---|
| B5 | "00/01/1900" (Value2 0, formato `dd/mm/aaaa`) | **""** (Value2 0, formato `dd/mm/aaaa;;`) |
| D5 | "00/01/1900" | **""** |
| D4 | "2026" | **""**; fórmula `=IF(OR(Painel!K3="",AND(Painel!$G$3="",Painel!$G$4="")),"",Painel!K3)` |

O H5 continua "histórico inteiro do lote" e o J5 continua vazio. O texto das linhas 1–12 não tem nenhum "1900" (a base
tinha). A fórmula do B5/D5 (`=Painel!G3/G4`) e o Value2 0 ficaram intactos, como o ADR-095 §3 descreve.

**Fluxo original, período.** Cenário S2: K3 2026 + K4 "Set".
- B5 = "01/09/2026" e D5 = "30/09/2026" (dd/mm/aaaa), iguais ao Painel G3/G4.
- D4 = "2026", e o Painel K3 é 2026.
- H5/J5 = 01/09/2026 a 30/09/2026.

**Números.** Em todos os 9 cenários, a única célula diferente entre base e nova em A1:AU200 é o **D4** nos cenários
"Tudo" (S1, S5, S9), que passa de 2026 para "". Nos outros 6 cenários há 0 diferenças. Em todos os cenários o
`engPainel` e o Painel são iguais. O `n`, a média, o DP, o CV, o bias, o ET, o Sigma e os status da Estatística são
idênticos bit a bit aos de antes da entrega. A mudança foi só de exibição.

### Fluxo vizinho

| Cenário | Painel | Estatística na nova (D4 · B5 · D5 · H5–J5) | Diferenças com a base |
|---|---|---|---|
| S3 trimestre | K3 2026, K4 T3 (01/07–30/09/2026) | 2026 · 01/07/2026 · 30/09/2026 · 01/07/2026 a 30/09/2026 | 0 |
| S4 ano inteiro | K4 "Ano inteiro" (01/01–31/12/2026) | 2026 · 01/01/2026 · 31/12/2026 · idem | 0 |
| S5 todos os lotes | H3 vazio (ADR-086) + "Tudo" | "" · "" · "" · "histórico inteiro do lote" | só D4 (2026 → "") |
| S6 todos os lotes | H3 vazio + Out/2026 | 2026 · 01/10/2026 · 31/10/2026 · idem | 0 |
| S7 De/Até atravessando o ano | G3/G4 01/12/2025–31/01/2026 → K4 "Personalizado", K3 vazio | "" · 01/12/2025 · 31/01/2026 · idem; nenhum "Ano 0" | 0 |
| S8 só o De | Set/2026 e depois Até apagado e De 01/09/2026 → K4 "Personalizado", K3 2026 | 2026 · 01/09/2026 · **""** · H5 01/09/2026, **J5 "00/01/1900"** | 0 (a base tinha "00/01/1900" no D5 **e** no J5) |
| S9 volta a "Tudo" | depois do personalizado | "" · "" · "" · "histórico inteiro do lote" (limpou de novo) | só D4 |

Também conferi:
- nenhum "yyyy" literal nas linhas 1–12 da Estatística, inclusive E8:F9 e C10, que o ADR-095 §5 (OC-0072) cita, em
  nenhum cenário;
- nenhum "Ano 0" em nenhum cenário.

Não há efeito colateral: nos vizinhos, nada mudou além do que a correção pretendia. O S8 tem um resto do mesmo tipo de
defeito, que já existia antes da entrega 42 (R5-01).

## Ocorrências novas

### R5-01: Estatística J5 mostra "00/01/1900" quando o Painel tem só o De (Até vazio)

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela/módulo:** aba Estatística, cabeçalho "PERÍODO ANALISADO",
  J5 (`=IF(AND(Painel!$G$3="",Painel!$G$4=""),"",JanelaFim($B$4,$D$4,$F$4,$B$5,$D$5))`, formato `dd/mm/aaaa`).
- **Ação:** no Painel, apagar o Até (G4) e deixar só o De (01/09/2026). O K4 vira "Personalizado". Depois, abrir a
  Estatística.
- **Esperado:** J5 vazio, ou o fim real da janela (o último dado, por exemplo). Nunca uma data sem sentido.
- **Observado:** D5 vazio (o `;;` do ADR-095 resolveu esse), H5 01/09/2026 e **J5 "00/01/1900"** (Value2 0). Na base
  (entrega 41), D5 e J5 mostravam "00/01/1900". Não foi a entrega 42 que criou isso: é o mesmo defeito da OC-0071 numa
  célula vizinha, que a correção não cobriu (a condição `SEM_DEATE` só trata De e Até vazios ao mesmo tempo).
- **Dificuldade:** quem confere ou imprime a Estatística lê "até 00/01/1900" e não sabe até quando os dados foram
  considerados.
- **Consequência:** só de leitura. Os números são iguais aos da base (0 diferenças no S8).
- **Sugestão:** aplicar ao J5 (e ao H5, para o caso simétrico "só o Até") o mesmo formato `dd/mm/aaaa;;` ou uma
  condição por célula (`Painel!$G$4=""`), sem mexer no `JanelaFim`. Outra opção é o Painel avisar quando só um dos
  dois (De ou Até) estiver preenchido.
- **Evidência:** `Bioquimica_F1.json`, `NOVA.cen.S8_so_De.hdr.J5` (txt "00/01/1900", v2 0.0, fmt `dd/mm/aaaa`) e
  `BASE.cen.S8_so_De.hdr` (D5 e J5 "00/01/1900").
- **Não executado:** o caso simétrico (só o Até preenchido).

## Dúvidas de regra de negócio

1. Com só o De preenchido, o motor usa "do De até o último dado"? O J5 com Value2 0 não diz. Vale definir o que o
   cabeçalho deve mostrar nesse caso (vazio, "até o último dado" ou a data do último resultado).

## Impressões de uso (Relatório A)

- Em "Tudo", o cabeçalho da Estatística ficou coerente: "Ano" e De/Até vazios ao lado de "histórico inteiro do lote".
  Quem chega pela dica do aviso de 180 corridas (OC-0012) não vê mais a data de 1900.
- Com o mês, o trimestre ou o ano, as datas aparecem em dd/mm/aaaa, iguais ao Painel. Atravessando o ano, o Ano fica
  vazio, o que é coerente com o K3 vazio do Painel.
- O caso "só o De" ainda mostra "00/01/1900" no fim da janela (R5-01). É pouco provável na rotina, mas acontece quando
  alguém apaga só uma das datas.

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste05_usuario_f1.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-15_reteste05\Bioquimica_F1.json` e `.log`.
