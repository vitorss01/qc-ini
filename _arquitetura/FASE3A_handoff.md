# FASE 3A — Relatório técnico de handoff

**Emitido por:** Agente 1 (gestor / diagnóstico)
**Destinatário:** Agente 2 (correção)
**Data:** 01/08/2026
**Status do projeto:** 1 bloqueador aberto, escopo congelado

---

## 0. REGRA DE OURO DESTA ETAPA

**Nenhuma funcionalidade nova.** Nem gráficos, nem `frmNaoConforme`, nem eixo X, nem
padronização textual. Apenas estabilizar. Se surgir ideia, registre no fim deste arquivo.

**Proibido `On Error Resume Next` como correção.**
**Proibido trocar estrutura de dados por suspeita** — só com causa comprovada.

---

## 1. ESTADO DOS ARQUIVOS

| Arquivo | Estado | Ação |
|---|---|---|
| `QC_Hematologia.xlsm` | ⚠️ Motor instalado, **congela** em `RegistrarEventosWestgard` | Alvo do trabalho |
| `QC_Bioquimica.xlsm` | ✅ Fase 2, funcional | **NÃO TOCAR** |
| `QC_Imunologia.xlsm` | ✅ Fase 2, funcional | **NÃO TOCAR** |

Backups: `_backup_pre_fase1/`, `_backup_pre_fase2/`, `_backup_pre_fase3/` (todos válidos).
Trabalhe sobre **cópia**. Restauração: `cp _backup_pre_fase3/QC_Hematologia.xlsm .`

---

## 2. O QUE JÁ ESTÁ PROVADO — não reinvestigar

### 2.1 O projeto COMPILA
Provado por execução bem-sucedida de `D_Alvo` e `D_Calc` via `Application.Run`.
Em VBA não existe API de "Compile Project"; a compilação é forçada ao executar
qualquer procedimento. Se houvesse erro de compilação, **nenhum** procedimento rodaria.

### 2.2 `AlvoAnalito` está correta
```vba
Public Sub AlvoAnalito(ByVal analito As String, ByVal nivel As Long, _
                       ByRef alvoMedia As Double, ByRef alvoDP As Double, ByRef etp As Double)
```
- É **Sub** (não Function) → a sintaxe `AlvoAnalito a, b, c, d, e` está **correta**
- 5 call sites, todos com 5 argumentos e tipos compatíveis (linhas 382, 622, 712, 779, 886)
- Teste isolado: `D_Alvo: OK media=3,0038 dp=0,0529 etp=7`

**A linha `AlvoAnalito analitoN, nivelN, aM, aS, etp` NÃO é a causa.** O VBA reporta a
linha onde a execução estava parada, não onde está o defeito.

### 2.3 Análise estática — limpa
- Sem procedimento chamado e não definido
- Sem declaração de módulo após procedimento
- Sem nome duplicado entre módulos
- Sem `ScreenUpdating`/`EnableEvents`/`Calculation` deixados alterados
- Sem variável sombreando função nativa (o Erro 9 anterior — array `val()` vs `Val()` — já corrigido)

**Achado menor (não bloqueador):** `mLotes.bas:45,69` declara `Dim stR As Worksheet`.
`stR` colide com a função nativa `Str()`. Inócuo hoje (usado só como objeto), mas
renomear para `wsReg` elimina risco futuro.

### 2.4 Motor estatístico — validado matematicamente
Confere com referência independente em Python, 6 casas decimais:
média 14,000000 · DP 3,162278 · CV 22,587698 · Bias 40,000000 · ET 21,500000 · Sigma 5,000000 · Z 1,000000

Westgard: **9/9 asserções**, incluindo o caso sutil (níveis em lados opostos → R4s, não 22s).
Elegibilidade CLSI EP05/C24: 8 estados, desconhecido nunca entra no cálculo.
`AtualizarCalc`: **funciona** (`D_Calc: OK Calc`).

---

## 3. O BLOQUEADOR

**`RegistrarEventosWestgard` congela.** Não lança erro — trava. Excel headless fica preso
e exige encerramento forçado.

### Já tentado, sem resolver
1. `col(j)(0)` → variável intermediária (indexação encadeada de membro padrão)
2. `mAgg` definido no início e em todas as saídas (evitar re-disparo em cascata)
3. Buffer e `ClearContents` de 20.000 → 5.000 linhas

### Estrutura atual da rotina
```
Entrada:  mDB (snapshot do banco), lote ativo
Saída:    aba Eventos_Westgard (A4:H) + mAgg (Dictionary de agregados)

Estágios:
  1. Varredura de mDB      -> grupos: Dictionary(chave -> Collection de Array(RUN,valor,data))
  2. Por grupo: deduplica  -> arrays rs()/ys()/ds()
  3. Ordenação por inserção
  4. AlvoAnalito + z-scores
  5. AvaliarWestgard1N
  6. Acumula eventos em ev()
  7. Escrita em bloco + Set mAgg
```

---

## 4. TAREFA DO AGENTE 2

### 4.1 PROVAR o ponto de congelamento (obrigatório antes de corrigir)

Bissecção com marcador de progresso que **sobreviva ao encerramento forçado**.

⚠️ **Armadilha já cometida duas vezes:** literais de string do VBA **não usam escape de
barra invertida**. Use `"C:\pasta\trace.txt"`, nunca `"C:\\pasta\\trace.txt"`. O `Open`
falha, o erro estoura dentro do handler e o Excel trava — mascarando o problema real.

Método sugerido:
```vba
Public Sub TR(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open "C:\caminho\trace.txt" For Append As #f     ' caminho SEM escape
    Print #f, Format(Now, "hh:nn:ss") & " " & s
    Close #f
End Sub
```
Instrumente os 7 estágios acima. Rode com timeout curto, mate o Excel, leia o trace.
**O último estágio registrado indica onde travou.**

### 4.2 Descartar explicitamente antes de culpar a estrutura
- Recálculo automático disparado a cada escrita (`Application.Calculation`)
- Eventos de planilha em cascata na escrita de `Eventos_Westgard`
- Recursão involuntária `AgregadoWestgard` ↔ `RegistrarEventosWestgard`
- `Dictionary` de `Collection` — **só depois de eliminar os anteriores**

### 4.3 Corrigir
Somente com causa comprovada. Se a evidência apontar a estrutura aninhada, substitua por
arrays planos com índice. Documente a evidência que justificou a decisão.

---

## 5. TAREFA DO AGENTE 3 (teste)

Rodar, **sem intervenção manual**, e reportar ao Agente 1:

| # | Critério | Esperado |
|---|---|---|
| 1 | `AtualizarEstatistica` completa | < 30 s, sem travar |
| 2 | Primitivas vs. referência Python | idênticas em 6 casas |
| 3 | Westgard | 9/9 asserções |
| 4 | Elegibilidade | 6/6 estados |
| 5 | `Eventos_Westgard` populada | nº de eventos > 0 e coerente |
| 6 | Varredura de erros de fórmula | **NENHUM** |
| 7 | Fase 1 intacta | RUN, exclusão lógica, troca de lote, login |
| 8 | Fase 2 intacta | 3 formulários carregam, upsert não duplica |
| 9 | Estado global | `ScreenUpdating`/`EnableEvents`/`Calculation` restaurados |

Scripts prontos: `_arquitetura/scripts_fase1/`, `scripts_fase2/`.
Harness de diagnóstico: envolver a chamada em `On Error GoTo eh` e devolver
`Err.Number & Err.Description` — foi o que provou que `AlvoAnalito` estava sã.

---

## 6. PROTOCOLO OPERACIONAL (todos os agentes)

Antes de **cada** execução:
```powershell
Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery" -Recurse -Force -EA SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"    -Recurse -Force -EA SilentlyContinue
Get-ChildItem "$env:APPDATA\Microsoft\Excel" -Recurse -Include *.xlsb -EA SilentlyContinue | Remove-Item -Force
```
Motivo: kills forçados enfileiram Recuperação de Documentos; na abertura seguinte o Excel
tenta exibir o painel de recuperação e **trava sob `/automation`** (`CO_E_SERVER_EXEC_FAILURE`).
Isso já custou uma sessão inteira de diagnóstico.

⚠️ **AutoSave do OneDrive está ativo** nesta pasta e já persistiu estado parcial após falha.
Considere desativá-lo para estes arquivos.

---

## 7. CRITÉRIO DE ENCERRAMENTO DA FASE 3A

Os 9 critérios da seção 5, verdes, em execução limpa. Só então liberar para:
`frmNaoConforme` → aba Registros · eixo X com RUN alinhado · padronização "Corrida"→"RUN" ·
documentação com rastreabilidade normativa (fórmula, entradas, referência CLSI por cálculo).

---

## 8. DÍVIDA TÉCNICA REGISTRADA (não corrigir agora)

1. `mLotes.bas` — `Dim stR As Worksheet` colide com `Str()` nativo → renomear para `wsReg`
2. `Calc` — referências absolutas (`$BT$1`/`$BU$1`) frágeis a inserção de coluna
3. `LiberStore` — gravação preguiçosa (só em `BeforeSave`), risco em fechamento por crash
4. 15 módulos de classe de planilha, só 5 com código
5. Projeto VBA sem senha de proteção

---

# FASE 4 — MODO ENGENHARIA (hardening, QA, UX)

**Status:** aguardando. **NÃO INICIAR** enquanto o bloqueador da seção 3 estiver aberto.

## Critério de entrada (gate obrigatório)

O ciclo de hardening só começa quando os **9 critérios da seção 5** estiverem verdes.
Rodar QA sobre um sistema que congela produz relatório sobre código que não executa —
e o próprio pipeline exige "nunca avançar sem validar a etapa anterior".

## Papéis, aplicados a ESTE sistema

| Papel | Foco específico já mapeado |
|---|---|
| **1 · Arquiteto** | Acoplamento `Calc` ↔ Painel via posição de coluna (`CF0 + t*NFD`); os 4 armazéns ocultos (`LotesStore`/`LiberStore`/`RegistrosStore`/`Cfg_Status`) seguem o mesmo padrão mas com código duplicado — candidato a rotina única de swap. Dívidas 1–5 acima |
| **2 · Desenvolvedor** | Só executa o aprovado. Regra desta base: **nunca** declarar variável com nome de função nativa (custou o Erro 9) |
| **3 · QA** | Cenários que ainda não existem: DB vazio · 1 único resultado · analito sem média/DP configurados · lote sem nenhum resultado · Status inválido digitado à mão · data futura · RUN duplicado manualmente |
| **4 · UX/UI** | Fluxo diário real: abrir → lançar RUN → importar em massa → excluir → ler gráfico → interpretar Westgard. Medir **cliques por tarefa**. Ponto fraco conhecido: `frmMassa` usa área de colagem, não grade navegável |
| **5 · Auditor** | Rastreabilidade já existe (RUN + exclusão lógica + elegibilidade CLSI). **Lacunas reais:** sem trilha de auditoria (quem/quando/o quê), sem senha no projeto VBA, `Cfg_Status` editável sem registro. Fraude possível hoje: editar `DB_Resultados` direto com Modo Desenvolvedor sem deixar rastro |
| **6 · Estresse** | Alvos: 5.000 RUNs · 40 analitos × 3 níveis · 50 lotes · importação de 500 linhas · `AtualizarEstatistica` repetido. Limites atuais: `DB_Resultados` 15.000 linhas, view 3.000/bloco, eventos 5.000, `Calc` 180 RUNs |

## Pipeline com portões

```
Arquitetura → Desenvolvimento → Compilação(*) → Unitário → Integrado →
Regressão → Desempenho → Usabilidade → Usuário final → Auditor →
Refatoração → nova bateria → relatório → nova revisão → (repetir)
```

(*) VBA **não tem** API de "Compile Project". A compilação é forçada executando
qualquer procedimento — se travar, há erro de compilação. Foi assim que o Erro 9 apareceu.

## Congelamento de arquitetura (recomendação aceita)

Após a Fase 4, congelar a arquitetura. Toda funcionalidade nova passa a exigir:
análise de impacto → implementação → testes → regressão → aprovação.
Registrar cada mudança neste diretório.

---

# ESTADO DA SESSÃO (encerrada no limite de contexto)

## Concluído e validado
- **Fase 1** — arquitetura em camadas, RUN, exclusão lógica, Modo Desenvolvedor · `FASE1_documentacao.md`
- **Fase 2** — DB_Resultados como fonte única, view em 3 blocos, 3 UserForms, upsert sem duplicar · `FASE2_documentacao.md`
- **Fase 3 (parcial)** — motor estatístico validado contra referência Python (6 casas),
  Westgard 9/9, elegibilidade CLSI EP05/C24, `mWestgardKnowledge` com Enum e Type,
  aba `Eventos_Westgard`, Erro 9 corrigido

## Aberto
1. 🔴 **`RegistrarEventosWestgard` congela** — único bloqueador (seção 3)
2. ⏸ `frmNaoConforme` → aba Registros — **não iniciado**
3. ⏸ Eixo X com RUN alinhado — **não iniciado**
4. ⏸ Padronização textual "Corrida" → "RUN" — **não iniciado**
5. ⏸ Documentação de rastreabilidade normativa (fórmula + entradas + referência CLSI por cálculo) — **não iniciada**

## Não refazer (já feito e validado)
Fases 1 e 2 completas · validação matemática do motor · 9 asserções de Westgard ·
tabela de elegibilidade · correção do Erro 9 · análise estática (limpa).

## Primeiro comando da próxima sessão
Ler este arquivo, restaurar ambiente limpo (seção 6) e atacar **apenas** a seção 4.1:
provar o ponto de congelamento por bissecção instrumentada.
