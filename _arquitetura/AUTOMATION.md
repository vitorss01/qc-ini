# QC_INI — Automação diária e alertas

Estado: **arquitetura definida e documentada; nada agendado ainda.** Nenhum
credencial foi armazenada e nenhuma tarefa foi criada sem o seu aval — agendar
envio automático para gestores é ação externa e é sua decisão, não minha.

---

## 1. A cadeia

```
06:00 (parametrizável)
   │
   ├─ 1. Excel abre o QC_INI e roda AtualizarOperacao + AtualizarBIData
   │      → BI_Data reconstruída, flags e nomes coerentes
   │
   ├─ 2. Power BI Service atualiza o dataset (OneDrive/SharePoint, sem gateway)
   │
   ├─ 3. Assinatura de e-mail do Power BI, ou Power Automate, gera o snapshot
   │
   └─ 4. E-mail chega aos gestores com o resumo do dia
```

## 2. Etapa 1 — atualizar a camada de dados

Script sugerido (**não criado ainda** — precisa da sua confirmação sobre horário
e sobre rodar em máquina que fica ligada):

```powershell
# atualizar_bi_diario.ps1  (esboço)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open('<caminho do QC_Bioquimica.xlsm>')
try {
    $xl.Run('AtualizarOperacao')   | Out-Null
    $xl.Run('AtualizarBIData')     | Out-Null
    $r = $xl.Run('ReconciliarComCalc')
    if (($r -split '\|')[1] -ne '0') { throw "BI diverge do motor: $r" }
    $wb.Save()
} finally {
    $wb.Close($true); $xl.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
```

**A reconciliação fica dentro da rotina diária, não só do build.** Se um dia os
números divergirem, é melhor a atualização falhar e alguém investigar do que o
painel ser publicado errado e ninguém perceber.

Agendamento: Agendador de Tarefas do Windows, gatilho diário, horário
parametrizável. **Não fixar o horário no script** — vive na tarefa.

## 3. Etapa 2 — atualizar o dataset

Como o `.xlsm` está no OneDrive, o Power BI Service atualiza pelo conector
OneDrive/SharePoint **sem gateway**. Configurar em: dataset → Configurações →
Atualização agendada, com horário posterior ao da etapa 1 (sugestão: 06:30, para
dar folga).

Se o arquivo for movido para disco local, passa a exigir *On-premises data
gateway (personal mode)*.

## 4. Etapa 3 — o e-mail

Duas opções, e a primeira é claramente melhor para começar:

### A) Assinatura de e-mail nativa do Power BI *(recomendada)*
No relatório publicado: **Assinar por e-mail** → escolher a página
"Qualidade do Dia" → destinatários → frequência diária.

- Não exige Power Automate, nem licença extra além do Pro
- Envia a imagem da página com um link
- **Não guarda credencial em lugar nenhum** — a assinatura vive no Service

### B) Power Automate *(quando precisar de condicional)*
Fluxo agendado: `Atualizar dataset` → `Exportar relatório para PDF` →
`Enviar e-mail (V2)`. Necessário quando o envio for condicional — por exemplo,
"só enviar se houver violação".

Em nenhum dos dois casos senha ou token entra na planilha ou no repositório. A
autenticação é do conector, gerida pelo Microsoft 365.

## 5. Conteúdo do resumo diário

Da página "Qualidade do Dia", já suportada pelas medidas existentes:

| Item | Medida |
|---|---|
| Data | `Data Mais Recente` |
| Resultados analisados | `n Hoje` |
| Controles fora de controle | `Rejeitados Hoje` |
| Regras violadas | `Viol 1_3s`, `Viol 2_2s`, `Viol R_4s` |
| Analitos críticos | `Analitos Críticos` |
| Situação geral | `Status Global` |
| Tendência | `Deslocamento vs Alvo (SD)` |

## 6. Alertas

Alertas do Power BI Service disparam sobre **cartão, KPI ou medidor**, com limiar
numérico. Critérios propostos — todos derivados de meta cadastrada, nenhum
arbitrário:

| Alerta | Condição | Justificativa |
|---|---|---|
| Violação de Westgard | `Viol Total` > 0 no dia | regra violada exige conduta documentada |
| Sigma baixo | `Sigma` < 3 | abaixo de 3σ o desempenho é inaceitável |
| Fora de controle | `% Fora de Controle` > 5% | mais de 1 em 20 corridas rejeitada |
| Deslocamento | `|Deslocamento vs Alvo (SD)|` > 1,5 | sugere erro sistemático |

Sem critério definido, alerta vira ruído e as pessoas param de ler — que é a
única falha pior do que não ter alerta.

## 7. O que ainda depende de você

- [ ] confirmar o horário da rotina diária
- [ ] confirmar em que máquina ela roda (precisa estar ligada)
- [ ] criar a tarefa no Agendador (posso preparar o script quando disser)
- [ ] publicar o relatório no Service e configurar a atualização
- [ ] definir a lista de destinatários e criar a assinatura de e-mail
