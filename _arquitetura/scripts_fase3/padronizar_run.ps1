# padronizar_run.ps1 - o identificador da corrida chama-se RUN em toda a interface
#
# Itens 6.2 e 6.3 do QUALITY_GATE.
#
# O MODELO, e o que dele decorre
#
#   corrida = o EVENTO (uma execucao do equipamento)
#   RUN     = o IDENTIFICADOR imutavel desse evento
#
# Disso decorre o que muda e o que NAO muda. A palavra "corrida" nao e um erro
# a ser eliminado: "Lancar corrida", "Registro de corridas" e "DataCorrida"
# estao CERTOS -- falam do evento. Errado e chamar o IDENTIFICADOR por outro
# nome. "Seq" nao existe desde a Fase 1 e sobreviveu em tres rotulos.
#
# A ARMADILHA, e a razao de este script nao ser um localizar-e-substituir
#
#   Audit_Legenda!B22 = "...o valor ORIGINAL esta sempre em Seq = 1."
#
# Aqui "Seq" e OUTRO conceito: a sequencia de AUDITORIA (1a, 2a, 3a vez que o
# resultado foi alterado), produzida por mAuditoria.ProximoSeq. Nao tem relacao
# com a corrida. Um replace global de "Seq"->"RUN" escreveria "o valor ORIGINAL
# esta sempre em RUN = 1" -- uma frase que inverte o significado da trilha de
# auditoria justamente na aba que a explica ao auditor. Fica como esta.
#
# SEGURANCA DA RENOMEACAO: verificado antes de escrever que nenhuma formula
# casa com o literal "Seq"/"Corrida" (nenhum MATCH/PROCV por cabecalho) e que
# nenhum VBA compara com essas strings. Sao rotulos visuais.
#
# IDEMPOTENTE: cada alvo aceita tanto o texto antigo quanto o ja padronizado.
# Texto DIFERENTE dos dois para o build -- e sinal de que a planilha mudou e a
# substituicao precisa ser reavaliada, nao aplicada as cegas.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\padronizar_run.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

# Acentos por codigo: o .ps1 e lido como ANSI e literal acentuado chegaria
# corrompido ao Excel. [string] em cada um porque [char]+[char] e SOMA NUMERICA.
$i_ = [string][char]0x00ED   # i agudo
$u_ = [string][char]0x00FA   # u agudo
$ni = 'n' + $i_ + 'veis'     # "niveis"

# ---------------------------------------------------------------------------
# Alvos: Aba, Celula, texto ACEITO (antigo), texto FINAL
# ---------------------------------------------------------------------------
$ALVOS = @(
    @{
        Aba = 'DB_Resultados'; Cel = 'A2'
        De  = 'Uma linha por resultado. Seq = n' + $u_ + 'mero da corrida (mesmo Seq nos ' + $ni + ' da mesma corrida).'
        Para = 'Uma linha por resultado. RUN = identificador da corrida (mesmo RUN nos ' + $ni + ' da mesma corrida).'
    },
    @{ Aba = 'Calc'; Cel = 'B2'; De = 'Seq'; Para = 'RUN' },
    @{ Aba = 'RegistrosStore'; Cel = 'D1'; De = 'Seq'; Para = 'RUN' }
)

$EIXO_DE = 'Corrida (RUN)'
$EIXO_PARA = 'RUN'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$erros = New-Object System.Collections.ArrayList
$feitos = New-Object System.Collections.ArrayList
$jaOk = New-Object System.Collections.ArrayList

try {
    $estruturaProtegida = $wb.ProtectStructure
    if ($estruturaProtegida) { $wb.Unprotect($SENHA) }

    # ------------------------- 6.2 rotulos de celula -------------------------
    $abas = @($wb.Worksheets | ForEach-Object { $_.Name })
    foreach ($alvo in $ALVOS) {
        if ($abas -notcontains $alvo.Aba) {
            [void]$erros.Add("aba $($alvo.Aba) nao existe")
            continue
        }
        $ws = $wb.Worksheets($alvo.Aba)

        $reprot = $false
        if ($ws.ProtectContents) {
            try { $ws.Unprotect($SENHA) } catch { }
            if ($ws.ProtectContents) { try { $ws.Unprotect() } catch { } }
            if ($ws.ProtectContents) {
                [void]$erros.Add("$($alvo.Aba) nao destravou com a senha do projeto")
                continue
            }
            $reprot = $true
        }

        $cel = $ws.Range($alvo.Cel)
        $atual = "$($cel.Value2)"

        if ($atual -eq $alvo.Para) {
            [void]$jaOk.Add("$($alvo.Aba)!$($alvo.Cel)")
        }
        elseif ($atual -eq $alvo.De) {
            # .Value2 e nao .Formula: Formula faria o Excel INTERPRETAR o texto.
            $cel.Value2 = $alvo.Para
            # Ler de volta: atribuicao COM que nao liga nao levanta erro.
            $lido = "$($cel.Value2)"
            if ($lido -ne $alvo.Para) {
                [void]$erros.Add("$($alvo.Aba)!$($alvo.Cel) nao gravou: ficou '$lido'")
            }
            else {
                [void]$feitos.Add("$($alvo.Aba)!$($alvo.Cel): '$atual' -> '$($alvo.Para)'")
            }
        }
        else {
            [void]$erros.Add("$($alvo.Aba)!$($alvo.Cel) tem texto inesperado (nao e nem o antigo nem o novo): '$atual'")
        }

        if ($reprot) { try { $ws.Protect($SENHA) } catch { } }
    }

    # ---------------- 6.3 titulo do eixo X (categorias) ----------------
    # Percorre TODO grafico embutido de TODA aba: a contagem varia por produto
    # (Bioquimica 2 niveis, Hematologia 3) e nao deve ser fixada aqui.
    $nGraf = 0
    foreach ($ws in $wb.Worksheets) {
        foreach ($co in $ws.ChartObjects()) {
            $nGraf++
            $ch = $co.Chart
            try {
                $ax = $ch.Axes(1)          # 1 = xlCategory
                if (-not $ax.HasTitle) {
                    [void]$erros.Add("$($ws.Name)!$($co.Name): eixo X sem titulo")
                    continue
                }
                $t = "$($ax.AxisTitle.Text)"
                if ($t -eq $EIXO_PARA) {
                    [void]$jaOk.Add("$($ws.Name)!$($co.Name) eixo X")
                }
                elseif ($t -eq $EIXO_DE) {
                    $ax.AxisTitle.Text = $EIXO_PARA
                    $lido = "$($ax.AxisTitle.Text)"
                    if ($lido -ne $EIXO_PARA) {
                        [void]$erros.Add("$($ws.Name)!$($co.Name): eixo X nao gravou, ficou '$lido'")
                    }
                    else {
                        [void]$feitos.Add("$($ws.Name)!$($co.Name) eixo X: '$t' -> '$EIXO_PARA'")
                    }
                }
                else {
                    [void]$erros.Add("$($ws.Name)!$($co.Name): titulo de eixo X inesperado '$t'")
                }
            }
            catch {
                [void]$erros.Add("$($ws.Name)!$($co.Name): falha no eixo X - $($_.Exception.Message)")
            }
        }
    }
    if ($nGraf -eq 0) { [void]$erros.Add('nenhum grafico embutido encontrado') }

    if ($estruturaProtegida) { $wb.Protect($SENHA, $true, $false) }

    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Padronizacao RUN incompleta: $($erros.Count) problema(s). Nada foi salvo."
    }

    $wb.Save()
}
finally {
    try { if ($erros.Count -gt 0) { $wb.Close($false) } else { $wb.Close($true) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}

"graficos inspecionados: $nGraf"
"alterados ($($feitos.Count)):"
if ($feitos.Count -eq 0) { "  (nenhum)" } else { $feitos | ForEach-Object { "  $_" } }
"ja padronizados ($($jaOk.Count)): $(if ($jaOk.Count) { $jaOk -join ', ' } else { '-' })"
"preservado de proposito: Audit_Legenda!B22 (Seq = sequencia de AUDITORIA, outro conceito)"
"Salvo: $Workbook"
