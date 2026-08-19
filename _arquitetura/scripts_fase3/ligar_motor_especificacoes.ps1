# ligar_motor_especificacoes.ps1 - inclui o motor formal na cadeia operacional
#
# Nao redireciona nenhuma formula. Na Hematologia, os consumidores existentes
# continuam em Analitos!Q:R; Eng_Especificacoes e atualizado em paralelo e
# preserva o fallback legado ate o cadastro formal estar utilizavel.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Workbook -notmatch '(?i)(_EM_DESENVOLVIMENTO|build_hardening)') {
    throw 'Recusado: integracao somente em clone _EM_DESENVOLVIMENTO ou build_hardening.'
}

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null; Start-Sleep 5 } catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

$salvou = $false
$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }

try {
    if ($wb.ReadOnly) { throw "Somente leitura: $Workbook" }
    $proj = $wb.VBProject
    $operacao = $null
    $espec = $null
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'mOperacao') { $operacao = $c }
        if ($c.Name -eq 'mEspecificacoes') { $espec = $c }
    }
    if ($null -eq $espec) { throw 'mEspecificacoes nao esta instalado.' }
    if ($null -eq $operacao) { throw 'mOperacao nao esta instalado.' }

    $cm = $operacao.CodeModule
    $texto = $cm.Lines(1, $cm.CountOfLines)
    if ($texto -notmatch '(?m)^Public Sub AtualizarOperacao\(\)') {
        throw 'AtualizarOperacao nao encontrado em mOperacao.'
    }
    if ($texto -notmatch '\bAtualizarEngEspec\b') {
        $inicio = $cm.ProcBodyLine('AtualizarOperacao', 0)
        $cm.InsertLines($inicio + 1, "    AtualizarEngEspec        ' especificacao formal + fallback legado")
        'AtualizarOperacao passa a atualizar Eng_Especificacoes'
    }
    else {
        'AtualizarOperacao ja atualizava Eng_Especificacoes'
    }

    # Guarda de nao regressao: este passo nao pode antecipar o corte da interface.
    $referencias = 0
    foreach ($ws in @($wb.Worksheets)) {
        $cel = $null
        try { $cel = $ws.Cells.Find('engETp', [System.Reflection.Missing]::Value, -4123, 2) } catch { }
        if ($null -ne $cel) { $referencias++ }
    }
    if ($referencias -gt 0) {
        throw "Encontradas formulas consumidoras de engETp ($referencias abas). O fallback Hemato exige manter Analitos!Q:R nesta entrega."
    }

    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
