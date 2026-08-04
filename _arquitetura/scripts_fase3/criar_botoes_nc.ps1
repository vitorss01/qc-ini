# criar_botoes_nc.ps1 - botoes da Sprint NC nas abas Resultados e Registros
#
# O botao chama um PONTO DE ENTRADA em mRegistros, nunca o formulario direto:
# o pipeline reconstroi os formularios a cada build, e um botao amarrado ao
# nome do formulario quebraria silenciosamente se o nome mudasse.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Novo-Excel {
    $ultimo = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

try {
    $botoes = @(
        @{ Aba = 'Resultados'; Nome = 'btnNaoConforme'; Texto = 'Registrar Resultado Nao Conforme'; Macro = 'AbrirFormNaoConforme'; L = 470; T = 8; W = 210; H = 26 },
        @{ Aba = 'Registros'; Nome = 'btnExcluirNC'; Texto = 'Excluir Registro'; Macro = 'AbrirFormExcluirRegistroNC'; L = 700; T = 8; W = 140; H = 26 }
    )
    foreach ($b in $botoes) {
        $ws = $wb.Worksheets.Item($b.Aba)
        $vis = $ws.Visible
        $ws.Visible = -1
        try { $ws.Unprotect($SENHA) } catch { }

        foreach ($sh in @($ws.Shapes)) { if ($sh.Name -eq $b.Nome) { $sh.Delete() } }

        $sh = $ws.Shapes.AddShape(5, $b.L, $b.T, $b.W, $b.H)   # msoShapeRoundedRectangle
        $sh.Name = $b.Nome
        $sh.TextFrame2.TextRange.Text = $b.Texto
        $sh.TextFrame2.TextRange.Font.Size = 9
        $sh.TextFrame2.TextRange.Font.Bold = $true
        $sh.Fill.ForeColor.RGB = 12419407
        $sh.Line.Visible = $false
        $sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        $sh.OnAction = $b.Macro

        $ws.Visible = $vis
        "  $($b.Aba): botao $($b.Nome) -> $($b.Macro)"
    }
    $wb.Save()
    "botoes da Sprint NC instalados"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
