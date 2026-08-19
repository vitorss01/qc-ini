# preparar_contrato_bi.ps1 - cria identidade persistente do contrato Excel -> BI

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$Produto,
    [string]$Versao = '2.0.0'
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

function Novo-Excel {
    $ultimo = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $ultimo = $_; Start-Sleep -Seconds ($t * 2) }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

function Definir-NomeConstante($wb, [string]$nome, [string]$valor) {
    try { $wb.Names.Item($nome).Delete() } catch { }
    $n = $wb.Names.Add($nome, ('="' + $valor.Replace('"', '""') + '"'))
    $n.Visible = $false
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $xl.Workbooks.Open($Workbook)
try {
    if ($wb.ReadOnly) { throw "Somente leitura: $Workbook" }
    $id = [guid]::NewGuid().ToString('D').ToUpperInvariant()
    Definir-NomeConstante $wb 'biProduto' $Produto
    Definir-NomeConstante $wb 'biWorkbookID' $id
    Definir-NomeConstante $wb 'biContratoVersao' $Versao
    $wb.Save()
    "Produto: $Produto"
    "WorkbookID: $id"
    "VersaoContrato: $Versao"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
