# normalizar_protecao.ps1 - deixa o artefato de build com as abas destravadas
#
# POR QUE ISTO EXISTE
#
# O arquivo de producao e o arquivo em que o usuario TRABALHA. Ao logar, o
# LockApp protege todas as abas; se o usuario salvar depois de logar, o arquivo
# fica gravado nesse estado. O estado de protecao do .xlsm, portanto, depende de
# como a ultima sessao terminou -- e nao de nenhuma decisao de projeto.
#
# O build assumia abas destravadas porque era assim que o arquivo estava no Git.
# Em 05/08/2026 o arquivo de producao chegou ao build com as 16 abas protegidas
# e redirecionar_calc.ps1 morreu com "A celula ou grafico ... esta protegida".
# A causa nao era o script de redirecionamento: era o build depender de um
# estado incidental.
#
# Este passo remove essa dependencia. Roda logo apos a copia da producao e
# devolve o artefato num estado CONHECIDO. A protecao final de distribuicao e
# reaplicada no fim do build (travar_estrutura.ps1) e, em uso, pelo LockApp.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\normalizar_protecao.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $u = $_; if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }; Start-Sleep -Seconds ($t * 2) }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$salvou = $false
$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

try {
    if ($wb.ProtectStructure) {
        $wb.Unprotect($SENHA)
        "estrutura: destravada"
    }
    else { "estrutura: ja estava destravada" }

    $destravadas = @()
    $resistentes = @()
    foreach ($ws in @($wb.Worksheets)) {
        if (-not $ws.ProtectContents) { continue }
        # Duas tentativas: com a senha do projeto e sem senha. Uma aba pode ter
        # sido protegida pela interface do Excel, sem senha nenhuma.
        $ok = $false
        try { $ws.Unprotect($SENHA); $ok = -not $ws.ProtectContents } catch { }
        if (-not $ok) { try { $ws.Unprotect(); $ok = -not $ws.ProtectContents } catch { } }
        if ($ok) { $destravadas += $ws.Name } else { $resistentes += $ws.Name }
    }

    if ($destravadas.Count -gt 0) { "abas destravadas ($($destravadas.Count)): $($destravadas -join ', ')" }
    else { "abas destravadas (0): nenhuma estava protegida" }

    # Aba que nao abre com a senha conhecida e um bloqueio DESCONHECIDO: seguir
    # em frente escreveria um artefato pela metade, exatamente o modo de falha
    # que este passo existe para impedir.
    if ($resistentes.Count -gt 0) {
        throw "abas que nao abriram com a senha do projeto: $($resistentes -join ', ')"
    }

    $wb.Save()
    $salvou = $true
    "protecao normalizada"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
