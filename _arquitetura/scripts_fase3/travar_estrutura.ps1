# travar_estrutura.ps1 - trava a criacao de abas pela interface
#
# DECISAO DO GESTOR: usuario comum e ADM NAO criam aba (nem na tela de login,
# nem depois de logar). So o Modo Desenvolvedor (QCDEV@2026) pode, para
# manutencao.
#
# COMO. Patcha mSeguranca (LockApp, UnlockApp, ModoDesenvolvedor) para chamar
# ProtegerEstrutura/DesprotegerEstrutura nos pontos certos, e deixa o ARQUIVO
# SALVO ja com ProtectStructure=True -- assim a trava vale mesmo com macros
# desabilitadas, nao so em tempo de execucao.
#
# Substituicao por NOME, nao por linha: mSeguranca e a mesma nos tres produtos,
# mas amarrar em numero de linha quebraria a cada mudanca a montante.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\travar_estrutura.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$patchFile = Join-Path $arq 'src_hardening1\mSeguranca_ESTRUTURA.txt'
$SENHA = 'qcini2025'

# --- le o patch: blocos '@@FUNC nome e um bloco '@@APPEND ---
$linhas = [System.IO.File]::ReadAllLines($patchFile, (New-Object System.Text.UTF8Encoding($false)))
$blocos = @{}
$append = New-Object System.Collections.ArrayList
$nome = $null
$buf = New-Object System.Collections.ArrayList
foreach ($l in $linhas) {
    if ($l -match "^'@@FUNC\s+(\w+)\s*$") {
        if ($nome -eq '@APPEND') { foreach ($x in $buf) { [void]$append.Add($x) } }
        elseif ($nome) { $blocos[$nome] = $buf.ToArray() }
        $nome = $Matches[1]; $buf = New-Object System.Collections.ArrayList
    }
    elseif ($l -match "^'@@APPEND\s*$") {
        if ($nome -and $nome -ne '@APPEND') { $blocos[$nome] = $buf.ToArray() }
        $nome = '@APPEND'; $buf = New-Object System.Collections.ArrayList
    }
    elseif ($nome) { [void]$buf.Add($l) }
}
if ($nome -eq '@APPEND') { foreach ($x in $buf) { [void]$append.Add($x) } }
elseif ($nome) { $blocos[$nome] = $buf.ToArray() }

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $u = $_; if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }; Start-Sleep -Seconds ($t * 2) }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

try {
    $comp = $null
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mSeguranca') { $comp = $c; break } }
    if ($comp -eq $null) { throw "mSeguranca nao encontrado no projeto" }
    $cm = $comp.CodeModule

    # substitui cada rotina pelo nome
    foreach ($n in $blocos.Keys) {
        # localiza a linha da declaracao
        $ini = 0
        for ($i = 1; $i -le $cm.CountOfLines; $i++) {
            if ($cm.Lines($i, 1) -match "^\s*(Public|Private|Friend)?\s*Sub\s+$n\s*(\(|$)") { $ini = $i; break }
        }
        if ($ini -eq 0) { throw "rotina $n nao encontrada em mSeguranca" }
        # inclui comentarios imediatamente acima
        while ($ini -gt 1 -and $cm.Lines($ini - 1, 1) -match "^\s*'") { $ini-- }
        # fim = primeiro 'End Sub' a partir da declaracao
        $fim = 0
        for ($i = $ini; $i -le $cm.CountOfLines; $i++) {
            if ($cm.Lines($i, 1).Trim() -eq 'End Sub') { $fim = $i; break }
        }
        if ($fim -eq 0) { throw "fim de $n nao encontrado" }
        $cm.DeleteLines($ini, $fim - $ini + 1)
        $cm.InsertLines($ini, ($blocos[$n] -join "`r`n"))
        "  rotina substituida: $n"
    }

    # anexa os helpers, se ainda nao existirem
    $txt = $cm.Lines(1, $cm.CountOfLines)
    if ($txt -notmatch '\bSub\s+ProtegerEstrutura\b') {
        $cm.AddFromString(($append -join "`r`n"))
        "  helpers acrescentados: ProtegerEstrutura, DesprotegerEstrutura"
    }

    # ESTADO EM REPOUSO: o arquivo sai com a estrutura protegida, entao a trava
    # vale mesmo antes de qualquer macro rodar.
    if (-not $wb.ProtectStructure) {
        $wb.Protect($SENHA, $true, $false)     # Structure=$true, Windows=$false
    }
    "  ProtectStructure no arquivo salvo: $($wb.ProtectStructure)"

    $wb.Save()
    "estrutura travada; abas so pelo Modo Desenvolvedor"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
