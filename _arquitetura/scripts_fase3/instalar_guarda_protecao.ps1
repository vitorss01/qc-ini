# instalar_guarda_protecao.ps1 - ADR-046: o par LiberarEscrita/RestaurarProtecao
#
# POR QUE UM SCRIPT PROPRIO
#
# mSeguranca e fonte de PRODUCAO: o build nao a regera, ela vem do workbook e
# so se altera por script versionado (ADR-021). travar_estrutura.ps1 ja faz
# isso, mas roda tarde no build -- depois de rodar_motor.ps1, que executa
# AtualizarEstatistica. Como as rotinas do motor passaram a chamar o par, o
# artefato precisa delas ANTES da primeira execucao, ou o projeto nao compila
# e o sintoma e "Sub ou Function nao definida".
#
# Idempotente: se as rotinas ja existem no modulo, sai sem tocar em nada.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso: .\instalar_guarda_protecao.ps1 -Workbook <arquivo.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$SENHA = 'qcini2025'
$arq = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$fonte = Join-Path $PSScriptRoot '..\src_producao\mSeguranca_GUARDA.txt'
$fonte = [System.IO.Path]::GetFullPath($fonte)
if (-not (Test-Path $fonte)) { throw "fonte nao encontrada: $fonte" }

# O .txt e UTF-8 (editor moderno); o modulo VBA e cp1252. Ler com a
# codificacao errada transformaria acento em lixo dentro do codigo.
$encFonte = New-Object System.Text.UTF8Encoding($false)
$bloco = [System.IO.File]::ReadAllText($fonte, $encFonte)

function Novo-Excel {
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { Start-Sleep -Seconds (1 + $t) }
    }
    throw 'nao consegui criar a instancia do Excel'
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvou = $false
try {
    $comp = $null
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Name -eq 'mSeguranca') { $comp = $c; break }
    }
    if ($comp -eq $null) { throw 'mSeguranca nao encontrado no projeto' }
    $cm = $comp.CodeModule
    $txt = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }

    # A constante da senha vem antes das rotinas que a usam. Declaracao de
    # modulo TEM de preceder a primeira procedure -- regra que ja custou tres
    # quebras de compilacao neste projeto.
    if ($txt -notmatch '\bSENHA_PROT\b') {
        $decl = @(
            "",
            "' Senha das abas tecnicas, uma vez so neste modulo. As rotinas que",
            "' escrevem em aba protegida usam LiberarEscrita/RestaurarProtecao e",
            "' nao repetem o literal.",
            "Private Const SENHA_PROT As String = ""$SENHA"""
        ) -join "`r`n"
        # depois do Option Explicit, ainda na area de declaracoes
        $linhaDecl = 1
        for ($i = 1; $i -le [Math]::Min(20, $cm.CountOfLines); $i++) {
            if ($cm.Lines($i, 1) -match '^\s*Option\s+Explicit') { $linhaDecl = $i; break }
        }
        $cm.InsertLines($linhaDecl + 1, $decl)
        "  constante SENHA_PROT declarada em mSeguranca"
    }
    else {
        "  constante SENHA_PROT ja existia"
    }

    $txt = $cm.Lines(1, $cm.CountOfLines)
    if ($txt -match '\bFunction\s+LiberarEscrita\b') {
        "  guarda ja instalada, nada a fazer"
    }
    else {
        $cm.AddFromString($bloco)
        "  LiberarEscrita e RestaurarProtecao acrescentadas a mSeguranca"
    }

    # Conferencia: as duas tem de estar visiveis como Public.
    $txt = $cm.Lines(1, $cm.CountOfLines)
    foreach ($n in @('LiberarEscrita', 'RestaurarProtecao')) {
        if ($txt -notmatch "Public\s+(Function|Sub)\s+$n\b") {
            throw "$n nao ficou publica em mSeguranca"
        }
    }

    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { $wb.Close($salvou) } catch { }
    try { $xl.Quit() } catch { }
}
