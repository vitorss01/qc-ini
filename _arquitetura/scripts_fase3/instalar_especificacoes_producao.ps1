# instalar_especificacoes_producao.ps1 - poe mEspecificacoes na PRODUCAO
#
# POR QUE ISTO PRECISOU EXISTIR
#
# mEspecificacoes sao 575 linhas que decidem CONFORME / NAO CONFORME, e Painel
# e Estatistica passaram a ler delas. Isso e ENGENHARIA, e o ADR-021 e explicito:
# "tudo que for engenharia entra por script versionado e nunca e aplicado a mao
# na producao". O modulo, porem, chegou ao arquivo pela VBE: nenhum script o
# importava. criar_form_especificacoes.ps1 apenas EXIGE que ele ja esteja la
# ("mEspecificacoes nao esta no projeto") -- confere, nao instala.
#
# A consequencia nao e teorica. E o mesmo defeito de julho que o ADR-021 diz
# existir para impedir: a fonte versionada e o que roda dentro do arquivo podem
# divergir sem que nada acuse. Corrigir src_producao\mEspecificacoes.bas nao
# mudava nada em lugar nenhum, porque nada levava o arquivo para dentro do
# .xlsm.
#
# CONFERE, NAO CONFIA: o modulo so e dado por instalado depois de RESPONDER.
# Contar linhas prova que entrou texto; chamar EspecAtiva prova que compila e
# que a versao instalada e a que corrige o acento do "Nao".
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\instalar_especificacoes_producao.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Fonte = ''
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Fonte -eq '') {
    $Fonte = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\mEspecificacoes.bas'
}
$Fonte = (Resolve-Path -LiteralPath $Fonte).ProviderPath

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
# 1 = msoAutomationSecurityLow, e NAO 3 (ForceDisable) como nos scripts que so
# mexem em celula. A conferencia abaixo CHAMA rotina do modulo, e com 3 o Excel
# desabilita as macros: o Run falha com "as macros foram desabilitadas" e a
# instalacao e recusada por um motivo que nao tem nada a ver com o modulo.
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $proj = $wb.VBProject

    $antes = 0
    $alvo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEspecificacoes') { $alvo = $c; break } }
    if ($alvo -ne $null) {
        $antes = $alvo.CodeModule.CountOfLines
        $proj.VBComponents.Remove($alvo)
        "removido: mEspecificacoes ($antes linhas)"
    }
    else { "mEspecificacoes nao existia no projeto" }

    $proj.VBComponents.Import($Fonte) | Out-Null

    $novo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEspecificacoes') { $novo = $c; break } }
    if ($novo -eq $null) { throw "Import nao criou o componente mEspecificacoes" }
    $depois = $novo.CodeModule.CountOfLines
    if ($depois -lt 100) { throw "mEspecificacoes entrou com apenas $depois linhas" }
    "importado: mEspecificacoes ($depois linhas, fonte $(Split-Path -Leaf $Fonte))"

    # ---- CONFERE, NAO CONFIA ----
    # Chamar uma rotina do modulo prova tres coisas de uma vez: o componente
    # existe, o projeto COMPILA (VBA compila sob demanda -- um erro de sintaxe
    # ficaria escondido ate alguem usar) e a versao instalada e a corrigida.
    $erros = @()

    # EspecAtiva com a grafia ACENTUADA. Com a versao antiga do modulo a funcao
    # nem existe, e a chamada falha -- que e o resultado certo para um script
    # que promete ter instalado a correcao.
    $naoAcentuado = [string][char]0x004E + [string][char]0x00E3 + 'o'
    try {
        $rNao = $xl.Run('EspecAtiva', $naoAcentuado)
        $rSim = $xl.Run('EspecAtiva', 'Sim')
        $rVazio = $xl.Run('EspecAtiva', '')
        if ($rNao) { $erros += "EspecAtiva('$naoAcentuado') devolveu True: a correcao do acento NAO esta instalada" }
        if (-not $rSim) { $erros += "EspecAtiva('Sim') devolveu False" }
        if (-not $rVazio) { $erros += "EspecAtiva('') devolveu False: linha em branco perderia a meta em silencio" }
    }
    catch { $erros += "EspecAtiva nao respondeu: $($_.Exception.Message)" }

    # MetasDaLinha: prova que o protocolo continua invariante de localidade.
    try {
        $m = [string]$xl.Run('MetasDaLinha', 'CV_BIAS_DIRETO', '', '', '', '', 2.5, 1.25)
        if ($m -match ',') { $erros += "MetasDaLinha devolveu virgula decimal: '$m'" }
        $p = $m -split '\|'
        if ($p.Count -ne 3) { $erros += "MetasDaLinha devolveu $($p.Count) campos: '$m'" }
        elseif ([Math]::Abs([double]$p[2] - (1.25 + 1.65 * 2.5)) -gt 0.0001) {
            $erros += "MetasDaLinha derivou ETp errado: '$m'"
        }
        else { "conferencia: EspecAtiva e MetasDaLinha responderam ('$m')" }
    }
    catch { $erros += "MetasDaLinha nao respondeu: $($_.Exception.Message)" }

    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Instalacao de mEspecificacoes rejeitada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }

    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
