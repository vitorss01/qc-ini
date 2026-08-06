# instalar_estat_periodo.ps1 - poe mEstatPeriodo na producao e CONFERE
#
# CONFERE, NAO CONFIA: o modulo so e dado por instalado depois de RESPONDER com
# o mesmo numero que as formulas antigas ja produziam. Media, DP e CV do periodo
# anual de 2026 tem de bater com a linha de base capturada antes da mudanca --
# se o motor novo divergir do caminho antigo, e ele que esta errado, e a
# instalacao e recusada.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\instalar_estat_periodo.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Fonte = ''
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Fonte -eq '') {
    $Fonte = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\mEstatPeriodo.bas'
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
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $proj = $wb.VBProject
    $alvo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $alvo = $c; break } }
    if ($alvo -ne $null) { $proj.VBComponents.Remove($alvo); "removido: mEstatPeriodo (versao anterior)" }

    $proj.VBComponents.Import($Fonte) | Out-Null
    $novo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $novo = $c; break } }
    if ($novo -eq $null) { throw 'Import nao criou mEstatPeriodo' }
    "importado: mEstatPeriodo ($($novo.CodeModule.CountOfLines) linhas)"

    # ---------------- conferencia contra o caminho antigo ----------------
    $est = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $est = $w; break } }
    if ($est -eq $null) { throw 'aba Estatistica ausente' }

    $ini = [double](Get-Date -Year 2026 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).ToOADate()
    $fim = [double](Get-Date -Year 2026 -Month 12 -Day 31 -Hour 0 -Minute 0 -Second 0).ToOADate()

    $erros = @()
    $conferidas = 0
    for ($r = 7; $r -le 86; $r++) {
        $a = "$($est.Cells($r,1).Value2)".Trim()
        if ($a -eq '') { continue }
        $niv = "$($est.Cells($r,2).Value2)".Trim()
        $nAnt = $est.Cells($r, 3).Value2
        if ("$nAnt" -eq '' -or [double]$nAnt -lt 2) { continue }

        $nNovo = $xl.Run('EstatPeriodo', $a, $niv, 'N', $ini, $fim, '', '', '')
        $mNovo = $xl.Run('EstatPeriodo', $a, $niv, 'MEDIA', $ini, $fim, '', '', '')
        $dNovo = $xl.Run('EstatPeriodo', $a, $niv, 'DP', $ini, $fim, '', '', '')
        $cNovo = $xl.Run('EstatPeriodo', $a, $niv, 'CV', $ini, $fim, '', '', '')

        $mAnt = $est.Cells($r, 4).Value2
        $dAnt = $est.Cells($r, 5).Value2
        $cAnt = $est.Cells($r, 6).Value2

        if ([double]$nNovo -ne [double]$nAnt) { $erros += "$a N$niv : n antigo $nAnt, novo $nNovo" }
        if ([Math]::Abs([double]$mNovo - [double]$mAnt) -gt 0.0001) { $erros += "$a N$niv : media antiga $mAnt, nova $mNovo" }
        if ([Math]::Abs([double]$dNovo - [double]$dAnt) -gt 0.0001) { $erros += "$a N$niv : DP antigo $dAnt, novo $dNovo" }
        if ([Math]::Abs([double]$cNovo - [double]$cAnt) -gt 0.0001) { $erros += "$a N$niv : CV antigo $cAnt, novo $cNovo" }
        $conferidas++
    }

    # janela de exclusao tem de REDUZIR o n
    $exI = [double](Get-Date -Year 2026 -Month 1 -Day 1).ToOADate()
    $exF = [double](Get-Date -Year 2026 -Month 1 -Day 9).ToOADate()
    $nCheio = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, '', '', '')
    $nCorte = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, $exI, $exF, '')
    if ([double]$nCorte -ge [double]$nCheio) { $erros += "exclusao nao reduziu o n: cheio $nCheio, cortado $nCorte" }

    # exclusao cobrindo tudo tem de zerar
    $nZero = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, $ini, $fim, '')
    if ([double]$nZero -ne 0) { $erros += "exclusao total deveria zerar o n, deu $nZero" }

    # status: ausencia de limite NAO pode virar aprovacao
    $sSemLim = [string]$xl.Run('StatusCV', 2.5, '', '', '')
    $sSoCLIA = [string]$xl.Run('StatusCV', 2.5, 5.67, '', '')
    $sFora = [string]$xl.Run('StatusCV', 9.9, 5.67, '', '')
    $eSemLim = [string]$xl.Run('StatusETP', 10, '', '')
    $eAmbos = [string]$xl.Run('StatusETP', 30, 17, 12)
    if ($sSemLim -ne 'SEM LIMITE') { $erros += "StatusCV sem limite devolveu '$sSemLim'" }
    if ($sSoCLIA -ne 'OK (CLIA)') { $erros += "StatusCV so CLIA devolveu '$sSoCLIA'" }
    if ($sFora -notlike 'FORA*') { $erros += "StatusCV fora devolveu '$sFora'" }
    if ($eSemLim -ne 'SEM LIMITE') { $erros += "StatusETP sem limite devolveu '$eSemLim'" }
    if ($eAmbos -ne 'CRITICO: FORA AMBOS') { $erros += "StatusETP fora ambos devolveu '$eAmbos'" }

    # LimEspec: especificacao ausente tem de virar "", NUNCA 0.
    # Com 0 o StatusCV leria "limite zero" e reprovaria os 31 analitos em VB.
    #
    # O NOME DO ANALITO E MONTADO POR CODIGO DE CARACTERE, e nao escrito como
    # literal. Windows PowerShell 5.1 le o .ps1 como ANSI: um "Acido urico"
    # acentuado digitado aqui chega corrompido ao Excel, nao casa com nenhuma
    # linha do banco e LimEspec devolve vazio -- fazendo a conferencia reprovar
    # um modulo correto. Ja aconteceu nesta sessao.
    $A_ = [string][char]0x00C1      # A agudo maiusculo
    $u_ = [string][char]0x00FA      # u agudo
    $anTeste = $A_ + 'cido ' + $u_ + 'rico'
    $limCLIA = $xl.Run('LimEspec', $anTeste, 2026, 'CLIA', 'ETP')
    $limVB = $xl.Run('LimEspec', $anTeste, 2026, 'Variacao Biologica', 'CV')
    $limNada = $xl.Run('LimEspec', 'ZZ_NAO_EXISTE', 2026, 'CLIA', 'ETP')
    if ("$limCLIA" -eq '' -or [double]$limCLIA -le 0) { $erros += "LimEspec CLIA/ETp de Acido urico devolveu '$limCLIA', esperado 17" }
    if ("$limVB" -ne '') { $erros += "LimEspec VB (nao cadastrada) devolveu '$limVB', esperado vazio" }
    if ("$limNada" -ne '') { $erros += "LimEspec de analito inexistente devolveu '$limNada', esperado vazio" }
    # e o encadeamento que importa: limite ausente nao pode reprovar
    $sVazio = [string]$xl.Run('StatusCV', 2.39, $limCLIA, $limVB, '')
    if ($sVazio -ne 'OK (CLIA)') { $erros += "StatusCV com VB ausente devolveu '$sVazio', esperado 'OK (CLIA)'" }

    # janela derivada
    $jIni = $xl.Run('JanelaInicio', 'TRIMESTRAL', 2026, 1, '', '')
    $jFim = $xl.Run('JanelaFim', 'TRIMESTRAL', 2026, 1, '', '')
    $espIni = (Get-Date -Year 2026 -Month 1 -Day 1).Date
    $espFim = (Get-Date -Year 2026 -Month 3 -Day 31).Date
    if ([datetime]$jIni -ne $espIni) { $erros += "JanelaInicio TRIMESTRAL/1 deu $jIni" }
    if ([datetime]$jFim -ne $espFim) { $erros += "JanelaFim TRIMESTRAL/1 deu $jFim" }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 12 | ForEach-Object { "  FALHA: $_" }
        throw "Instalacao recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }

    "conferencia: $conferidas linha(s) batem com o caminho antigo (n, media, DP, CV)"
    "conferencia: exclusao 01/01-09/01 levou n de $nCheio para $nCorte; exclusao total zerou"
    "conferencia: StatusCV/StatusETP nao aprovam sem limite; TRIMESTRAL/1 = 01/01 a 31/03"

    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
