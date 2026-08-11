# corrigir_painel_nivel_e_rotulos.ps1 - tres defeitos silenciosos da Estatistica v2
#
# 1. O PAINEL MOSTRAVA O BIAS DO NIVEL 1 EM TODOS OS NIVEIS
#
#    Painel!G7 e G8 eram a MESMA formula:
#      INDEX(Estatistica!$K$14:$K$93, MATCH(selAnalito, Estatistica!$A$14:$A$137, 0))
#
#    MATCH sem criterio de nivel devolve a PRIMEIRA ocorrencia do analito, que e
#    sempre o Nivel 1. Medido: Lactato N1 = -3,31 e N2 = -2,68; o Painel exibia
#    -3,31 nas duas linhas. E o erro nao para no bias -- H8 calcula
#    ET% = CV*1,65 + bias, entao o ET do Nivel 2 saia com o bias do Nivel 1:
#    6,96% no lugar de 7,58%.
#
#    Havia ainda faixas incompativeis: INDEX sobre 80 linhas (K14:K93) e MATCH
#    sobre 124 (A14:A137). Uma posicao acima de 80 devolveria erro ou valor de
#    outra linha.
#
#    A correcao casa ANALITO E NIVEL, e le o nivel do proprio rotulo da linha
#    ("N1", "N2", "N3"), para continuar certa se um terceiro nivel aparecer.
#    COUNTIFS antes do SUMIFS porque SUMIFS sem correspondencia devolve ZERO --
#    e zero num campo de bias e um numero plausivel e errado, nao um vazio.
#
# 2. A DATA VOLTOU A EXIBIR O ANO LITERAL
#
#    H5 mostrava "01/01/yyyy". O codigo do ano depende do IDIOMA DE FORMATOS da
#    instalacao: nesta maquina (1046, pt-BR) e "aaaa", e "yyyy" vira TEXTO.
#    Mesmo defeito ja corrigido no DB_Resultados, reintroduzido nas celulas novas
#    do painel. O script TENTA candidatos e so aceita o que RENDERIZA o ano.
#
# 3. CABECALHOS COM A MAIUSCULA ACENTUADA NO LUGAR DA MINUSCULA
#
#    "NIvel" e "MEdia" com I e E maiusculos acentuados (codigos 205 e 201) onde
#    cabem i e e minusculos (237 e 233). Erro de codigo de caractere na montagem.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\corrigir_painel_nivel_e_rotulos.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

$i_ = [string][char]0x00ED   # i minusculo agudo
$e_ = [string][char]0x00E9   # e minusculo agudo

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
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    $est = $null; $pa = $null
    foreach ($w in $wb.Worksheets) {
        if ($w.Name -like 'Estat*') { $est = $w }
        if ($w.Name -eq 'Painel') { $pa = $w }
    }
    if ($est -eq $null -or $pa -eq $null) { throw 'aba Estatistica ou Painel ausente' }
    if ($est.ProtectContents) { try { $est.Unprotect($SENHA) } catch { } }
    if ($pa.ProtectContents) { try { $pa.Unprotect($SENHA) } catch { } }
    $nomeEst = $est.Name

    # ---------- localiza a tabela PELO ROTULO, nao por linha fixa ----------
    $hdr = 0
    for ($r = 1; $r -le 40; $r++) { if ("$($est.Cells($r,1).Value2)".Trim() -eq 'Analito') { $hdr = $r; break } }
    if ($hdr -eq 0) { throw 'cabecalho da Estatistica nao encontrado' }
    $r0 = $hdr + 1
    $rn = $r0
    for ($r = $r0; $r -le $r0 + 200; $r++) { if ("$($est.Cells($r,1).Value2)".Trim() -ne '') { $rn = $r } }

    # coluna do Bias pelo cabecalho
    $colBias = 0
    for ($c = 1; $c -le 34; $c++) {
        if ("$($est.Cells($hdr,$c).Value2)" -like 'Bias*') { $colBias = $c; break }
    }
    if ($colBias -eq 0) { throw 'coluna de Bias nao encontrada pelo rotulo' }
    $letraBias = ("$($est.Cells($hdr,$colBias).Address(0,0))" -replace '\d+$', '')
    "tabela: cabecalho linha $hdr, dados $r0..$rn, Bias na coluna $letraBias"

    # ---------- 1. Painel: bias por NIVEL ----------
    $linhasPainel = @()
    for ($r = 1; $r -le 40; $r++) {
        $a = "$($pa.Cells($r,1).Value2)".Trim()
        if ($a -match '^N(\d+)$') { $linhasPainel += @{ Linha = $r; Nivel = [int]$Matches[1] } }
    }
    if ($linhasPainel.Count -eq 0) { throw 'nenhuma linha N1/N2/N3 encontrada no Painel' }

    foreach ($lp in $linhasPainel) {
        $r = $lp.Linha
        # O NIVEL VEM DO ROTULO DA PROPRIA LINHA ("N2" -> 2), para a formula
        # continuar certa se um terceiro nivel aparecer. A coluna B da
        # Estatistica guarda o NUMERO, entao o "N" sai e o "--" converte.
        #
        # O criterio e montado UMA vez e usado nos dois lugares. Antes eu o
        # inseria com um Replace sobre a formula pronta, procurando "$A8," -- e
        # a ocorrencia dentro do SUMIFS termina em "$A8))", sem virgula. So o
        # COUNTIFS recebia a conversao: o SUMIFS seguia comparando a coluna B
        # (numeros) com o texto "N2", nao casava nada e devolvia ZERO. Zero num
        # campo de bias e exatamente o numero plausivel e errado que este script
        # existe para eliminar.
        $crit = "--SUBSTITUTE(`$A$r,`"N`",`"`")"
        $rgA = "'$nomeEst'!`$A`$$r0`:`$A`$$rn"
        $rgB = "'$nomeEst'!`$B`$$r0`:`$B`$$rn"
        $rgK = "'$nomeEst'!`$$letraBias`$$r0`:`$$letraBias`$$rn"
        $f = "=IF(COUNTIFS($rgA,selAnalito,$rgB,$crit,$rgK,`"<>`")=0,`"`",SUMIFS($rgK,$rgA,selAnalito,$rgB,$crit))"
        $pa.Cells($r, 7).Formula = $f
    }
    "Painel: bias por nivel escrito em $($linhasPainel.Count) linha(s)"

    # ---------- 2. formato de data no painel da Estatistica ----------
    $cand = @(
        @{ Prop = 'NumberFormat'; Valor = 'dd/mm/yyyy' },
        @{ Prop = 'NumberFormatLocal'; Valor = 'dd/mm/aaaa' },
        @{ Prop = 'NumberFormatLocal'; Valor = 'dd/mm/yyyy' }
    )
    $celData = @()
    for ($r = 1; $r -lt $hdr; $r++) {
        for ($c = 1; $c -le 24; $c++) {
            $cel = $est.Cells($r, $c)
            $t = "$($cel.Text)"
            if ($t -match 'yyyy|aaaa' -or ("$($cel.NumberFormatLocal)" -match 'yyyy|aaaa')) {
                if ("$($cel.Value2)" -ne '' -or "$($cel.Formula)" -like '=*') { $celData += $cel }
            }
        }
    }
    $venc = $null
    foreach ($cel in $celData) {
        if ("$($cel.Value2)" -eq '' -or -not ($cel.Value2 -is [double])) { continue }
        $ano = [DateTime]::FromOADate([double]$cel.Value2).Year.ToString()
        foreach ($cd in $cand) {
            try {
                if ($cd.Prop -eq 'NumberFormat') { $cel.NumberFormat = $cd.Valor } else { $cel.NumberFormatLocal = $cd.Valor }
                if ("$($cel.Text)" -like "*$ano*") { $venc = $cd; break }
            }
            catch { }
        }
        if ($venc -ne $null) { break }
    }
    if ($venc -eq $null) { throw 'nenhum formato de data renderizou o ano no painel da Estatistica' }
    foreach ($cel in $celData) {
        if ($venc.Prop -eq 'NumberFormat') { $cel.NumberFormat = $venc.Valor } else { $cel.NumberFormatLocal = $venc.Valor }
    }
    "formato de data: $($venc.Prop)='$($venc.Valor)' em $($celData.Count) celula(s)"

    # colunas largas o bastante para a data caber (##### e uma data invisivel)
    for ($c = 1; $c -le 12; $c++) { if ($est.Columns($c).ColumnWidth -lt 12) { $est.Columns($c).ColumnWidth = 12 } }

    # ---------- 3. cabecalhos com a caixa certa ----------
    $corr = @{ 'Nivel' = 'N' + $i_ + 'vel'; 'Media' = 'M' + $e_ + 'dia' }
    $nCab = 0
    for ($c = 1; $c -le 34; $c++) {
        $h = "$($est.Cells($hdr,$c).Value2)"
        $semAcento = ($h -replace [string][char]0x00CD, 'i') -replace [string][char]0x00C9, 'e'
        $semAcento = ($semAcento -replace [string][char]0x00ED, 'i') -replace [string][char]0x00E9, 'e'
        if ($corr.ContainsKey($semAcento)) {
            # -cne, NAO -ne.
            #
            # Em PowerShell, -ne compara string SEM diferenciar maiuscula de
            # minuscula: 'NIvel' com I maiusculo acentuado e 'Nivel' com i
            # minusculo acentuado sao considerados IGUAIS, e a correcao nunca
            # era escrita. O defeito e justamente a CAIXA do caractere, entao a
            # comparacao precisa enxerga-la.
            if ($h -cne $corr[$semAcento]) {
                $est.Cells($hdr, $c).Value2 = [string]$corr[$semAcento]
                $nCab++
            }
        }
    }
    "cabecalhos corrigidos: $nCab"

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # ---------- conferencia ----------
    $erros = @()

    # o bias de cada nivel do Painel tem de bater com a linha correspondente da Estatistica
    $sel = "$($wb.Names.Item('selAnalito').RefersToRange.Value2)".Trim()
    foreach ($lp in $linhasPainel) {
        $esperado = ''
        for ($r = $r0; $r -le $rn; $r++) {
            if ("$($est.Cells($r,1).Value2)".Trim() -eq $sel -and [string]$est.Cells($r, 2).Value2 -eq [string]$lp.Nivel) {
                $esperado = "$($est.Cells($r,$colBias).Value2)"
                break
            }
        }
        $obtido = "$($pa.Cells($lp.Linha,7).Value2)"
        if ($esperado -eq '' -and $obtido -ne '') { $erros += "Painel N$($lp.Nivel): esperado vazio, veio '$obtido'" }
        elseif ($esperado -ne '' -and [Math]::Abs([double]$obtido - [double]$esperado) -gt 0.0001) {
            $erros += "Painel N$($lp.Nivel): esperado $esperado, veio $obtido"
        }
    }
    # os niveis nao podem ser todos iguais quando a Estatistica os tem diferentes
    if ($linhasPainel.Count -ge 2) {
        $v = @()
        foreach ($lp in $linhasPainel) { $v += "$($pa.Cells($lp.Linha,7).Value2)" }
        $distintosEst = @()
        foreach ($lp in $linhasPainel) {
            for ($r = $r0; $r -le $rn; $r++) {
                if ("$($est.Cells($r,1).Value2)".Trim() -eq $sel -and [string]$est.Cells($r, 2).Value2 -eq [string]$lp.Nivel) {
                    $distintosEst += "$($est.Cells($r,$colBias).Value2)"; break
                }
            }
        }
        if (($distintosEst | Select-Object -Unique).Count -gt 1 -and ($v | Select-Object -Unique).Count -eq 1) {
            $erros += "Painel repete o mesmo bias ($($v[0])) em todos os niveis, mas a Estatistica tem valores distintos"
        }
    }

    # data tem de mostrar o ano
    for ($r = 1; $r -lt $hdr; $r++) {
        for ($c = 1; $c -le 24; $c++) {
            $t = "$($est.Cells($r,$c).Text)"
            if ($t -match 'yyyy|aaaa') { $erros += "$($est.Cells($r,$c).Address(0,0)) ainda mostra o codigo do formato: '$t'" }
            if ($t -match '^#+$') { $erros += "$($est.Cells($r,$c).Address(0,0)) esta estreita demais: '$t'" }
        }
    }

    # cabecalhos
    # -cmatch, NAO -match: -match tambem ignora a caixa, entao o 'i' minusculo
    # acentuado casaria com o padrao 'I' maiusculo e a prova acusaria erro num
    # cabecalho ja corrigido. Mesma armadilha do -ne, do outro lado.
    foreach ($c in 1..34) {
        $h = "$($est.Cells($hdr,$c).Value2)"
        if ($h -cmatch ([string][char]0x00CD) -or $h -cmatch ([string][char]0x00C9)) {
            $erros += "cabecalho col $c ainda com maiuscula acentuada: '$h'"
        }
    }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 10 | ForEach-Object { "  FALHA: $_" }
        throw "Correcao recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }

    $res = @()
    foreach ($lp in $linhasPainel) { $res += "N$($lp.Nivel)=$($pa.Cells($lp.Linha,7).Text)" }
    "conferencia: bias por nivel no Painel -> $($res -join '  ')  (analito '$sel')"

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
