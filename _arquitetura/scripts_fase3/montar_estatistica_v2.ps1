# montar_estatistica_v2.ps1 - novo layout e painel de filtros da aba Estatistica
#
# ORDEM DAS COLUNAS pedida pelo gestor:
#   A Analito | B Nivel | C n | D Media | E DP | F CV% |
#   G Lim CV CLIA% | H Lim CV VB% | I Lim CV FAB% | J Status CV |
#   K Bias% | L Lim Bias VB% | M Lim Bias FAB% |
#   N ET% | O ETP CLIA% | P ETP VB% | Q Status ETP |
#   R Sigma | S Classificacao
#
# O BLOCO P:U SAI DE CASA. Ele guardava as medias de bias vindas do EQC_Dados e
# agora colide com ETP VB (P) e Status ETP (Q). Vai para AA:AF, preservado: nao
# ha dado de EQC hoje, mas apagar a estrutura fecharia a porta do bias por
# ensaio de proficiencia, que e o metodologicamente correto.
#
# BIAS: a coluna K passa a usar o ALVO DO LOTE (LotesStore), unica origem com
# dado. O cabecalho DIZ isso. Bias por EQC continua sendo o certo para avaliacao
# de metodo e volta a valer quando houver EQC cadastrado.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI):
# todo texto acentuado e montado por [char]. [char]+[char] e SOMA NUMERICA, por
# isso cada pedaco leva [string].
#
# Uso:
#   .\montar_estatistica_v2.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

# --- acentos por codigo ---
$a_ = [string][char]0x00E1   # a agudo
$e_ = [string][char]0x00E9   # e agudo
$i_ = [string][char]0x00ED   # i agudo
$o_ = [string][char]0x00F3   # o agudo
$c_ = [string][char]0x00E7   # c cedilha
$a2 = [string][char]0x00E3   # a til
$I_ = [string][char]0x00CD   # I agudo
$E_ = [string][char]0x00C9   # E agudo
$ni = 'N' + $i_ + 'vel'
$med = 'M' + $e_ + 'dia'
$per = 'Per' + $i_ + 'odo'
$vis = 'Vis' + $a2 + 'o'
$cla = 'Classifica' + $c_ + $a2 + 'o'
$exc = 'EXCLUS' + [string][char]0x00C3 + 'O'
$per_ef = $per + ' efetivo'
$anual = 'ANUAL'

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

    $est = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $est = $w; break } }
    if ($est -eq $null) { throw 'aba Estatistica ausente' }
    if ($est.ProtectContents) { try { $est.Unprotect($SENHA) } catch { } }
    $nomeAba = $est.Name

    $R0 = 7
    $RN = 86

    # ---------- 1. realoca o bloco de bias do EQC (P:U -> AA:AF) ----------
    $jaMoveu = ("$($est.Cells(6,27).Value2)".Trim() -ne '')
    if (-not $jaMoveu) {
        $est.Range($est.Cells(6, 16), $est.Cells($RN, 21)).Copy($est.Cells(6, 27)) | Out-Null
        $est.Range($est.Cells(6, 16), $est.Cells($RN, 21)).ClearContents() | Out-Null
        "bloco de bias do EQC realocado: P:U -> AA:AF"
    }
    else { "bloco AA:AF ja existe, preservado" }

    # ---------- 2. painel de filtros (linhas 1..5) ----------
    #
    # DESMESCLAR ANTES DE ESCREVER.
    #
    # O cabecalho antigo tinha celulas mescladas nas linhas 1..5. Atribuir
    # .Formula a um membro nao-ancora de uma mesclagem NAO GRAVA e NAO LEVANTA
    # ERRO pela via do COM: a celula fica vazia e o script segue reportando
    # sucesso. Foi assim que H3/H4 (a janela efetiva) ficaram em branco enquanto
    # F5 -- fora da mesclagem -- calculava normalmente. Com a janela vazia o
    # motor recebe 0 e 0, entende "sem limite" e calcula sobre o historico
    # INTEIRO: numeros plausiveis na tela e filtro de periodo inerte.
    $est.Range('A1:S5').UnMerge() | Out-Null
    $est.Range('A1:S5').ClearContents() | Out-Null

    $est.Range('A1').Value2 = 'ESTAT' + $I_ + 'STICA DE DESEMPENHO ANAL' + $I_ + 'TICO'
    $est.Range('A1').Font.Bold = $true
    $est.Range('A2').Value2 = 'Imprecis' + $a2 + 'o (CV), inexatid' + $a2 + 'o (Bias) e erro total (ET) confrontados com as especifica' + $c_ + $o_ + 'es de qualidade.'

    $est.Range('A3').Value2 = $vis
    $est.Range('C3').Value2 = 'Ano'
    $est.Range('E3').Value2 = 'M' + $e_ + 's/Tri/Sem'
    $est.Range('G3').Value2 = 'Janela de'
    $est.Range('A4').Value2 = 'Lote (vazio = todos)'
    $est.Range('C4').Value2 = $per + ' de'
    $est.Range('E4').Value2 = $a_ + 't' + $e_
    $est.Range('G4').Value2 = $a_ + 't' + $e_
    $est.Range('A5').Value2 = $exc + ' de'
    $est.Range('C5').Value2 = $a_ + 't' + $e_
    $est.Range('E5').Value2 = $per_ef + ':'
    foreach ($c in @('A3', 'C3', 'E3', 'G3', 'A4', 'C4', 'E4', 'G4', 'A5', 'C5', 'E5')) { $est.Range($c).Font.Bold = $true }

    $est.Range('B3').Value2 = $anual
    $est.Range('D3').Value2 = [double]2026
    $est.Range('F3').Value2 = [double]1
    # B4 = lote, preservado se ja tiver valor
    $est.Range('D4').NumberFormatLocal = 'dd/mm/aaaa'
    $est.Range('F4').NumberFormatLocal = 'dd/mm/aaaa'
    $est.Range('B5').NumberFormatLocal = 'dd/mm/aaaa'
    $est.Range('D5').NumberFormatLocal = 'dd/mm/aaaa'
    $est.Range('H3').NumberFormatLocal = 'dd/mm/aaaa'
    $est.Range('H4').NumberFormatLocal = 'dd/mm/aaaa'

    # validacao do seletor de visao
    $v = $est.Range('B3').Validation
    try { $v.Delete() } catch { }
    $v.Add(3, 1, 1, 'MENSAL,TRIMESTRAL,SEMESTRAL,ANUAL,PERSONALIZADO') | Out-Null
    $v.IgnoreBlank = $true
    $v.InCellDropdown = $true

    # janela efetiva: UMA pergunta, UMA resposta -- o seletor decide, e o
    # resultado fica VISIVEL em H3/H4 e escrito por extenso em F5.
    $est.Range('H3').Formula = '=JanelaInicio($B$3,$D$3,$F$3,$D$4,$F$4)'
    $est.Range('H4').Formula = '=JanelaFim($B$3,$D$3,$F$3,$D$4,$F$4)'
    $est.Range('F5').Formula = '=PeriodoEfetivo($H$3,$H$4,$B$5,$D$5)'
    $est.Range('F5').Font.Bold = $true

    # ---------- 3. nomes definidos ----------
    $nomes = @{
        'Estat_Visao'          = 'B3'
        'Estat_Ano'            = 'D3'
        'Estat_Parte'          = 'F3'
        'Estat_Lote'           = 'B4'
        'Data_Inicio_Visao'    = 'D4'
        'Data_Fim_Visao'       = 'F4'
        'Data_Inicio_Exclusao' = 'B5'
        'Data_Fim_Exclusao'    = 'D5'
        'Estat_Ini_Efetiva'    = 'H3'
        'Estat_Fim_Efetiva'    = 'H4'
    }
    foreach ($n in $nomes.Keys) {
        try { $wb.Names.Item($n).Delete() } catch { }
        $wb.Names.Add($n, "='$nomeAba'!`$$($nomes[$n].Substring(0,1))`$$($nomes[$n].Substring(1))") | Out-Null
    }
    "nomes definidos: $($nomes.Count)"

    # ---------- 4. cabecalho ----------
    $cab = @(
        'Analito', $ni, 'n', $med, 'DP', 'CV %',
        'Lim CV CLIA %', 'Lim CV VB %', 'Lim CV FAB %', 'Status CV',
        'Bias % (alvo do lote)', 'Lim Bias VB %', 'Lim Bias FAB %',
        'ET %', 'ETP CLIA %', 'ETP VB %', 'Status ETP',
        'Sigma', $cla
    )
    $est.Range($est.Cells(6, 1), $est.Cells(6, 26)).ClearContents() | Out-Null
    for ($k = 0; $k -lt $cab.Count; $k++) {
        $est.Cells(6, $k + 1).Value2 = [string]$cab[$k]
    }
    $est.Range($est.Cells(6, 1), $est.Cells(6, $cab.Count)).Font.Bold = $true
    $est.Range($est.Cells(6, 1), $est.Cells(6, $cab.Count)).Interior.Color = 14277081

    # ---------- 5. formulas ----------
    # A e B (analito/nivel) NAO sao tocadas: sao a identidade da linha.
    $JAN = 'Estat_Ini_Efetiva,Estat_Fim_Efetiva,Data_Inicio_Exclusao,Data_Fim_Exclusao,$B$4'
    $ANO = 'YEAR(Estat_Fim_Efetiva)'
    $VB = '"Variacao Biologica"'
    $FB = '"Fabricante"'

    $f = @{
        3  = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"N`",$JAN))"
        4  = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"MEDIA`",$JAN))"
        5  = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"DP`",$JAN))"
        6  = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"CV`",$JAN))"
        7  = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,`"CLIA`",`"CV`"))"
        8  = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,$VB,`"CV`"))"
        9  = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,$FB,`"CV`"))"
        10 = "=IF(`$A7=`"`",`"`",StatusCV(`$F7,`$G7,`$H7,`$I7))"
        11 = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"BIAS`",$JAN))"
        12 = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,$VB,`"BIAS`"))"
        13 = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,$FB,`"BIAS`"))"
        14 = "=IF(`$A7=`"`",`"`",EstatPeriodo(`$A7,`$B7,`"ET`",$JAN))"
        15 = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,`"CLIA`",`"ETP`"))"
        16 = "=IF(`$A7=`"`",`"`",LimEspec(`$A7,$ANO,$VB,`"ETP`"))"
        17 = "=IF(`$A7=`"`",`"`",StatusETP(`$N7,`$O7,`$P7))"
        # Sigma = (ETp - |Bias|) / CV. Sem ETp, sem Bias ou com CV zero nao ha
        # sigma -- devolver 0 ou um numero grande seria inventar desempenho.
        18 = "=IF(OR(`$F7=`"`",`$F7=0,`$O7=`"`",`$K7=`"`"),`"`",(`$O7-ABS(`$K7))/`$F7)"
        19 = "=IF(`$R7=`"`",`"`",IF(`$R7>=6,`"Excelente`",IF(`$R7>=4,`"Bom`",IF(`$R7>=3,`"Marginal`",`"Inaceit" + $a_ + "vel`"))))"
    }

    $est.Range($est.Cells($R0, 3), $est.Cells($RN, 19)).ClearContents() | Out-Null
    foreach ($c in ($f.Keys | Sort-Object)) {
        $est.Range($est.Cells($R0, $c), $est.Cells($RN, $c)).Formula = $f[$c]
    }
    "formulas escritas: colunas C..S, linhas $R0..$RN"

    $est.Columns('A:S').AutoFit() | Out-Null
    $est.Columns('A').ColumnWidth = 24
    $est.Columns('K').ColumnWidth = 20

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # ---------- 6. conferencia ----------
    $erros = @()
    for ($k = 0; $k -lt $cab.Count; $k++) {
        $lido = "$($est.Cells(6, $k + 1).Value2)"
        if ($lido -ne [string]$cab[$k]) { $erros += "cabecalho col $($k+1): esperado '$($cab[$k])', lido '$lido'" }
    }
    $comN = 0; $comCV = 0; $comStatus = 0; $forcaFalsa = 0
    for ($r = $R0; $r -le $RN; $r++) {
        if ("$($est.Cells($r,1).Value2)".Trim() -eq '') { continue }
        $n = $est.Cells($r, 3).Value2
        if ("$n" -ne '' -and [double]$n -gt 0) { $comN++ }
        if ("$($est.Cells($r,6).Value2)" -ne '') { $comCV++ }
        $s = "$($est.Cells($r,10).Value2)"
        if ($s -ne '') { $comStatus++ }
        # o defeito que nao pode existir: limite VB ausente virando reprovacao
        if ("$($est.Cells($r,8).Value2)" -eq '' -and $s -like '*VB*') { $forcaFalsa++ }
    }
    if ($comN -lt 60) { $erros += "so $comN linha(s) com n>0 (esperado ~62)" }
    if ($comCV -lt 60) { $erros += "so $comCV linha(s) com CV (esperado ~62)" }
    if ($forcaFalsa -gt 0) { $erros += "$forcaFalsa linha(s) acusam VB sem ter limite de VB cadastrado" }

    # A JANELA TEM DE ESTAR DEFINIDA, E ISSO PRECISA SER CONFERIDO.
    #
    # Com H3/H4 vazias o motor recebe 0 e 0, que significa "sem limite dos dois
    # lados" -- ou seja, calcula sobre o historico INTEIRO. Os numeros aparecem
    # certos e o filtro de periodo nao esta funcionando. Foi o que aconteceu na
    # primeira montagem: 62 linhas com n=18 e a janela sem derivar.
    $hIni = $est.Range('H3').Value2
    $hFim = $est.Range('H4').Value2
    if ("$hIni" -eq '' -or "$hFim" -eq '') { $erros += 'H3/H4 (janela efetiva) nao derivaram data' }
    $txtPer = "$($est.Range('F5').Text)"
    if ($txtPer -like '*invalido*' -or $txtPer -like '*nao definido*') { $erros += "periodo efetivo: '$txtPer'" }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 10 | ForEach-Object { "  FALHA: $_" }
        throw "Montagem recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }
    "conferencia: $comN linha(s) com n, $comCV com CV, $comStatus com Status CV; nenhuma acusa VB sem limite"
    "periodo efetivo: $($est.Range('F5').Text)"

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
