# fase0_ajustes.ps1 - correcoes objetivas da Fase 0
#
# O QUE FOI MEDIDO ANTES DE MEXER (auditoria, nao suposicao)
#
# 1. Painel C:J ja estao com largura 12. Nada a fazer -- registrado aqui para
#    que a proxima leitura do script nao refaca a pergunta.
#
# 2. A mescla Painel!O3:X4 e SEGURA e fica. Contem UMA formula real (o aviso
#    "FILTRO ATIVO"), nenhuma outra aba a referencia, nenhum nome definido cai
#    nela e nenhum shape a cobre (btnDev esta em R1, fora). Nao ha razao
#    tecnica para remover, e preferencia contra mescla nao e razao.
#
# 3. SEPARADOR DE VALIDACAO -- e aqui a medicao mudou o diagnostico.
#
#    O COM traduz o separador para o da LOCALIDADE ao ler e ao escrever. Nesta
#    maquina (pt-BR) o separador e ";", e nove listas ja estao corretas assim:
#    'OTI;DES;MIN', 'Ativo;Excluido', 'Sim;Nao', 'ANALISTA;TECNICO;ADM'...
#
#    TRES estao com virgula e por isso o Excel as le como UM item so:
#      DB_Resultados!C4  '1,2'
#      Estatistica!B4    'ANUAL,SEMESTRAL,TRIMESTRAL,MENSAL,PERSONALIZADO'
#      Estatistica!F4    '1,2,3,4,5,6,7,8,9,10,11,12'
#
#    A prova de que o defeito e real, e nao teorico: o VALOR de F4 e o texto
#    inteiro "1,2,3,4,5,6,7,8,9,10,11,12". O usuario abriu a lista, viu um item
#    unico e selecionou-o. JanelaInicio recebe isso como "parte" e Val() para na
#    virgula, devolvendo 1 -- numero plausivel vindo de um campo corrompido.
#
#    Nenhum VBA cria ou altera validacao: todas sao estaticas.
#
# 4. B10 CARREGA RESIDUO DO LAYOUT ANTIGO.
#
#    A linha 10 hoje e a EXCLUSAO 3 (A10='3', B10=inicio, C10=fim). Mas B10
#    ainda tem o valor 'PERSONALIZADO' e um dropdown de Visao herdados de quando
#    o painel morava ali. O espelho V10 (=IF($B$10="","",$B$10)) copia esse
#    texto para a coluna que o motor le como DATA de exclusao.
#
#    Nao corrompe o calculo -- ComoData("PERSONALIZADO") devolve 0 e o par e
#    ignorado --, mas deixa a exclusao 3 inutilizavel e poe um seletor de visao
#    dentro de um campo de data. E uma armadilha esperando o usuario.
#
# 5. SEIS NOMES DEFINIDOS APONTAM PARA O LUGAR ERRADO.
#
#    Estat_Visao=B10, Estat_Ano=D10, Estat_Parte=F10, Estat_Lote=B11,
#    Data_Inicio_Visao=D11, Data_Fim_Visao=F11 -- todos na geometria ANTIGA.
#    As entradas reais estao em B4, D4, F4, H4, B5 e D5.
#
#    Medido: os seis sao usados por ZERO formulas e ZERO linhas de VBA. Nao ha
#    impacto hoje. Corrigi-los mesmo assim porque a Fase 1 vai escrever nessas
#    celulas a partir de um UserForm, e um nome que aponta para o campo errado
#    e exatamente como se grava dado no lugar errado sem ninguem notar.
#
#    Os que o motor USA (Estat_Ini_Efetiva, Estat_Fim_Efetiva, Estat_Exclusoes)
#    estao certos e NAO sao tocados.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\fase0_ajustes.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

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

function Aba { param($W, $N) foreach ($x in $W.Worksheets) { if ($x.Name -like $N) { return $x } } return $null }

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

    $pa = Aba $wb 'Painel'
    $est = Aba $wb 'Estat*'
    $db = Aba $wb 'DB_Resultados'
    foreach ($w in @($pa, $est, $db)) {
        if ($w -ne $null -and $w.ProtectContents) { try { $w.Unprotect($SENHA) } catch { } }
    }
    $nomeEst = $est.Name

    # ---------------- 1. larguras C:J ----------------
    $ajust = 0
    for ($c = 3; $c -le 10; $c++) {
        if ([Math]::Abs($pa.Columns($c).ColumnWidth - 12) -gt 0.01) { $pa.Columns($c).ColumnWidth = 12; $ajust++ }
    }
    "1. Painel C:J -> largura 12 ($ajust coluna(s) ajustada(s); as demais ja estavam)"

    # ---------------- 2. mescla O3:X4 ----------------
    "2. Painel!O3:X4 -> MANTIDA (mescla=$($pa.Range('O3:X4').MergeCells); sem referencia externa, sem nome, sem shape)"

    # ---------------- 3. separador das validacoes ----------------
    # Escreve com ';' e LE DE VOLTA contando os itens. O criterio de sucesso e o
    # numero de itens que o Excel reconhece, nao o texto que eu mandei.
    $alvos = @(
        @{ Aba = $db; End = 'C4:C1048576'; Itens = @('1', '2') },
        @{ Aba = $est; End = 'B4'; Itens = @('ANUAL', 'SEMESTRAL', 'TRIMESTRAL', 'MENSAL', 'PERSONALIZADO') },
        @{ Aba = $est; End = 'F4'; Itens = @('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12') }
    )
    $nVal = 0
    foreach ($a in $alvos) {
        $rg = $a.Aba.Range($a.End)
        $lista = ($a.Itens -join ';')
        try { $rg.Validation.Delete() } catch { }
        # 3=xlValidateList, 1=xlValidAlertStop, 1=xlBetween
        $rg.Validation.Add(3, 1, 1, $lista) | Out-Null
        $rg.Validation.IgnoreBlank = $true
        $rg.Validation.InCellDropdown = $true
        $nVal++
        "3. $($a.Aba.Name)!$($a.End) -> $($a.Itens.Count) itens"
    }

    # ---------------- 4. F4: o valor era a lista inteira ----------------
    $f4antes = "$($est.Range('F4').Value2)"
    if ($f4antes -like '*,*' -or $f4antes -eq '') {
        $est.Range('F4').Value2 = [double]1
        "4. Estatistica!F4 -> 1 (continha o texto '$f4antes')"
    }
    else { "4. Estatistica!F4 -> ja continha '$f4antes', preservado" }

    # ---------------- 5. B10: residuo do layout antigo ----------------
    $b10antes = "$($est.Range('B10').Value2)"
    try { $est.Range('B10').Validation.Delete() } catch { }
    if ($b10antes -ne '') { $est.Range('B10').ClearContents() | Out-Null }
    "5. Estatistica!B10 -> limpo (continha '$b10antes' + dropdown de Visao no campo da exclusao 3)"

    # ---------------- 6. nomes definidos para as celulas reais ----------------
    $corr = @{
        'Estat_Visao'       = 'B4'
        'Estat_Ano'         = 'D4'
        'Estat_Parte'       = 'F4'
        'Estat_Lote'        = 'H4'
        'Data_Inicio_Visao' = 'B5'
        'Data_Fim_Visao'    = 'D5'
    }
    foreach ($n in $corr.Keys) {
        try { $wb.Names.Item($n).Delete() } catch { }
        $col = $corr[$n].Substring(0, 1)
        $lin = $corr[$n].Substring(1)
        $wb.Names.Add($n, "='$nomeEst'!`$$col`$$lin") | Out-Null
    }
    "6. nomes redirecionados: $($corr.Count) (Estat_Ini_Efetiva/Fim_Efetiva/Exclusoes NAO tocados)"

    # ---------------- 7. B11: o "Periodo efetivo" tinha perdido a formula ----
    #
    # A prova 4.1 acusou Estatistica!B11 AUSENTE. B11 e a celula que ESCREVE POR
    # EXTENSO a janela realmente aplicada -- inclusive as exclusoes. Sem ela, o
    # usuario nao tem como saber que ha um periodo sendo descartado do calculo.
    #
    # Isso deixou de ser hipotetico: o arquivo chegou com uma exclusao ATIVA de
    # 05/01 a 12/01/2026 em B8:C8, e as 578 divergencias de valor da prova 4.2
    # sao consequencia dela. Com a formula de volta, essa exclusao passa a estar
    # ESCRITA na tela em vez de agir em silencio.
    $b11 = "$($est.Range('B11').Formula)"
    if ($b11 -notlike '=*') {
        $est.Range('B11').Formula = '=PeriodoEfetivo(Estat_Ini_Efetiva,Estat_Fim_Efetiva,Estat_Exclusoes)'
        "7. Estatistica!B11 -> formula do periodo efetivo restaurada (estava vazia)"
    }
    else { "7. Estatistica!B11 -> ja tinha formula, preservada" }

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # ================= CONFERENCIA =================
    $erros = @()

    for ($c = 3; $c -le 10; $c++) {
        if ([Math]::Abs($pa.Columns($c).ColumnWidth - 12) -gt 0.01) { $erros += "Painel col $c largura $($pa.Columns($c).ColumnWidth)" }
    }
    if (-not $pa.Range('O3:X4').MergeCells) { $erros += 'Painel!O3:X4 deixou de ser mescla' }

    foreach ($a in $alvos) {
        $f1 = ''
        try { $f1 = "$($a.Aba.Range($a.End).Validation.Formula1)" } catch { $erros += "$($a.Aba.Name)!$($a.End) ficou sem validacao"; continue }
        $sep = if ($f1 -like '*;*') { ';' } else { ',' }
        $n = ($f1 -split [regex]::Escape($sep)).Count
        if ($n -ne $a.Itens.Count) { $erros += "$($a.Aba.Name)!$($a.End): Excel reconhece $n item(ns), esperado $($a.Itens.Count) -- F1='$f1'" }
    }

    $f4 = "$($est.Range('F4').Value2)"
    if ($f4 -like '*,*') { $erros += "F4 ainda contem a lista: '$f4'" }
    if ("$($est.Range('B10').Value2)" -ne '') { $erros += "B10 ainda tem valor: '$($est.Range('B10').Value2)'" }
    # NAO testar por .Validation.Type.
    #
    # Apos Validation.Delete() o Excel NAO passa a levantar erro em .Type: ele
    # devolve o tipo "somente entrada", e o try/catch nunca dispara. A prova
    # acusava validacao numa celula ja limpa. O criterio util e a LISTA: se
    # Formula1 esta vazia, nao ha dropdown.
    $f1b10 = ''
    try { $f1b10 = "$($est.Range('B10').Validation.Formula1)" } catch { }
    if ($f1b10.Trim() -ne '') { $erros += "B10 ainda tem lista de validacao: '$f1b10'" }

    foreach ($n in $corr.Keys) {
        $r = ''
        try { $r = "$($wb.Names.Item($n).RefersTo)" } catch { $erros += "nome $n sumiu"; continue }
        if ($r -notlike "*`$$($corr[$n].Substring(0,1))`$$($corr[$n].Substring(1))") { $erros += "nome $n -> '$r', esperado $($corr[$n])" }
    }
    # os nomes que o motor usa nao podem ter mudado
    foreach ($par in @(@('Estat_Ini_Efetiva', 'H5'), @('Estat_Fim_Efetiva', 'J5'))) {
        $r = "$($wb.Names.Item($par[0]).RefersTo)"
        if ($r -notlike "*`$$($par[1].Substring(0,1))`$$($par[1].Substring(1))") { $erros += "$($par[0]) mudou: '$r'" }
    }

    # o motor tem de continuar produzindo numero
    $hdr = 0
    for ($r = 1; $r -le 40; $r++) { if ("$($est.Cells($r,1).Value2)".Trim() -eq 'Analito') { $hdr = $r; break } }
    $comN = 0
    for ($r = $hdr + 1; $r -le $hdr + 80; $r++) {
        if ("$($est.Cells($r,1).Value2)".Trim() -eq '') { continue }
        $n = $est.Cells($r, 3).Value2
        if ("$n" -ne '' -and [double]$n -gt 0) { $comN++ }
    }
    if ($comN -lt 60) { $erros += "so $comN linha(s) com n>0 apos os ajustes (esperado ~62)" }

    $txtPer = "$($est.Range('B11').Value2)"
    if ($txtPer -eq '' -or $txtPer -like '*nao definido*') { $erros += "B11 (periodo efetivo) nao calculou: '$txtPer'" }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 12 | ForEach-Object { "  FALHA: $_" }
        throw "Fase 0 recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }
    ""
    "conferencia: larguras ok, mescla mantida, 3 validacoes com o numero certo de itens, F4=$f4, B10 limpa, 6 nomes redirecionados, $comN linhas com n"

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
