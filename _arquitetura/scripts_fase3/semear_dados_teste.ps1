# semear_dados_teste.ps1 - da ao ARTEFATO os dados que a suite precisa
#
# POR QUE ISTO PASSOU A SER NECESSARIO
#
# A suite sempre exercitou o sistema sobre os 1.000 resultados ficticios que
# moravam no arquivo de PRODUCAO. Funcionava, e escondia uma falha de projeto:
# uma suite de regressao nao pode depender de dado de demonstracao vivendo no
# arquivo de trabalho do laboratorio.
#
# Em 05/08/2026 a producao foi limpa para o uso real, e a suite ficou cega: o
# artefato nasceu vazio, a verificacao 3.6 leu a linha 4 do DB_Resultados,
# recebeu tudo vazio e travou o Excel chamando UpsertResultados com um registro
# degenerado (RUN 0, analito "", data vazia).
#
# Agora a fixture pertence ao BUILD, nao a producao. O laboratorio trabalha com
# a planilha limpa; o artefato de verificacao nasce com dado suficiente para os
# controles terem o que exercitar.
#
# O lote de teste tem nucleo 999999 -- impossivel de confundir com lote real.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\semear_dados_teste.ps1 -Workbook <build.xlsm> [-NLV 2] [-Corridas 25]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [int]$NLV = 2,
    [int]$Corridas = 25
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'
$LOTE = '999999'

function Norm { param([string]$s)
    if (-not $s) { return '' }
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark') { [void]$sb.Append($ch) }
    }
    return $sb.ToString().ToUpperInvariant()
}
function Aba { param($P, [string]$N)
    $a = Norm $N
    foreach ($ws in @($P.Worksheets)) { if ((Norm $ws.Name) -eq $a) { return $ws } }
    throw "aba '$N' nao encontrada"
}
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
$xl.Calculation = -4135

try {
    # RESTAURAR o estado de entrada, nao impor um.
    #
    # normalizar_protecao.ps1 deixa a estrutura DESTRAVADA de proposito: os
    # passos seguintes do build criam abas. Reproteger aqui fez criar_eng_saida
    # falhar logo depois -- uma etapa impondo estado que a seguinte precisava
    # que continuasse aberto.
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ProtectContents) { try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } } }
    }

    # ---- lote de teste no registro --------------------------------------
    $rngL = $wb.Names.Item('regLoteCol').RefersToRange
    $rngL.NumberFormat = '@'
    $rngL.Cells.Item(1, 1).Value2 = $LOTE
    try { $wb.Names.Item('loteAtivo').RefersToRange.Value2 = $LOTE } catch { }
    try { $wb.Names.Item('loteCarregado').RefersToRange.Value2 = $LOTE } catch { }
    "lote de teste registrado: $LOTE"

    # ---- especificacoes minimas -----------------------------------------
    # Sem meta, os blocos que avaliam conformidade nao teriam o que comparar.
    $an = Aba $wb 'Analitos'
    $nAn = 0
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -eq $null -or "$v".Trim() -eq '') { continue }
        $nAn++
        if ("$($an.Cells.Item($i,5).Value2)".Trim() -eq '') {
            $an.Cells.Item($i, 5).Value2 = [double](80 + ($nAn * 3) % 40)    # media N1
            $an.Cells.Item($i, 6).Value2 = [double](2 + ($nAn % 3))          # DP N1
            for ($nv2 = 2; $nv2 -le $NLV; $nv2++) {
                $cM = 5 + ($nv2 - 1) * 2
                $an.Cells.Item($i, $cM).Value2 = [double](100 * $nv2 + ($nAn * 5) % 90)
                $an.Cells.Item($i, $cM + 1).Value2 = [double](2 * $nv2 + ($nAn % 4))
            }
        }
        if ("$($an.Cells.Item($i,11).Value2)".Trim() -eq '') { $an.Cells.Item($i, 11).Value2 = [double]15 }
    }
    "analitos com media/DP de teste: $nAn"

    # ---- resultados ------------------------------------------------------
    $db = Aba $wb 'DB_Resultados'
    $analitos = @()
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -ne '') { $analitos += "$v" }
    }
    if ($analitos.Count -eq 0) { throw "nenhum analito cadastrado -- nada a semear" }

    $nLin = $Corridas * $NLV * $analitos.Count
    $dados = New-Object 'object[,]' $nLin, 7
    # [datetime]::new, nao Get-Date.
    #
    # Get-Date -Hour 0 -Minute 0 -Second 0 NAO zera os milissegundos: a data
    # carrega o instante do relogio. Duas execucoes da semeadura produziam
    # datas com fracao de segundo diferente, e o diff de formulas acusava 50
    # divergencias de valor -- fixture nao deterministica deixa a suite
    # instavel para sempre. A semente do gerador aleatorio ja era fixa; faltava
    # o relogio.
    $base = [datetime]::new(2026, 1, 6, 0, 0, 0)
    $k = 0
    $rnd = New-Object System.Random 20260805      # semente fixa: build reprodutivel
    for ($corr = 1; $corr -le $Corridas; $corr++) {
        $dt = $base.AddDays(7 * ($corr - 1))
        for ($nv = 1; $nv -le $NLV; $nv++) {
            for ($a = 0; $a -lt $analitos.Count; $a++) {
                $lin = $an.Cells.Item(4 + $a, 1)
                # Media/DP POR NIVEL: E=5/F=6 (N1), G=7/H=8 (N2), I=9/J=10 (N3).
                #
                # Antes era "5/6 se nivel 1, senao 7/8" -- e o nivel 3 nascia com
                # a media do nivel 2. Na Hematologia (NLV=3, produto PADRAO do
                # build) o WBC tem N2=6,8 e N3=16,5: os resultados do N3 caiam
                # dezenas de desvios fora do alvo. A fixture inteira do nivel 3
                # era lixo, e e sobre ela que a suite mede o aceite.
                $cMed = 5 + ($nv - 1) * 2
                $cDp = $cMed + 1
                $media = [double]$an.Cells.Item(4 + $a, $cMed).Value2
                $dp = [double]$an.Cells.Item(4 + $a, $cDp).Value2
                if ($dp -le 0) { $dp = 1 }
                # normal por Box-Muller: dado de teste tem de PARECER controle,
                # senao Westgard nunca dispara e as regras nao sao exercitadas.
                $u1 = $rnd.NextDouble(); $u2 = $rnd.NextDouble()
                if ($u1 -le 0) { $u1 = 0.0001 }
                $z = [Math]::Sqrt(-2 * [Math]::Log($u1)) * [Math]::Cos(2 * [Math]::PI * $u2)
                $dados[$k, 0] = [double]$corr
                $dados[$k, 1] = $dt
                $dados[$k, 2] = [double]$nv
                $dados[$k, 3] = 'QC-' + $LOTE + ('{0:D2}' -f $nv)
                $dados[$k, 4] = [string]$analitos[$a]
                $dados[$k, 5] = [double][Math]::Round($media + $z * $dp, 4)
                $dados[$k, 6] = 'Ativo'
                $k++
            }
        }
    }
    $db.Range($db.Cells.Item(4, 4), $db.Cells.Item(3 + $nLin, 4)).NumberFormat = '@'
    $db.Range($db.Cells.Item(4, 1), $db.Cells.Item(3 + $nLin, 7)).Value = $dados
    "resultados semeados: $nLin ($Corridas corridas x $NLV niveis x $($analitos.Count) analitos)"

    # ---- flags derivadas (ADR-025) ---------------------------------------
    #
    # Desde o ADR-025 as colunas BA/BB/BC nao sao mais formula: sao valor
    # mantido por mBanco. Escrever direto em A:G, como esta rotina faz, deixaria
    # o nucleo do lote e as flags rFirst/rRunUnico VAZIOS -- e a fixture nasceria
    # invisivel para o Calc, o Painel e os graficos. Rapido e errado.
    #
    # AtualizarFlagsBanco tambem redimensiona os intervalos nomeados r* para a
    # ultima linha real, o que esta rotina precisa igualmente.
    # O modulo pode AINDA nao existir, e isso nao e erro.
    #
    # A semeadura roda no passo 0c; mBanco so e importado no passo 4, e o 4a
    # chama AtualizarFlagsBanco logo em seguida. Na Bioquimica isso passava
    # despercebido porque a PRODUCAO ja tem o modulo (ADR-025 foi aplicado la);
    # na Hematologia, nao -- e o build inteiro morria no passo 0c.
    #
    # Ausencia do modulo => aviso e segue, porque o 4a corrige.
    # Modulo presente que FALHA => continua fatal: ai a fixture ficaria mesmo
    # fora dos calculos e ninguem consertaria depois.
    $temBanco = $false
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mBanco') { $temBanco = $true } }
    if (-not $temBanco) {
        "mBanco ainda nao importado (passo 4) -- flags serao geradas no passo 4a"
    }
    else {
        try {
            $xl.Run('AtualizarFlagsBanco') | Out-Null
            "flags BA/BB/BC e intervalos nomeados atualizados por mBanco"
        }
        catch {
            throw "AtualizarFlagsBanco falhou -- fixture ficaria fora dos calculos: $($_.Exception.Message)"
        }
    }

    # ---- alinhar os FILTROS ao dado semeado ------------------------------
    #
    # Semear o dado nao basta: as telas tem estado de filtro PERSISTIDO, herdado
    # da producao, e ele nao conhece a fixture.
    #
    # Era o que estava acontecendo: o artefato nascia com loteAtivo = 999999 (o
    # lote de teste) mas com Estat_Lote = 8974 (o lote REAL da producao), e a
    # Estatistica devolvia n=0 para os 31 analitos. O Painel, pela mesma razao,
    # ficava presa no mes 'Jan' enquanto a fixture vai de janeiro a junho -- as
    # linhas de limite do grafico viravam #N/A da corrida 5 em diante.
    #
    # Nenhum dos dois e defeito do produto: e a bancada medindo com o filtro
    # fechado. Mas o diff acusava dezenas de divergencias de VALOR e o ruido
    # escondia o que a suite existe para ver.
    $nucleo = $LOTE
    $est = Aba $wb 'Estatistica'
    $est.Range('H4').Value2 = $nucleo              # Estat_Lote
    $est.Range('B4').Value2 = 'ANUAL'
    $est.Range('D4').Value2 = 2026
    "Estatistica alinhada: lote $nucleo, visao ANUAL de 2026"

    $pai = Aba $wb 'Painel'
    $pai.Range('I3').Value2 = 2026
    $pai.Range('I4').Value2 = '(todos)'
    # De/Ate VAZIOS, e nao um ano inteiro.
    #
    # A primeira versao escrevia 01/01/2026 a 31/12/2026 em G3/G4. Funcionava e
    # era pior: o Painel passava a exibir "FILTRO ATIVO: 01/01/26 a 31/12/26", e
    # a fixture ficava dependendo de a janela cobrir as datas semeadas. Sem
    # filtro nenhum o Calc enxerga tudo, que e o que a bancada quer, e o rotulo
    # continua igual ao da referencia.
    $pai.Range('G3').ClearContents() | Out-Null
    $pai.Range('G4').ClearContents() | Out-Null
    "Painel alinhado: sem filtro de data (a fixture vai de jan a jun de 2026)"

    # ---- conferencia -----------------------------------------------------
    $ult = $db.Cells.Item($db.Rows.Count, 1).End(-4162).Row
    if ("$($db.Cells.Item(4,54).Value2)".Trim() -eq '') { throw "linha 4 sem flag rFirst -- mBanco nao rodou" }
    if ($ult -lt 4) { throw "nada foi gravado no DB_Resultados" }
    if ("$($db.Cells.Item(4,5).Value2)".Trim() -eq '') { throw "linha 4 sem analito -- a suite le exatamente essa linha" }
    if ("$($db.Cells.Item(4,7).Value2)" -ne 'Ativo') { throw "linha 4 sem status Ativo" }
    "conferencia: linha 4 = RUN $($db.Cells.Item(4,1).Value2) / $($db.Cells.Item(4,5).Value2) / $($db.Cells.Item(4,7).Value2)"

    $xl.Calculation = -4105
    $xl.CalculateFullRebuild()
    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { $xl.Calculation = -4105 } catch { }
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
