# reordenar_analitos.ps1 - poe a aba Analitos na ordem do CSV e cadastra os novos
#
# O QUE ESTE SCRIPT RESOLVE, E POR QUE ELE E DELICADO
#
# A aba Analitos e a lista mestra: a ordem dela define a ordem do spinner do
# Painel, das linhas da Estatistica e das colunas da aba Importar. Reordenar
# parece cosmetico. NAO E.
#
#   LotesStore guarda as especificacoes POR POSICAO, nao por nome.
#
# Cada lote ocupa um bloco de 40 linhas (CAP_ANALITOS em mLotes.bas), e a
# posicao da linha DENTRO do bloco e a unica ligacao com o analito. Reordenar a
# aba Analitos sem permutar os blocos faria a media/DP da Glicose passar a valer
# para o analito que assumisse a posicao 1 -- em silencio, sem erro, sem aviso.
# Seria corrupcao de especificacao analitica, que e o insumo do Levey-Jennings.
#
# Por isso este script permuta as duas coisas com a MESMA permutacao, e depois
# CONFERE, lote a lote e analito a analito, que a especificacao continua com o
# dono certo. Se um unico par divergir, o arquivo nao e salvo.
#
# O resto do sistema ja e seguro para reordenacao, o que foi verificado antes de
# escrever isto:
#   Calc          - area de UM analito por vez (selAnalito); nao tem layout por analito
#   Painel        - MATCH por nome; faixas ja vao ate a linha 43 (40 analitos)
#   Estatistica   - segue a ordem de Analitos; ja dimensionada para 40
#   RegistrosStore- chaveia por NOME
#   Graficos      - leem o Calc, que e por analito selecionado
#
# A ordem e os nomes vem de src_producao/analitos_<produto>.csv, lido como
# UTF-8. Nome de analito tem acento; manter isso em .ps1 (que o PowerShell 5.1
# le como ANSI) ja corrompeu tabela neste projeto.
#
# Nao usar acentos neste arquivo.
#
# Uso:
#   .\reordenar_analitos.ps1 -Workbook <QC_Bioquimica.xlsm> -Csv <analitos_bioquimica.csv>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Csv
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if (-not $Csv) {
    $Csv = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\analitos_bioquimica.csv'
}
$Csv = (Resolve-Path -LiteralPath $Csv).ProviderPath

$SENHA = 'qcini2025'
$A_R0 = 4          # primeira linha de analito na aba Analitos
$A_CAP = 40        # capacidade (CAP_ANALITOS em mLotes.bas)
$A_RN = $A_R0 + $A_CAP - 1
$LS_R0 = 2         # primeira linha de dado em LotesStore
$LS_C0 = 3         # coluna C = primeira coluna de spec (equivale a Analitos!E)
$LS_NC = 12        # C..N  =  E..P (aInput)

# --- CSV em UTF-8 --------------------------------------------------------
$linhas = [System.IO.File]::ReadAllLines($Csv, [System.Text.Encoding]::UTF8)
$def = @()
for ($i = 1; $i -lt $linhas.Count; $i++) {
    $l = $linhas[$i].Trim()
    if ($l -eq '') { continue }
    $p = $l.Split(',')
    if ($p.Count -lt 5) { throw "linha $($i+1) do CSV mal formada: $l" }
    $def += [pscustomobject]@{
        Ordem = [int]$p[0]; Sigla = $p[1]; Nome = $p[2]; Unid = $p[3]; Casas = $p[4]
    }
}
if ($def.Count -eq 0) { throw "CSV sem definicoes" }
if ($def.Count -gt $A_CAP) { throw "CSV tem $($def.Count) analitos; a capacidade e $A_CAP" }
"definicoes no CSV : $($def.Count)"

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
# Calculation so aceita atribuicao com uma pasta ja aberta (0x800A03EC antes
# disso). Manual durante a permutacao: recalculo no meio veria a lista mestra
# em estado intermediario.
$xl.Calculation = -4135

try {
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ProtectContents) {
            try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } }
            if ($ws.ProtectContents) { throw "aba '$($ws.Name)' nao abriu com a senha do projeto" }
        }
    }

    $an = $wb.Worksheets.Item('Analitos')
    $ls = $wb.Worksheets.Item('LotesStore')

    # --- 1. estado ANTES -------------------------------------------------
    # Analitos A..P das 40 linhas
    $antes = $an.Range($an.Cells.Item($A_R0, 1), $an.Cells.Item($A_RN, 16)).Value2
    $nomeAntes = @()
    for ($i = 1; $i -le $A_CAP; $i++) {
        $v = $antes[$i, 1]
        $nomeAntes += $(if ($v -eq $null) { '' } else { "$v".Trim() })
    }
    $existentes = @($nomeAntes | Where-Object { $_ -ne '' })
    "analitos hoje     : $($existentes.Count)"

    # blocos de lote existentes
    $ultLS = $ls.Cells.Item($ls.Rows.Count, 1).End(-4162).Row
    $nBlocos = [Math]::Ceiling(($ultLS - $LS_R0 + 1) / $A_CAP)
    if ($nBlocos -lt 1) { $nBlocos = 0 }
    "blocos de lote    : $nBlocos"

    $blocoAntes = @{}
    $loteDoBloco = @{}
    for ($b = 0; $b -lt $nBlocos; $b++) {
        $r0 = $LS_R0 + $b * $A_CAP
        $loteDoBloco[$b] = "$($ls.Cells.Item($r0, 1).Value2)".Trim()
        $blocoAntes[$b] = $ls.Range($ls.Cells.Item($r0, $LS_C0), $ls.Cells.Item($r0 + $A_CAP - 1, $LS_C0 + $LS_NC - 1)).Value2
    }

    # --- 2. permutacao ---------------------------------------------------
    # origem[k] = posicao (1..40) que o analito da nova posicao k+1 ocupava
    #             antes;  0 quando e analito novo.
    $origem = @()
    $novos = @()
    foreach ($d in $def) {
        $idx = 0
        for ($i = 0; $i -lt $A_CAP; $i++) {
            if ($nomeAntes[$i] -eq $d.Nome) { $idx = $i + 1; break }
        }
        $origem += $idx
        if ($idx -eq 0) { $novos += $d.Nome }
    }
    "ja cadastrados    : $(($origem | Where-Object { $_ -ne 0 }).Count)"
    "novos a cadastrar : $($novos.Count)"
    if ($novos.Count -gt 0) { "  $($novos -join ' | ')" }

    # nenhum analito existente pode sumir: isso apagaria especificacao e
    # desligaria historico do DB_Resultados da lista mestra
    $mantidos = @()
    foreach ($o in $origem) { if ($o -ne 0) { $mantidos += $nomeAntes[$o - 1] } }
    $perdidos = @($existentes | Where-Object { $mantidos -notcontains $_ })
    if ($perdidos.Count -gt 0) {
        throw "o CSV nao contempla analito(s) ja cadastrado(s): $($perdidos -join ', '). Isso apagaria especificacao e orfanaria dados do banco."
    }

    # --- 3. monta a nova aba Analitos (A..P) -----------------------------
    $novoA = New-Object 'object[,]' $A_CAP, 16
    for ($k = 0; $k -lt $A_CAP; $k++) {
        for ($c = 0; $c -lt 16; $c++) { $novoA[$k, $c] = $null }
    }
    for ($k = 0; $k -lt $def.Count; $k++) {
        $d = $def[$k]
        $o = $origem[$k]
        if ($o -ne 0) {
            # existente: leva A..P inteiro, para nao perder nenhuma spec
            #
            # PARENTESES OBRIGATORIOS em ($c - 1). A virgula liga mais forte que
            # o menos: $novoA[$k, $c - 1] e lido como ($k, $c) - 1, ou seja um
            # ARRAY menos 1, e o erro sai como "Object[] nao contem
            # op_Subtraction" -- longe da causa. Mesma familia do @('a'+$x,'b')
            # que corrompeu o Cfg_Status.
            for ($c = 1; $c -le 16; $c++) { $novoA[$k, ($c - 1)] = $antes[$o, $c] }
        }
        else {
            # novo: identidade e formato. Media/DP/TEa ficam VAZIOS de
            # proposito -- sao dados analiticos do lote, que so o laboratorio
            # pode informar. Inventar numero aqui seria fabricar evidencia.
            $novoA[$k, 0] = $d.Nome
            $novoA[$k, 1] = 'Bioquimica'
            $novoA[$k, 2] = $d.Unid
            $novoA[$k, 3] = [double]$d.Casas
            $novoA[$k, 14] = 'MIN'      # O: perfil de desempenho (fi/fb minimo)
        }
    }
    $an.Range($an.Cells.Item($A_R0, 1), $an.Cells.Item($A_RN, 16)).Value2 = $novoA
    "aba Analitos reescrita: $($def.Count) analitos"

    # --- 4. permuta os blocos do LotesStore ------------------------------
    for ($b = 0; $b -lt $nBlocos; $b++) {
        $orig = $blocoAntes[$b]
        $novo = New-Object 'object[,]' $A_CAP, $LS_NC
        for ($k = 0; $k -lt $A_CAP; $k++) {
            for ($c = 0; $c -lt $LS_NC; $c++) { $novo[$k, $c] = $null }
        }
        for ($k = 0; $k -lt $def.Count; $k++) {
            $o = $origem[$k]
            if ($o -ne 0) {
                for ($c = 1; $c -le $LS_NC; $c++) { $novo[$k, ($c - 1)] = $orig[$o, $c] }
            }
        }
        $r0 = $LS_R0 + $b * $A_CAP
        $ls.Range($ls.Cells.Item($r0, $LS_C0), $ls.Cells.Item($r0 + $A_CAP - 1, $LS_C0 + $LS_NC - 1)).Value2 = $novo
    }
    if ($nBlocos -gt 0) { "blocos de lote permutados: $nBlocos" }

    # --- 5. atualiza a linha de/para da aba Importar ---------------------
    $imp = $null
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -eq 'Importar') { $imp = $ws } }
    if ($imp -ne $null) {
        for ($k = 0; $k -lt $def.Count; $k++) {
            $imp.Cells.Item(3, 5 + $k).Value2 = [string]$def[$k].Nome     # linha 3 = mapeamento
            $cab = $imp.Cells.Item(4, 5 + $k)
            $cab.Value2 = [string]$def[$k].Sigla
            $cab.Interior.ColorIndex = 0
            $cab.Interior.Color = 14277081
            $cab.Font.Color = 0
        }
        "aba Importar: de/para atualizado para $($def.Count) colunas"
    }

    # --- 6. CONFERENCIA: a spec continua com o dono certo? ---------------
    $depois = $an.Range($an.Cells.Item($A_R0, 1), $an.Cells.Item($A_RN, 16)).Value2
    $divergencias = @()

    # 6a. Analitos: para cada analito que ja existia, E..P tem de ser igual
    for ($i = 1; $i -le $A_CAP; $i++) {
        $nome = $(if ($antes[$i, 1] -eq $null) { '' } else { "$($antes[$i,1])".Trim() })
        if ($nome -eq '') { continue }
        $novoIdx = 0
        for ($k = 1; $k -le $A_CAP; $k++) {
            $n2 = $(if ($depois[$k, 1] -eq $null) { '' } else { "$($depois[$k,1])".Trim() })
            if ($n2 -eq $nome) { $novoIdx = $k; break }
        }
        if ($novoIdx -eq 0) { $divergencias += "Analitos: '$nome' sumiu"; continue }
        for ($c = 2; $c -le 16; $c++) {
            $va = "$($antes[$i, $c])"
            $vd = "$($depois[$novoIdx, $c])"
            if ($va -ne $vd) { $divergencias += "Analitos: '$nome' col $c  '$va' -> '$vd'" }
        }
    }

    # 6b. LotesStore: mesma spec, mesmo dono, lote a lote
    for ($b = 0; $b -lt $nBlocos; $b++) {
        $r0 = $LS_R0 + $b * $A_CAP
        $dep = $ls.Range($ls.Cells.Item($r0, $LS_C0), $ls.Cells.Item($r0 + $A_CAP - 1, $LS_C0 + $LS_NC - 1)).Value2
        $orig = $blocoAntes[$b]
        for ($i = 1; $i -le $A_CAP; $i++) {
            $nome = $(if ($antes[$i, 1] -eq $null) { '' } else { "$($antes[$i,1])".Trim() })
            if ($nome -eq '') { continue }
            $novoIdx = 0
            for ($k = 0; $k -lt $def.Count; $k++) { if ($def[$k].Nome -eq $nome) { $novoIdx = $k + 1; break } }
            if ($novoIdx -eq 0) { continue }
            for ($c = 1; $c -le $LS_NC; $c++) {
                $va = "$($orig[$i, $c])"
                $vd = "$($dep[$novoIdx, $c])"
                if ($va -ne $vd) {
                    $divergencias += "LotesStore lote $($loteDoBloco[$b]): '$nome' col $c  '$va' -> '$vd'"
                }
            }
        }
    }

    if ($divergencias.Count -gt 0) {
        $divergencias | Select-Object -First 15 | ForEach-Object { "  $_" }
        throw "CONFERENCIA FALHOU: $($divergencias.Count) divergencia(s). Arquivo NAO salvo."
    }
    "conferencia: ok - toda especificacao continua com o analito de origem"

    $xl.Calculation = -4105     # automatico
    $wb.Application.CalculateFullRebuild()
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
