# limpar_dados_teste.ps1 - deixa a planilha pronta para o ambiente real
#
# O QUE APAGA (dado de demonstracao)
#   DB_Resultados A:G   todos os resultados
#   Registros / RegistrosStore   ocorrencias e repeticoes
#   Liberacao / LiberStore       assinaturas
#   Importar                     area de colagem
#   Analitos E:J   medias e DP por nivel      -- so no escopo Tudo
#   LotesStore     blocos de spec por lote    -- so no escopo Tudo
#   Configuracao   registro de lotes          -- so no escopo Tudo
#
# O QUE NUNCA APAGA
#   a lista de analitos (nome, area, unidade, casas decimais)
#   TEa CLIA, CV Fab, CVi, CVg e o perfil de desempenho -- sao constantes
#     regulatorias e dados de referencia, nao dado ficticio
#   usuarios, senhas e rubricas
#   qualquer formula
#
# POR QUE UMA ROTINA, E NAO "selecionar e apagar"
#
# Apagar so o DB_Resultados deixa residuo coerente com nada: Registros continua
# apontando ocorrencia de corrida que nao existe mais, e a Liberacao continua
# assinada para um lote sem resultado. Um sistema de CQI com historico
# inconsistente e pior do que um vazio -- e a inconsistencia so aparece na
# auditoria, quando ja e tarde.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\limpar_dados_teste.ps1 -Workbook <x.xlsm> [-Escopo Tudo|Movimento]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [ValidateSet('Tudo', 'Movimento')][string]$Escopo = 'Movimento'
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

# Aba por nome NORMALIZADO.
#
# "Liberacao" e "Configuracao" tem cedilha e til. Este arquivo e .ps1, que o
# Windows PowerShell 5.1 le como ANSI: o nome acentuado chega corrompido ao COM
# e o Excel devolve DISP_E_BADINDEX -- indice invalido -- como se a aba nao
# existisse. Comparar sem acento resolve sem depender da codificacao do arquivo.
function Norm {
    param([string]$s)
    if (-not $s) { return '' }
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark') { [void]$sb.Append($ch) }
    }
    return $sb.ToString().ToUpperInvariant()
}

function Aba {
    param($Pasta, [string]$Nome)
    $alvo = Norm $Nome
    foreach ($ws in @($Pasta.Worksheets)) { if ((Norm $ws.Name) -eq $alvo) { return $ws } }
    throw "aba '$Nome' nao encontrada"
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
$xl.Calculation = -4135     # manual: 56.000 formulas nao podem recalcular a cada limpeza

try {
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ProtectContents) {
            try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } }
        }
    }

    $relato = @()

    # ---- resultados ------------------------------------------------------
    $db = Aba $wb 'DB_Resultados'
    $ult = $db.Cells.Item($db.Rows.Count, 1).End(-4162).Row
    if ($ult -ge 4) {
        # SO as colunas do banco. O bloco BA:BD e formula e desce ate a linha
        # 15.003; apagar linha inteira levaria essas formulas junto.
        $db.Range($db.Cells.Item(4, 1), $db.Cells.Item($ult, 7)).ClearContents()
        $relato += "DB_Resultados : $($ult - 3) resultados apagados"
    }
    else { $relato += "DB_Resultados : ja estava vazio" }

    # ---- ocorrencias -----------------------------------------------------
    $rg = Aba $wb 'Registros'
    $visAntes = $rg.Visible; $rg.Visible = -1
    $rg.Range($rg.Cells.Item(4, 2), $rg.Cells.Item(203, 13)).ClearContents()
    $rg.Visible = $visAntes
    $rs = Aba $wb 'RegistrosStore'
    $vs = $rs.Visible; $rs.Visible = -1
    $rs.Range($rs.Cells.Item(2, 1), $rs.Cells.Item($rs.Rows.Count, 12)).ClearContents()
    $rs.Visible = $vs
    $relato += "Registros     : ocorrencias e armazem por lote limpos"

    # ---- assinaturas -----------------------------------------------------
    $lb = Aba $wb 'Liberacao'
    $vl = $lb.Visible; $lb.Visible = -1
    $lb.Range($lb.Cells.Item(4, 3), $lb.Cells.Item(203, 6)).ClearContents()
    $lb.Visible = $vl
    $ls = Aba $wb 'LiberStore'
    $vls = $ls.Visible; $ls.Visible = -1
    $ls.Range($ls.Cells.Item(1, 1), $ls.Cells.Item($ls.Rows.Count, 4)).ClearContents()
    $ls.Visible = $vls
    $relato += "Liberacao     : assinaturas limpas"

    # ---- area de colagem -------------------------------------------------
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.Name -eq 'Importar') {
            $ws.Range($ws.Cells.Item(5, 2), $ws.Cells.Item(204, 40)).ClearContents()
            $relato += "Importar      : area de colagem limpa"
        }
    }

    # ---- vitrine ---------------------------------------------------------
    $rv = Aba $wb 'Resultados'
    $rv.Range($rv.Cells.Item(4, 2), $rv.Cells.Item(303, 40)).ClearContents()
    $relato += "Resultados    : vitrine limpa"

    if ($Escopo -eq 'Tudo') {
        # ---- medias e DP por nivel (E..J). K..P PERMANECE: TEa CLIA, CV Fab,
        # CVi, CVg e o perfil sao referencia, nao dado de teste.
        $an = Aba $wb 'Analitos'
        $an.Range($an.Cells.Item(4, 5), $an.Cells.Item(43, 10)).ClearContents()
        $relato += "Analitos      : medias e DP apagadas (TEa CLIA e referencias mantidas)"

        # ---- armazem de spec por lote
        $st = Aba $wb 'LotesStore'
        $vst = $st.Visible; $st.Visible = -1
        $st.Range($st.Cells.Item(2, 1), $st.Cells.Item($st.Rows.Count, 14)).ClearContents()
        $st.Visible = $vst
        $relato += "LotesStore    : blocos de especificacao por lote limpos"

        # ---- registro de lotes
        $cfg = Aba $wb 'Configuracao'
        $rngL = $wb.Names.Item('regLoteCol').RefersToRange
        $rngL.ClearContents()
        try { $wb.Names.Item('loteAtivo').RefersToRange.ClearContents() } catch { }
        try { $wb.Names.Item('loteCarregado').RefersToRange.ClearContents() } catch { }
        $relato += "Configuracao  : registro de lotes limpo (cadastre o lote real)"
    }

    # ---- conferencia -----------------------------------------------------
    $sobrou = @()
    $u2 = $db.Cells.Item($db.Rows.Count, 1).End(-4162).Row
    if ($u2 -ge 4) { $sobrou += "DB_Resultados ainda com $($u2 - 3) linha(s)" }

    $an2 = Aba $wb 'Analitos'
    $nAn = 0
    for ($i = 4; $i -le 43; $i++) {
        $v = $an2.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -ne '') { $nAn++ }
    }
    if ($nAn -lt 31) { $sobrou += "a lista de analitos ficou com $nAn (esperado 31)" }

    $teaOK = $false
    for ($i = 4; $i -le 43; $i++) {
        if ($an2.Cells.Item($i, 11).Value2 -ne $null) { $teaOK = $true; break }
    }
    if (-not $teaOK) { $sobrou += "TEa CLIA foi apagado por engano"

    }
    $us = Aba $wb 'Usuarios'
    $vu = $us.Visible; $us.Visible = -1
    $nUsr = 0
    for ($i = 2; $i -le 60; $i++) { if ("$($us.Cells.Item($i,1).Value2)".Trim() -ne '') { $nUsr++ } }
    $us.Visible = $vu
    if ($nUsr -eq 0) { $sobrou += "nenhum usuario restou" }

    $relato | ForEach-Object { "  $_" }
    "  ----"
    "  analitos preservados : $nAn"
    "  usuarios preservados : $nUsr"

    if ($sobrou.Count -gt 0) {
        $sobrou | ForEach-Object { "  PROBLEMA: $_" }
        throw "conferencia falhou: $($sobrou.Count) problema(s). Arquivo NAO salvo."
    }
    "  conferencia: ok"

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()
    if ($wb.ProtectStructure -eq $false) { $wb.Protect($SENHA, $true, $false) }
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
