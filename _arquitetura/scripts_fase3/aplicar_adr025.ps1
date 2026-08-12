# aplicar_adr025.ps1 - etapa de BUILD do ADR-025
#
# POR QUE ESTA ETAPA EXISTE
#
# A conversao BA:BC de formula para valor e o redimensionamento dos nomes r*
# viviam apenas no instalar_capacidade60m.py -- um patch aplicado A MAO sobre a
# producao. O resultado foi que a producao tinha o ADR-025 e o artefato do build
# nao, e as provas 1.1 e 1.13 passavam porque rodavam contra o arquivo remendado.
# Um build so e build se reconstroi o produto inteiro a partir da fonte.
#
# DIVISAO DE PAPEIS
#   instalar_capacidade60m.py  MIGRACAO, uma vez, sobre um arquivo que ja tem
#                              historico: captura gabarito, converte e PROVA a
#                              equivalencia contra as formulas antigas.
#   aplicar_adr025.ps1 (aqui)  BUILD, toda vez: o artefato nasce sem as formulas
#                              expansivas e com os nomes dimensionados. Nao ha
#                              gabarito a capturar porque nao ha formula antiga
#                              -- mBanco.AtualizarFlagsBanco e a unica fonte.
#
# O modulo mBanco entra pelo aplicar_vba.ps1, junto dos demais. Aqui so se faz o
# que e de DADO: tirar as formulas e mandar o motor gravar os valores.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\aplicar_adr025.ps1 -Workbook <alvo.xlsm>

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'
$C_BA = 53
$C_BC = 55

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

try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    # ---- 0. o modulo tem de estar la ------------------------------------
    $temBanco = $false
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mBanco') { $temBanco = $true } }
    if (-not $temBanco) {
        throw "mBanco ausente no artefato -- aplicar_vba.ps1 precisa importa-lo ANTES desta etapa"
    }
    "mBanco presente no artefato"

    $db = Aba $wb 'DB_Resultados'
    $protEstava = $db.ProtectContents
    if ($protEstava) { $db.Unprotect($SENHA) }

    $ult = $db.Cells.Item($db.Rows.Count, 1).End(-4162).Row
    "banco: ultima linha com dado = $ult"

    # ---- 1. fora as formulas expansivas ---------------------------------
    # Toda a coluna, nao so ate $ult: o provisionamento antigo ia ate 15.003 e
    # deixar formula orfa abaixo do dado reintroduziria o custo n^2 que esta
    # etapa existe para eliminar.
    $antes = 0
    try { $antes = $db.Range($db.Cells.Item(4, $C_BA), $db.Cells.Item($db.Rows.Count, $C_BC)).SpecialCells(-4123).Count } catch { }
    $db.Range($db.Cells.Item(4, $C_BA), $db.Cells.Item($db.Rows.Count, $C_BC)).ClearContents()
    "formulas BA:BC removidas (havia $antes celula(s) de formula)"

    # ---- 2. o motor grava os valores e redimensiona os nomes ------------
    $xl.Run('AtualizarFlagsBanco') | Out-Null
    "AtualizarFlagsBanco executada"

    # ---- 3. conferencias que impedem um artefato mudo de passar ---------
    if ($ult -ge 4) {
        $depois = 0
        try { $depois = $db.Range($db.Cells.Item(4, $C_BA), $db.Cells.Item($db.Rows.Count, $C_BC)).SpecialCells(-4123).Count } catch { }
        if ($depois -gt 0) { throw "restaram $depois formula(s) em BA:BC" }

        if ($db.Cells.Item($ult, $C_BA).HasFormula) { throw "BA$ult ainda e formula" }
        $vBB = $db.Cells.Item($ult, 54).Value2
        if ($vBB -eq $null -or "$vBB" -eq '') {
            # so e aceitavel se a ultima linha nao estiver Ativa
            $st = "$($db.Cells.Item($ult, 7).Value2)".Trim()
            if ($st -eq 'Ativo') { throw "BB$ult vazio numa linha Ativa" }
        }

        foreach ($n in @('rRUN', 'rData', 'rNivel', 'rAnalito', 'rValor', 'rStatus', 'rLote', 'rFirst', 'rRunUnico')) {
            $ref = $wb.Names.Item($n).RefersTo
            if ($ref -notmatch "\`$$ult(\D|$)") { throw "nome $n nao acompanhou a ultima linha ($ult): $ref" }
        }
        "9 nomes r* dimensionados ate a linha $ult"

        # prova independente, pelo caminho ingenuo do proprio modulo
        $r = $xl.Run('ConferirFlagsBanco', 3000)
        "ConferirFlagsBanco: $r"
        if (($r -split '\|')[1] -ne '0') { throw "ConferirFlagsBanco acusou divergencia: $r" }
    }
    else {
        "banco vazio: nada a converter (nomes ficam na linha 4)"
    }

    # ---- 3b. AtualizarOperacao passa a refrescar as flags ----------------
    #
    # O vigia Worksheet_Change cobre a edicao manual, mas so quando os eventos
    # estao ligados. Automacao com EnableEvents=False -- a propria suite faz
    # isso ao restaurar um Status apos ExcluirLogico -- escaparia dele, e as
    # flags ficariam obsoletas ate a proxima gravacao. Foi o que sustentou a
    # falha da prova 4.2 mesmo depois do vigia instalado.
    #
    # Como formula, BA:BC se curavam sozinhas em qualquer recalculo. O
    # equivalente agora e pendura-las no ponto de refresh geral, ao lado de
    # AtualizarViewResultados e AtualizarEngEspec (mesmo padrao do ADR-022).
    # Custo: 0,74 s com 93.000 registros, numa rotina que ja e O(n).
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Name -eq 'mOperacao') {
            $cm = $c.CodeModule
            $txt = $cm.Lines(1, $cm.CountOfLines)
            if ($txt -notmatch 'AtualizarFlagsBanco') {
                $ini = $cm.ProcBodyLine('AtualizarOperacao', 0)
                $cm.InsertLines($ini + 1, "    AtualizarFlagsBanco      ' ADR-025: estado derivado do banco antes de quem o consome")
                "AtualizarOperacao passa a refrescar BA:BC"
            }
            else { "AtualizarOperacao ja chamava AtualizarFlagsBanco" }
        }
    }

    # Const nao e chamavel por Application.Run; UltimaLinhaCapacidade e a funcao
    # que a publica.
    $capL = $xl.Run('UltimaLinhaCapacidade')
    $livres = $xl.Run('LinhasLivres')
    "capacidade: ate a linha $capL  ($livres linha(s) livres)"
    if ([int]$capL -lt 66603) { throw "capacidade abaixo dos 60 meses (66.603 linhas): $capL" }

    if ($protEstava -and -not $db.ProtectContents) {
        $db.Protect($SENHA, $false, $true, $true, $true) | Out-Null
    }
    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
