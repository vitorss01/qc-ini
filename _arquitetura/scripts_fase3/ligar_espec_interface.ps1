# ligar_espec_interface.ps1 - Painel e Estatistica passam a ler o motor
#
# Fecha o ADR-022. Ate aqui existiam DUAS fontes de verdade para o ETp:
#   - mEspecificacoes, com historico por ano (novo)
#   - Analitos!R, cascata fixa sem ano (antigo)
# e a interface lia a ANTIGA. Conviver assim e a divergencia que o ADR-019
# existe para impedir, e ja custou caro neste projeto.
#
# O QUE FAZ
#   1. cria Eng_Especificacoes -- camada de saida, mesmo padrao do Eng_Saida
#   2. nomeia engEspAnalito / engCVtp / engBIAStp / engETp
#   3. redireciona TODA formula que referencia Analitos!$R$4:$R$43 para engETp
#   4. liga AtualizarEngEspec ao AtualizarOperacao
#
# A substituicao e cirurgica: so em celulas que citam $R$4:$R$43, e dentro
# dessas o Analitos!$A$4:$A$43 vira engEspAnalito. A coluna A da Analitos e
# usada em outros lugares (o spinner do Painel, por exemplo) e nao pode ser
# trocada em bloco.
#
# As colunas K..W da Analitos NAO sao apagadas nesta etapa: viram exibicao
# residual. Apagar formula em massa sem o diff da suite e como este projeto
# perdeu 4.362 celulas uma vez.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\ligar_espec_interface.ps1 -Workbook <x.xlsm>

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

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
    return $null
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
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ProtectContents) { try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } } }
    }

    # ---- 1. Eng_Especificacoes ------------------------------------------
    $eng = Aba $wb 'Eng_Especificacoes'
    if ($eng -ne $null) { $eng.Visible = -1; $eng.Delete() }
    $eng = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $eng.Name = 'Eng_Especificacoes'

    $eng.Range('A1').Value2 = 'Ano de contexto:'
    $eng.Range('C1').Value2 = 'Fonte:'
    $eng.Range('A2').Value2 = 'Saida do motor de especificacoes. NAO editar: reescrito a cada AtualizarOperacao.'
    $eng.Range('A2').Font.Italic = $true
    $cab = @('Analito', 'Fonte', 'Ano vigente', 'CVtp %', 'BIAStp %', 'ETp %', 'Rigor', 'ID')
    for ($k = 0; $k -lt $cab.Count; $k++) {
        $c = $eng.Cells.Item(3, $k + 1); $c.Value2 = [string]$cab[$k]; $c.Font.Bold = $true
    }
    $eng.Range($eng.Cells.Item(3, 1), $eng.Cells.Item(3, 8)).Interior.Color = 14277081
    $eng.Columns.Item(1).ColumnWidth = 26
    $eng.Columns.Item(2).ColumnWidth = 20
    $eng.Columns.Item(8).ColumnWidth = 14

    # ---- 2. nomes --------------------------------------------------------
    foreach ($n in @('engEspAnalito', 'engCVtp', 'engBIAStp', 'engETp', 'engEspAno')) {
        try { $wb.Names.Item($n).Delete() } catch { }
    }
    $wb.Names.Add('engEspAnalito', '=Eng_Especificacoes!$A$4:$A$43') | Out-Null
    $wb.Names.Add('engCVtp',       '=Eng_Especificacoes!$D$4:$D$43') | Out-Null
    $wb.Names.Add('engBIAStp',     '=Eng_Especificacoes!$E$4:$E$43') | Out-Null
    $wb.Names.Add('engETp',        '=Eng_Especificacoes!$F$4:$F$43') | Out-Null
    $wb.Names.Add('engEspAno',     '=Eng_Especificacoes!$B$1') | Out-Null
    "Eng_Especificacoes criada; nomes engEspAnalito/engCVtp/engBIAStp/engETp"

    # ---- 3. redireciona os consumidores ----------------------------------
    $alvo = 'Analitos!$R$4:$R$43'
    $alvoA = 'Analitos!$A$4:$A$43'
    $tocadas = @()
    # Find do Excel, nao varredura celula a celula.
    #
    # Iterar SpecialCells(xlCellTypeFormulas) parece natural e e inviavel aqui:
    # o DB_Resultados sozinho tem ~45.000 formulas e cada leitura de .Formula e
    # uma travessia COM. Mesma licao do UpsertResultados. O Find e indexado
    # pelo proprio Excel e devolve so o que interessa.
    foreach ($ws in @($wb.Worksheets)) {
        $cel = $null
        try { $cel = $ws.Cells.Find($alvo, [System.Reflection.Missing]::Value, -4123, 2) } catch { }
        if ($cel -eq $null) { continue }
        $primeiro = $cel.Address($true, $true)
        do {
            $f = $cel.Formula
            if ($f -like "*$alvo*") {
                $cel.Formula = $f.Replace($alvo, 'engETp').Replace($alvoA, 'engEspAnalito')
                $tocadas += "$($ws.Name)!$($cel.Address($false,$false))"
            }
            $cel = $ws.Cells.FindNext($cel)
            if ($cel -eq $null) { break }
        } while ($cel.Address($true, $true) -ne $primeiro)
    }
    if ($tocadas.Count -eq 0) { throw "nenhuma formula consumia Analitos!R -- verificar antes de seguir" }
    "formulas redirecionadas para o motor: $($tocadas.Count)"
    "  $((($tocadas | Select-Object -First 6) -join ', '))$(if($tocadas.Count -gt 6){' ...'})"

    # ---- 4. liga ao AtualizarOperacao ------------------------------------
    $vbp = $wb.VBProject
    foreach ($c in $vbp.VBComponents) {
        if ($c.Name -eq 'mOperacao') {
            $cm = $c.CodeModule
            $txt = $cm.Lines(1, $cm.CountOfLines)
            if ($txt -notmatch 'AtualizarEngEspec') {
                $ini = $cm.ProcBodyLine('AtualizarOperacao', 0)
                $cm.InsertLines($ini + 1, "    AtualizarEngEspec        ' meta vigente antes de quem a consome")
                "AtualizarOperacao passa a atualizar o Eng_Especificacoes"
            }
            else { "AtualizarOperacao ja chamava AtualizarEngEspec" }
        }
    }

    $eng.Visible = 2      # veryHidden: saida de motor nao e tela

    $xl.Calculation = -4105
    $xl.CalculateFullRebuild()
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
