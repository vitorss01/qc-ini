# aplicar_release_visual.ps1 - correcoes P0 de identidade/rotulo e UX P3 segura

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$Produto,
    [Parameter(Mandatory = $true)][int]$NLV
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Cor-Ole([int]$r, [int]$g, [int]$b) { return ($r + 256*$g + 65536*$b) }
function Aba($wb, [string]$nome) { try { return $wb.Worksheets.Item($nome) } catch { return $null } }
function Aba-Like($wb, [string]$padrao) {
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -like $padrao) { return $ws } }
    return $null
}

$brand900 = Cor-Ole 16 59 61
$brand500 = Cor-Ole 54 191 168
$surface = Cor-Ole 244 247 246
$border = Cor-Ole 201 212 209
$text900 = Cor-Ole 36 50 51
$warning = Cor-Ole 255 244 214
$inputBlue = Cor-Ole 0 0 255

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $xl.Workbooks.Open($Workbook)
try {
    if ($wb.ReadOnly) { throw "Somente leitura: $Workbook" }
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) { try { $ws.Unprotect($SENHA) } catch { } }

    # P0: a Bioquimica continha identidade copiada da Hematologia. Remove apenas
    # a identidade falsa. Ausencia inicial e um cadastro ainda nao preenchido,
    # nao um defeito e nem um valor a ser persistido como "PENDENTE".
    $bioIni = $null
    $bioCfg = $null
    if ($Produto -eq 'Bioquimica') {
        $bioIni = Aba-Like $wb 'In*cio'
        $bioCfg = Aba-Like $wb 'Configura*o'
        if ($null -ne $bioCfg) {
            $bioCfg.Range('C9:C11').ClearContents()
            $bioCfg.Range('C9:C11').Locked = $false
            $bioCfg.Range('B24').Value2 = 'Informe o codigo do lote sem prefixo/sufixo de nivel.'
            $bioCfg.Range('C25').Value2 = 'Codigo do lote'
        }
        if ($null -ne $bioIni) {
            if ($null -ne $bioCfg) {
                # CELULA VAZIA REFERENCIADA VIRA ZERO.
                #
                # O artefato sai SEM identidade preenchida (C9:C11 acabaram de
                # ser limpas acima), e "=Configuracao!C9" apontando para vazio
                # renderiza 0. A tela de abertura exibia "Equipamento: 0" e
                # "Controle: 0" -- e era isso que testar_release_visual_bio.ps1
                # acusava a cada build, corretamente.
                #
                # O par de aspas vem de [char]34 e nao de aspas escapadas: numa
                # string do PowerShell cada aspa literal exige duas, e a conta
                # sai errada com facilidade -- foi o que aconteceu na primeira
                # tentativa, que gravou formula invalida em silencio.
                $vz = [string][char]34 + [string][char]34
                $nc = $bioCfg.Name
                $bioIni.Range('C7').Formula = "=IF('$nc'!C9=$vz,$vz,'$nc'!C9)"
                $bioIni.Range('C8').Formula = "=IF('$nc'!C11=$vz,$vz,'$nc'!C11)"
            }
            else {
                $bioIni.Range('C7:C8').ClearContents()
            }
        }
    }

    # P0: os cinco rotulos precisam corresponder ao motor unico.
    $painel = Aba $wb 'Painel'
    if ($null -ne $painel) {
        $painel.Range('M6').Value2 = '1-3S'
        $painel.Range('N6').Value2 = '2-2S'
        $painel.Range('O6').Value2 = 'R4S'
        $painel.Range('P6').Value2 = '4-1S'
        $painel.Range('Q6').Value2 = '10X'
        $painel.Range('M6:Q6').Interior.Color = $brand900
        $painel.Range('M6:Q6').Font.Color = Cor-Ole 255 255 255
        $painel.Range('M6:Q6').Font.Bold = $true
    }

    $lib = Aba-Like $wb 'Libera*o'
    if ($null -ne $lib) {
        $lib.Range('A2').Value2 = 'Duplo-clique em INICIADO para a primeira conferencia. FINALIZADO exige ANALISTA/ADM, rubrica valida e CQI aprovado pelo motor.'
        $lib.Range('A2:F2').Interior.Color = $warning
        $lib.Range('A2:F2').WrapText = $true
    }

    # Datas sem mascara literal defeituosa.
    foreach ($nome in @('DB_Resultados','Resultados','Corridas','Audit_Log','EQC_Dados')) {
        $ws = Aba $wb $nome
        if ($null -eq $ws) { continue }
        $ult = [Math]::Max(4, $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1)
        if ($nome -eq 'Corridas') {
            $ws.Range("B4:B$ult").NumberFormatLocal = 'dd/mm/aaaa;@'
            $ws.Range("C4:C$ult").NumberFormat = 'hh:mm:ss'
        } elseif ($nome -eq 'Audit_Log') {
            $ws.Range("C4:C$ult").NumberFormatLocal = 'dd/mm/aaaa hh:mm:ss;@'
            $ws.Range("D4:D$ult").NumberFormatLocal = 'dd/mm/aaaa;@'
            $ws.Range("E4:E$ult").NumberFormat = 'hh:mm:ss'
        } else {
            $ws.Range("B4:B$ult").NumberFormatLocal = 'dd/mm/aaaa;@'
        }
    }

    # P3: sistema visual discreto e consistente, sem alterar geometria dos
    # graficos, formulas, validacoes ou areas de entrada.
    foreach ($ws in @($wb.Worksheets)) {
        $ws.UsedRange.Font.Name = 'Aptos'
        $ws.UsedRange.Font.Color = $text900
        try { $ws.Application.ActiveWindow.DisplayGridlines = $false } catch { }
        try { $ws.Application.ActiveWindow.Zoom = 90 } catch { }
        if ($ws.UsedRange.Rows.Count -gt 0) {
            $ws.Rows.Item(1).Font.Bold = $true
            $ws.Rows.Item(1).Font.Color = Cor-Ole 255 255 255
            $ws.Rows.Item(1).Interior.Color = $brand900
        }
    }

    foreach ($padrao in @('In*cio','Painel','Estat*stica','Libera*o')) {
        $ws = Aba-Like $wb $padrao; if ($null -ne $ws) { $ws.Tab.Color = $brand500 }
    }
    foreach ($nome in @('Analitos','Resultados','Importar','Registros')) {
        $ws = Aba $wb $nome; if ($null -ne $ws) { $ws.Tab.Color = $brand900 }
    }
    $ws = Aba-Like $wb 'Configura*o'; if ($null -ne $ws) { $ws.Tab.Color = $brand900 }
    foreach ($nome in @('DB_Resultados','Corridas','Audit_Log','BI_Data','Eng_Saida','Cfg_Status','Eventos_Westgard')) {
        $ws = Aba $wb $nome; if ($null -ne $ws) { $ws.Tab.Color = $border }
    }

    $bi = Aba $wb 'BI_Data'
    if ($null -ne $bi) {
        $bi.Rows.Item(1).Interior.Color = $brand900
        $bi.Rows.Item(1).Font.Color = Cor-Ole 255 255 255
        $bi.Rows.Item(1).Font.Bold = $true
        $bi.Columns.AutoFit() | Out-Null
    }

    # Reaplica a convencao visual de entrada depois do passe global de fontes.
    # C9:C11 permanecem editaveis quando a planilha for protegida; Inicio e
    # somente uma exibicao vinculada ao cadastro, portanto continua bloqueada.
    if ($Produto -eq 'Bioquimica') {
        if ($null -ne $bioCfg) {
            $bioCfg.Range('C9:C11').Interior.Color = Cor-Ole 255 255 255
            $bioCfg.Range('C9:C11').Font.Color = $inputBlue
            $bioCfg.Range('C9:C11').Borders.LineStyle = 1
            $bioCfg.Range('C9:C11').Locked = $false
        }
        if ($null -ne $bioIni) {
            $bioIni.Range('C7:C8').Interior.Color = $surface
            $bioIni.Range('C7:C8').Font.Color = $text900
            $bioIni.Range('C7:C8').Locked = $true
        }
    }

    $wb.Save()
    "P0 identidade/rotulos/data: aplicado"
    "P3 sistema visual: aplicado sem alterar formulas"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
