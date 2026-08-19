# criar_estrutura_especificacoes.ps1 - estrutura formal de especificacoes
#
# Cria Cfg_Especificacoes, DB_Especificacoes e Eng_Especificacoes sem semear
# fontes, rigor ou limites. Na Hematologia, Analitos!Q:R continua sendo o
# fallback operacional ate existir cadastro formal utilizavel.
#
# Idempotente: se as abas ja existem, valida o schema e preserva os dados.
# Recusa executar sobre um canonico.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$Produto
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Workbook -notmatch '(?i)(_EM_DESENVOLVIMENTO|build_hardening)') {
    throw "Recusado: estrutura de especificacoes so pode ser criada em clone _EM_DESENVOLVIMENTO ou build_hardening."
}
if ($Produto -ne 'Hematologia') {
    throw "Este passo entrega a pendencia da Hematologia. Produto recebido: $Produto"
}

$SENHA = 'qcini2025'

function Aba($wb, [string]$nome) {
    try { return $wb.Worksheets.Item($nome) } catch { return $null }
}

function Nova-Aba($wb, [string]$nome) {
    $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $nome
    return $ws
}

function Conferir-Cabecalho($ws, [int]$linha, [string[]]$esperado) {
    for ($i = 0; $i -lt $esperado.Count; $i++) {
        $atual = [string]$ws.Cells.Item($linha, $i + 1).Value2
        if ($atual -ne $esperado[$i]) {
            throw "$($ws.Name)!$($ws.Cells.Item($linha, $i + 1).Address($false,$false)): esperado '$($esperado[$i])', encontrado '$atual'."
        }
    }
}

function Nome-Definido($wb, [string]$nome, [string]$refere) {
    try { $wb.Names.Item($nome).Delete() } catch { }
    $wb.Names.Add($nome, $refere) | Out-Null
}

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null; Start-Sleep 5 } catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

$salvou = $false
$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }

try {
    if ($wb.ReadOnly) { throw "Somente leitura: $Workbook" }
    $estruturaProtegida = $wb.ProtectStructure
    if ($estruturaProtegida) { $wb.Unprotect($SENHA) }

    # ---- Cfg_Especificacoes --------------------------------------------
    $cfg = Aba $wb 'Cfg_Especificacoes'
    if ($null -eq $cfg) {
        $cfg = Nova-Aba $wb 'Cfg_Especificacoes'
        $cfg.Range('A1').Value2 = 'ESPECIFICACOES DE QUALIDADE - CONFIGURACAO'
        $cfg.Range('A2').Value2 = 'Fonte padrao'
        $cfg.Range('A3').Value2 = 'Rigor padrao (VB)'
        $cfg.Range('A4').Value2 = 'Ano de referencia (0 = ano do resultado)'
        $cfg.Range('B2:B3').ClearContents()
        $cfg.Range('B4').Value2 = [double]0
        $cfg.Range('A6').Value2 = 'FONTE'
        $cfg.Range('B6').Value2 = 'MODELO'
        $cfg.Range('C6').Value2 = 'DESCRICAO / REFERENCIA INTERNA'

        # Catalogo tecnico de modelos. Nao e cadastro de fonte nem especificacao.
        $cfg.Range('E6').Value2 = 'MODELO SUPORTADO'
        $cfg.Range('F6').Value2 = 'ENTRADAS'
        $cfg.Range('G6').Value2 = 'SAIDAS DERIVADAS'
        $modelos = @(
            @('ETP_DIRETO', 'ETp %', 'CVtp = ETp/3'),
            @('VB', 'CVi %, CVg %, rigor', 'CVtp, BIAStp, ETp'),
            @('CV_BIAS_DIRETO', 'CVtp %, BIAStp %', 'ETp')
        )
        for ($i = 0; $i -lt $modelos.Count; $i++) {
            for ($j = 0; $j -lt 3; $j++) { $cfg.Cells.Item(7 + $i, 5 + $j).Value2 = [string]$modelos[$i][$j] }
        }
        "Cfg_Especificacoes criada sem fonte ou rigor predefinido"
    }
    else {
        Conferir-Cabecalho $cfg 6 @('FONTE','MODELO','DESCRICAO / REFERENCIA INTERNA')
        "Cfg_Especificacoes existente: schema preservado"
    }

    try { $cfg.Unprotect($SENHA) } catch { }
    $cfg.Cells.Locked = $true
    $cfg.Range('B2:B4').Locked = $false
    $cfg.Range('A7:C40').Locked = $false
    $cfg.Range('A1:G1').Font.Bold = $true
    $cfg.Range('A6:G6').Font.Bold = $true
    $cfg.Range('A6:C6').Interior.Color = 14277081
    $cfg.Range('E6:G6').Interior.Color = 14277081
    $cfg.Columns.Item(1).ColumnWidth = 25
    $cfg.Columns.Item(2).ColumnWidth = 19
    $cfg.Columns.Item(3).ColumnWidth = 42
    $cfg.Columns.Item(5).ColumnWidth = 20
    $cfg.Columns.Item(6).ColumnWidth = 24
    $cfg.Columns.Item(7).ColumnWidth = 22

    $sep = [string]$xl.International(5)
    foreach ($endereco in @('B2','B3','B4','B7:B40')) { try { $cfg.Range($endereco).Validation.Delete() } catch { } }
    $cfg.Range('B2').Validation.Add(3, 1, 1, '=$A$7:$A$40')
    $cfg.Range('B3').Validation.Add(3, 1, 1, "MIN${sep}DES${sep}OTI")
    $cfg.Range('B4').Validation.Add(1, 1, 1, 0, 2200)
    $cfg.Range('B7:B40').Validation.Add(3, 1, 1, "ETP_DIRETO${sep}VB${sep}CV_BIAS_DIRETO")
    $cfg.Protect($SENHA, $true, $true, $true, $true)

    # ---- DB_Especificacoes ---------------------------------------------
    $cabDB = @('ID','Ano','Fonte','Modelo','Analito','ETp %','CVi %','CVg %',
               'Rigor','CVtp %','BIAStp %','Ativo','Usuario','Cadastrado em')
    $db = Aba $wb 'DB_Especificacoes'
    if ($null -eq $db) {
        $db = Nova-Aba $wb 'DB_Especificacoes'
        $db.Range('A1').Value2 = 'ESPECIFICACOES DE QUALIDADE - BANCO'
        $db.Range('A2').Value2 = 'Uma linha por Ano + Fonte + Analito. Nenhum limite e criado automaticamente.'
        for ($i = 0; $i -lt $cabDB.Count; $i++) { $db.Cells.Item(3, $i + 1).Value2 = [string]$cabDB[$i] }
        "DB_Especificacoes criada vazia"
    }
    else {
        Conferir-Cabecalho $db 3 $cabDB
        "DB_Especificacoes existente: dados preservados"
    }
    try { $db.Unprotect($SENHA) } catch { }
    $db.Cells.Locked = $true
    $db.Range('A1:N1').Font.Bold = $true
    $db.Range('A3:N3').Font.Bold = $true
    $db.Range('A3:N3').Interior.Color = 14277081
    $db.Columns.Item(1).ColumnWidth = 13
    $db.Columns.Item(3).ColumnWidth = 22
    $db.Columns.Item(4).ColumnWidth = 18
    $db.Columns.Item(5).ColumnWidth = 25
    $db.Columns.Item(14).ColumnWidth = 18
    $db.Columns.Item(14).NumberFormat = 'dd/mm/yyyy hh:mm'
    $db.Protect($SENHA, $true, $true, $true, $true)

    # ---- Eng_Especificacoes --------------------------------------------
    $cabEng = @('Analito','Fonte','Ano vigente','CVtp %','BIAStp %','ETp %','Rigor','ID','Situacao')
    $eng = Aba $wb 'Eng_Especificacoes'
    if ($null -eq $eng) {
        $eng = Nova-Aba $wb 'Eng_Especificacoes'
        $eng.Range('A1').Value2 = 'Ano de contexto:'
        $eng.Range('C1').Value2 = 'Fonte formal:'
        $eng.Range('A2').Value2 = 'Saida derivada. Cadastro formal prevalece; na ausencia, Hematologia preserva Analitos!Q:R como fallback.'
        for ($i = 0; $i -lt $cabEng.Count; $i++) { $eng.Cells.Item(3, $i + 1).Value2 = [string]$cabEng[$i] }
        "Eng_Especificacoes criada"
    }
    else {
        Conferir-Cabecalho $eng 3 $cabEng
        "Eng_Especificacoes existente: schema preservado"
    }
    try { $eng.Unprotect($SENHA) } catch { }
    $eng.Cells.Locked = $true
    $eng.Range('A1:I1').Font.Bold = $true
    $eng.Range('A3:I3').Font.Bold = $true
    $eng.Range('A3:I3').Interior.Color = 14277081
    $eng.Columns.Item(1).ColumnWidth = 25
    $eng.Columns.Item(2).ColumnWidth = 22
    $eng.Columns.Item(8).ColumnWidth = 14
    $eng.Columns.Item(9).ColumnWidth = 28
    $eng.Protect($SENHA, $true, $true, $true, $true)

    Nome-Definido $wb 'engEspAnalito' '=Eng_Especificacoes!$A$4:$A$43'
    Nome-Definido $wb 'engCVtp' '=Eng_Especificacoes!$D$4:$D$43'
    Nome-Definido $wb 'engBIAStp' '=Eng_Especificacoes!$E$4:$E$43'
    Nome-Definido $wb 'engETp' '=Eng_Especificacoes!$F$4:$F$43'
    Nome-Definido $wb 'engEspAno' '=Eng_Especificacoes!$B$1'
    Nome-Definido $wb 'engEspSituacao' '=Eng_Especificacoes!$I$4:$I$43'

    $cfg.Visible = 2
    $db.Visible = 2
    $eng.Visible = 2
    if ($estruturaProtegida -or -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }

    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
