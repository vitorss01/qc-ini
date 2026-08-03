# criar_audit_log.ps1 - Sprint HARDENING 3, itens 3.1 e 3.2
#
# Cria a aba Audit_Log: trilha de auditoria append-only, ENCADEADA POR HASH.
#
# Cada linha carrega o SHA-256 da linha anterior somado ao proprio conteudo.
# Alterar ou apagar uma linha quebra a cadeia, e mAuditoria.VerificarIntegridadeLog
# aponta em que linha quebrou.
#
# POR QUE A CADEIA E NECESSARIA. A protecao de planilha do Excel nao resiste a
# fraude deliberada: a senha esta em texto no projeto VBA e a marcacao
# <sheetProtection> pode ser removida descompactando o .xlsm. Sem a cadeia, o
# log seria "confie em mim". Com ela, adulteracao vira evidencia verificavel.
#
# A aba fica xlSheetVeryHidden e protegida. Isso NAO e a garantia -- e a
# primeira barreira. A garantia e a cadeia.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_audit_log.ps1 -Workbook <copia_de_trabalho.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'

# Espelha as constantes de mAuditoria.bas. Mudou aqui, muda la.
$AUDIT_R0 = 4

$colunas = @(
    'ID_Auditoria', 'Data/Hora da operacao', 'Acao', 'Origem',
    'RUN', 'Data da corrida', 'Nivel', 'Analito', 'Lote', 'Resultado',
    'Status antes', 'Status depois', 'Parecer Tecnico',
    'Usuario do sistema', 'Usuario Office', 'Usuario Windows',
    'Computador', 'Arquivo', 'Versao', 'Hash anterior', 'Hash'
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)

try {
    # ---- idempotencia: remove versao anterior ----
    # ATENCAO: so e seguro recriar porque o build parte sempre de uma copia
    # limpa da producao, que ainda nao tem log. Nunca rodar isto sobre um
    # arquivo em uso -- apagaria a trilha.
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq 'Audit_Log') {
            $ws.Visible = -1
            $usadas = $ws.UsedRange.Rows.Count
            if ($usadas -ge $AUDIT_R0) {
                "AVISO: Audit_Log anterior tinha $($usadas - $AUDIT_R0 + 1) registro(s) e sera recriada"
            }
            $ws.Delete()
            break
        }
    }

    $au = $wb.Worksheets.Add()
    $au.Name = 'Audit_Log'

    # ---- cabecalho ----
    $au.Range('A1').Value2 = 'TRILHA DE AUDITORIA - append-only, encadeada por hash SHA-256'
    $au.Range('A1').Font.Bold = $true
    $au.Range('A2').Value2 = 'Nao editar manualmente. Qualquer alteracao quebra a cadeia e e detectada por ConferirAuditoria.'
    $au.Range('A2').Font.Italic = $true

    for ($i = 0; $i -lt $colunas.Count; $i++) {
        $c = $au.Cells.Item($AUDIT_R0 - 1, $i + 1)
        $c.Value2 = $colunas[$i]
        $c.Font.Bold = $true
    }
    $cab = $au.Range($au.Cells.Item($AUDIT_R0 - 1, 1), $au.Cells.Item($AUDIT_R0 - 1, $colunas.Count))
    $cab.Interior.Color = 14277081       # cinza claro
    $au.Rows.Item($AUDIT_R0 - 1).AutoFilter() | Out-Null

    # Texto nas colunas de hash: 64 digitos hexadecimais nunca devem virar
    # notacao cientifica (um hash so de digitos seria convertido em numero e a
    # verificacao falharia por falso positivo).
    $au.Columns.Item(20).NumberFormat = '@'
    $au.Columns.Item(21).NumberFormat = '@'
    $au.Columns.Item(1).NumberFormat = '@'

    $au.Columns.Item(2).ColumnWidth = 19
    $au.Columns.Item(13).ColumnWidth = 45      # parecer tecnico
    $au.Columns.Item(20).ColumnWidth = 14
    $au.Columns.Item(21).ColumnWidth = 14

    $au.Rows.Item($AUDIT_R0 - 1).Font.Size = 9

    # ---- fecha ----
    # NAO chamar Select() aqui. Selecionar exige a aba ativa; escondendo-a em
    # seguida, a pasta e salva com a celula ativa numa aba xlSheetVeryHidden, e
    # o Application.Run seguinte morre com "nao e possivel mover o foco para o
    # controle porque esta invisivel". Ativar uma aba visivel antes de esconder.
    $wb.Worksheets.Item('Painel').Activate()

    $au.Visible = 2      # xlSheetVeryHidden
    $au.Protect('qcini2025', $false, $true, $true, $true)

    $wb.Save()
    "Audit_Log criada: $($colunas.Count) colunas, dados a partir da linha $AUDIT_R0"
    "  encadeamento : SHA-256 (bloco genese = 64 zeros)"
    "  visibilidade : xlSheetVeryHidden"
    "  protecao     : ativa (barreira, nao garantia)"
}
finally {
    $wb.Close($true)
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
