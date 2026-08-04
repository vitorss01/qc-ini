# instalar_cfg_watch.ps1 - item 2.5 do Quality Gate
#
# Instala o vigia da tabela de elegibilidade: o modulo de documento da aba
# Cfg_Status passa a chamar mConfig.AuditarMudancaCfg a cada alteracao, e a
# sombra e sincronizada na primeira instalacao.
#
# POR QUE UM SCRIPT E NAO UM .cls VERSIONADO. O nome de codigo da aba
# (Planilha1, Planilha7, ...) NAO e estavel entre os tres produtos nem entre
# versoes: depende da ordem em que as abas foram criadas. Amarrar o patch a
# "Planilha3" quebraria em Bioquimica e Imunologia. Aqui a aba e localizada
# PELO NOME e o componente correspondente e resolvido em tempo de execucao.
#
# ATENCAO: modulo de documento nao aceita AddFromFile com cabecalho -- VERSION,
# BEGIN e Attribute viram erro de sintaxe e derrubam a compilacao do PROJETO
# INTEIRO (invariante 9 do RC1_ARCHITECTURE_FREEZE). Por isso o codigo entra
# como string, sem cabecalho.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\instalar_cfg_watch.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

$codigo = @'
Option Explicit

' Vigia da tabela de elegibilidade (item 2.5). Ver mConfig.bas.
'
' EnableEvents = False durante a auditoria: AuditarMudancaCfg ESCREVE na propria
' aba (versao e sombra), e sem isso o evento se dispararia recursivamente.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fim
    If Intersect(Target, Me.Range("B4:C40")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Dim n As Long
    n = AuditarMudancaCfg()
    ' MsgBox so quando ha alguem para ler. Em automacao COM o Excel roda
    ' invisivel, e um dialogo modal ali prende o processo indefinidamente --
    ' foi assim que a caixa de senha travou o build por horas.
    If n > 0 And Application.Visible Then
        MsgBox n & " alteracao(oes) na tabela de elegibilidade registrada(s) na auditoria." & vbLf & vbLf & _
               "Configuracao agora na versao " & VersaoCfg() & "." & vbLf & _
               "Isto muda RETROATIVAMENTE o que compoe media, DP, CV, Bias, Sigma e Westgard " & _
               "de todo o historico.", vbInformation, "Elegibilidade alterada"
    End If
fim:
    Application.EnableEvents = True
End Sub
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
if ($wb.ReadOnly) {
    $wb.Close($false); $xl.Quit()
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}

try {
    $ws = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'Cfg_Status') { $ws = $w; break } }
    if ($ws -eq $null) { throw "Aba Cfg_Status nao encontrada" }

    $nomeCodigo = $ws.CodeName
    $comp = $wb.VBProject.VBComponents.Item($nomeCodigo)
    if ($comp.Type -ne 100) { throw "$nomeCodigo nao e modulo de documento (Type=$($comp.Type))" }

    $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
    $comp.CodeModule.AddFromString($codigo)

    $linhas = $comp.CodeModule.CountOfLines
    if ($linhas -lt 10) { throw "codigo do vigia entrou truncado ($linhas linhas)" }

    # marco zero da sombra, sem registrar nada
    $xl.Run('SincronizarSombraCfg')

    $versao = $xl.Run('VersaoCfg')

    $wb.Save()
    "aba Cfg_Status    : modulo $nomeCodigo ($linhas linhas de vigia)"
    "sombra            : sincronizada (colunas H e I, ocultas)"
    "versao inicial    : $versao"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
