# blindar_artefato.ps1 - itens 3.3 e 3.5 do Quality Gate
#
# O PROBLEMA. A protecao do sistema hoje e aplicada em tempo de execucao:
# Workbook_Open chama LockApp, que esconde as abas e chama ReprotectAll com
# UserInterfaceOnly:=True. Duas consequencias:
#
#   1. Workbook_Open NAO RODA com macros desabilitadas. Quem abre o arquivo com
#      macros desligadas -- o cenario do auditor, e tambem o do curioso -- pega o
#      arquivo no estado em que foi salvo.
#   2. UserInterfaceOnly NAO PERSISTE. E um atributo de sessao: some ao fechar.
#
# Como o arquivo vinha sendo salvo em sessao autenticada (ADM roda UnprotectAll),
# ele era distribuido DESTRAVADO: nenhuma das 18 abas tinha <sheetProtection> e
# DB_Resultados ficava a um clique de distancia.
#
# A CORRECAO. O artefato sai do build ja no estado travado, com a protecao
# GRAVADA NO ARQUIVO:
#   - toda aba protegida por senha (sem UserInterfaceOnly, para persistir)
#   - toda aba xlSheetVeryHidden, menos Login
#   - estrutura da pasta de trabalho protegida
#
# Com macros habilitadas nada muda para o usuario: Workbook_Open chama LockApp,
# o login roda e UnlockApp revela o que o papel permite -- e ReprotectAll
# reaplica a protecao com UserInterfaceOnly:=True, que e o que deixa o VBA
# gravar.
#
# O QUE ISTO NAO RESOLVE. Protecao de planilha do Excel e barreira, nao cofre: a
# senha esta em texto no projeto VBA e a marcacao pode ser removida
# descompactando o .xlsm. Por isso a trilha de auditoria e ENCADEADA POR HASH --
# a garantia de integridade esta la, nao aqui. Ver mAuditoria.bas.
#
# O item 3.4 (senha do PROJETO VBA) NAO e feito aqui: exige escrever os campos
# DPB/DPx do vbaProject.bin, o que corrompe o projeto com facilidade. E um passo
# manual no VBE: Ferramentas > Propriedades do VBAProject > Protecao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\blindar_artefato.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'

$SENHA = 'qcini2025'
$LOGIN = 'Login'

$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

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
    # A aba de Login precisa estar visivel ANTES de esconder as outras: uma pasta
    # de trabalho nao pode ficar sem nenhuma aba visivel.
    $wsLogin = $null
    foreach ($ws in $wb.Worksheets) { if ($ws.Name -eq $LOGIN) { $wsLogin = $ws; break } }
    if ($wsLogin -eq $null) { throw "Aba $LOGIN nao encontrada - nao da para travar o arquivo com seguranca" }
    $wsLogin.Visible = -1        # xlSheetVisible
    $wsLogin.Activate()

    $protegidas = 0
    $escondidas = 0
    foreach ($ws in $wb.Worksheets) {
        # Protecao SEM UserInterfaceOnly: e esta que fica gravada no arquivo.
        try { $ws.Unprotect($SENHA) } catch { }
        $ws.Protect($SENHA, $true, $true, $true)     # estrutura, conteudo, objetos, cenarios
        $protegidas++

        if ($ws.Name -ne $LOGIN) {
            $ws.Visible = 2      # xlSheetVeryHidden
            $escondidas++
        }
    }

    # Estrutura protegida: impede reexibir aba pelo menu de contexto.
    try { $wb.Unprotect($SENHA) } catch { }
    $wb.Protect($SENHA, $true, $false)

    $wb.Save()

    "abas protegidas   : $protegidas"
    "abas ocultas      : $escondidas (todas menos $LOGIN)"
    "estrutura         : protegida"
    "estado distribuido: TRAVADO (protecao gravada no arquivo, nao so em tempo de execucao)"
    "ATENCAO item 3.4  : a senha do projeto VBA e passo MANUAL no VBE"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
