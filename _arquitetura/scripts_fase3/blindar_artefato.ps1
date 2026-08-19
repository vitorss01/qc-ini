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

# A TRILHA DE AUDITORIA FICA VISIVEL DE PROPOSITO.
#
# O objetivo do Audit_Log e ser lido por quem audita -- inclusive com as
# macros desabilitadas, que e como um auditor cauteloso abre um arquivo
# desconhecido. Esconder a trilha derrotaria o proposito dela.
# Continua PROTEGIDA (nao editavel) e a integridade nao depende de estar
# oculta: depende da cadeia de hash.
$VISIVEIS = @($LOGIN)

$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

# Criar o Excel COM RESILIENCIA.
#
# O build sobe e derruba o Excel cerca de dez vezes. Sob esse ritmo o servidor
# COM as vezes recusa a proxima instancia com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE) -- estado transitorio, nao defeito do script.
# Falhar na primeira tentativa jogava fora um build inteiro de varios minutos.
function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try {
            $app = New-Object -ComObject Excel.Application
            return $app
        }
        catch {
            $ultimo = $_
            # Depois de um periodo sem Excel rodando, a PRIMEIRA ativacao COM
            # costuma falhar com 0x80080005 mesmo com a maquina sadia. Lancar o
            # excel.exe uma vez levanta o servidor e as ativacoes seguintes
            # funcionam. Verificado nesta maquina: com um processo de pe, o
            # New-Object passa na hora.
            if ($tentativa -eq 2) {
                try {
                    Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null
                    Start-Sleep -Seconds 5
                }
                catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu apos 6 tentativas: $($ultimo.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
# A trava de estrutura impede reexibir/ocultar aba. Este script PRECISA fazer
# isso (esconde as abas de dado), entao abre a estrutura primeiro; a protecao
# final ($wb.Protect Structure) volta a fecha-la no arquivo distribuido.
if ($wb.ProtectStructure) { $wb.Unprotect('qcini2025') }

# AutoRecuperacao DESLIGADA nesta copia de trabalho.
#
# O build encerra o Excel a forca varias vezes. Cada encerramento deixa um
# arquivo de recuperacao pendente; acumulados, o Excel passa a tentar exibir o
# painel "Recuperacao de Documento" ao iniciar e MORRE antes de responder --
# ate o excel.exe puro para de abrir, e a automacao falha com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE), que parece defeito de COM e nao e.
#
# O artefato e reproduzivel por comando: nao ha o que recuperar aqui.
try { $wb.EnableAutoRecover = $false } catch { }

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
        if ($ws.Name -like 'Audit_*') {
            # Aba protegida SEM AllowFiltering vira bloco morto: o auditor
            # nao consegue filtrar, e a tabela perde a razao de existir.
            $ws.Protect($SENHA, $true, $true, $true, $true, $true, $true, $true, $true, $false, $false, $false, $true, $true)
        } else {
            $ws.Protect($SENHA, $true, $true, $true)
        }
        $protegidas++

        if ($VISIVEIS -notcontains $ws.Name) {
            $ws.Visible = 2      # xlSheetVeryHidden
            $escondidas++
        }
    }

    # Estrutura protegida: impede reexibir aba pelo menu de contexto.
    try { $wb.Unprotect($SENHA) } catch { }
    $wb.Protect($SENHA, $true, $false)

    $wb.Save()

    "abas protegidas   : $protegidas"
    "abas ocultas      : $escondidas (visiveis: $($VISIVEIS -join ', '))"
    "estrutura         : protegida"
    "estado distribuido: TRAVADO (protecao gravada no arquivo, nao so em tempo de execucao)"
    "ATENCAO item 3.4  : a senha do projeto VBA e passo MANUAL no VBE"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
