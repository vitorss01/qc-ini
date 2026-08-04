# criar_audit_log.ps1 - itens 3.1 e 3.2
#
# Cria Audit_Log como BASE DE AUDITORIA, nao como arquivo tecnico.
#
# O auditor abre esta aba e responde tudo com os filtros nativos do Excel, sem
# conhecer a estrutura do sistema e sem nenhuma tela em VBA:
#   - so as alteracoes do WBC          - so um RUN, um lote, um nivel
#   - so o que fulano alterou          - so o que veio de tal computador
#   - so entre duas datas              - so entre duas horas
#   - so exclusoes                     - so mudancas de configuracao
#
# TRES DECISOES QUE FAZEM ISSO FUNCIONAR:
#
# 1. TABELA DO EXCEL (ListObject "tblAuditoria"). Sem ela os filtros nao
#    acompanham as linhas novas -- justamente os eventos que o auditor procura
#    primeiro -- e Power Query / Power BI nao tem o que referenciar pelo nome.
#
# 2. Data e Hora em COLUNAS PROPRIAS, alem do timestamp completo. O filtro do
#    Excel sobre data-hora nao resolve "entre 8h e 12h em qualquer dia".
#
# 3. Protecao com AllowFiltering e AllowSorting. Aba protegida sem essas
#    permissoes vira bloco morto e o objetivo se perde inteiro.
#
# A aba fica VISIVEL e protegida: e material de auditoria, feito para ser lido.
# A integridade nao vem de esconder -- vem da cadeia de hash (mAuditoria.bas).
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_audit_log.ps1 -Workbook <copia_de_trabalho.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

$AUDIT_R0 = 4        # espelha mAuditoria.bas
$SENHA = 'qcini2025'

# nome, largura
$colunas = @(
    @('ID_Evento', 22), @('VersaoSchema', 8), @('Timestamp', 18), @('Data', 11), @('Hora', 10),
    @('Categoria', 14), @('Acao', 24), @('Modulo', 14), @('ChaveRegistro', 18),
    @('RUN', 7), @('DataCorrida', 12), @('Equipamento', 14), @('Lote', 15),
    @('Nivel', 7), @('Analito', 12), @('SeqAlteracao', 8),
    @('ResultadoAnterior', 16), @('ResultadoNovo', 16), @('Delta', 11), @('DeltaPerc', 10),
    @('StatusAnterior', 15), @('StatusNovo', 15),
    @('Motivo', 26), @('ParecerTecnico', 48),
    @('UsuarioSistema', 16), @('Papel', 11), @('UsuarioOffice', 18), @('UsuarioWindows', 18),
    @('Computador', 18), @('Arquivo', 22), @('VersaoSistema', 13),
    @('HashAnterior', 14), @('Hash', 14)
)

$legenda = @(
    @('CATEGORIA', ''),
    @('DADO', 'Alteracao em resultado de controle: inclusao, alteracao, exclusao logica.'),
    @('CONFIGURACAO', 'Alteracao na tabela de elegibilidade (Cfg_Status) ou em parametro do sistema.'),
    @('SEGURANCA', 'Login, troca de senha, mudanca de papel de usuario.'),
    @('SISTEMA', 'Eventos internos do sistema.'),
    @('', ''),
    @('ACAO', ''),
    @('RESULTADO_INCLUIDO', 'Resultado gravado pela primeira vez para aquela chave RUN|Nivel|Analito.'),
    @('RESULTADO_ALTERADO', 'Valor numerico substituido por outro valor numerico.'),
    @('RESULTADO_APAGADO', 'Valor substituido por vazio. ResultadoNovo aparece como <VAZIO>.'),
    @('RESULTADO_EXCLUIDO', 'Exclusao LOGICA: a linha permanece no banco e sai dos calculos.'),
    @('REENVIO_BLOQUEADO', 'Tentativa de reenviar registro nao ativo. Recusada: nao ressuscita excluido.'),
    @('CFG_ELEGIBILIDADE_ALTERADA', 'Um estado passou a entrar (ou a nao entrar) na estatistica.'),
    @('CFG_ESTADO_INCLUIDO', 'Novo estado acrescentado a tabela de elegibilidade.'),
    @('CFG_ESTADO_REMOVIDO', 'Estado retirado da tabela de elegibilidade.'),
    @('', ''),
    @('COLUNAS QUE MERECEM EXPLICACAO', ''),
    @('ChaveRegistro', 'RUN|Nivel|Analito. Filtre por ela para ver TODA a vida de um resultado.'),
    @('SeqAlteracao', '1a, 2a, 3a vez que aquele resultado foi mexido. O valor ORIGINAL esta sempre em Seq = 1.'),
    @('Delta / DeltaPerc', 'Diferenca entre o valor novo e o anterior. Vazio quando um dos lados nao e numero.'),
    @('Motivo', 'Lista fechada: permite CONTAR (ex.: quantas correcoes foram erro de digitacao).'),
    @('ParecerTecnico', 'Texto livre do analista, minimo de 5 palavras. Explica o caso concreto.'),
    @('UsuarioSistema', 'Login autenticado no QC_INI (senha em hash). E a identidade DE REGISTRO.'),
    @('UsuarioOffice / UsuarioWindows / Computador', 'Corroboracao. Vem de variavel de ambiente e pode ser alterada pelo usuario.'),
    @('HashAnterior / Hash', 'Encadeamento SHA-256. Alterar ou apagar qualquer linha quebra a cadeia e e detectado.'),
    @('', ''),
    @('COMO CONFERIR A INTEGRIDADE', 'Executar a macro ConferirAuditoria. Ela recalcula a cadeia inteira e aponta a linha exata onde houver quebra.'),
    @('LIMITE DECLARADO', 'A hora vem do relogio da maquina, que o usuario pode alterar. Nao ha solucao dentro do Excel.')
)

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
    # Idempotencia. So e seguro recriar porque o build parte sempre de uma copia
    # limpa da producao, que ainda nao tem log. Nunca rodar sobre arquivo em uso.
    foreach ($nome in @('Audit_Log', 'Audit_Legenda')) {
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -eq $nome) {
                $ws.Visible = -1
                $usadas = $ws.UsedRange.Rows.Count
                if ($nome -eq 'Audit_Log' -and $usadas -ge $AUDIT_R0) {
                    "AVISO: Audit_Log anterior tinha $($usadas - $AUDIT_R0 + 1) evento(s) e sera recriada"
                }
                $ws.Delete()
                break
            }
        }
    }

    # ---------------------------------------------------------------- log ----
    $au = $wb.Worksheets.Add()
    $au.Name = 'Audit_Log'

    $au.Range('A1').Value2 = 'TRILHA DE AUDITORIA - QC_INI'
    $au.Range('A1').Font.Bold = $true
    $au.Range('A1').Font.Size = 13
    $au.Range('A2').Value2 = 'Append-only e encadeada por hash SHA-256. Uma linha = um evento. Nao editar: qualquer alteracao quebra a cadeia e e detectada pela macro ConferirAuditoria. Use os filtros da tabela para investigar. Ver aba Audit_Legenda.'
    $au.Range('A2').Font.Italic = $true

    for ($i = 0; $i -lt $colunas.Count; $i++) {
        $c = $au.Cells.Item($AUDIT_R0 - 1, $i + 1)
        $c.Value2 = $colunas[$i][0]
        $au.Columns.Item($i + 1).ColumnWidth = $colunas[$i][1]
    }

    # Texto nas colunas de hash e de ID: 64 digitos hexadecimais so de numeros
    # virariam notacao cientifica e a verificacao falharia por falso positivo.
    $au.Columns.Item(1).NumberFormat = '@'
    $au.Columns.Item(32).NumberFormat = '@'
    $au.Columns.Item(33).NumberFormat = '@'

    # ListObject: e o que faz o filtro acompanhar as linhas novas e o que
    # Power Query / Power BI referenciam pelo nome.
    $rngTab = $au.Range($au.Cells.Item($AUDIT_R0 - 1, 1), $au.Cells.Item($AUDIT_R0, $colunas.Count))
    $lo = $au.ListObjects.Add(1, $rngTab, $null, 1)    # xlSrcRange, xlYes
    $lo.Name = 'tblAuditoria'
    $lo.TableStyle = 'TableStyleMedium2'

    $au.Rows.Item($AUDIT_R0 - 1).Font.Size = 9
    $au.Activate()
    $xl.ActiveWindow.FreezePanes = $false
    $au.Range('A4').Select() | Out-Null
    $xl.ActiveWindow.FreezePanes = $true

    # ------------------------------------------------------------ legenda ----
    $lg = $wb.Worksheets.Add()
    $lg.Name = 'Audit_Legenda'
    $lg.Range('A1').Value2 = 'COMO LER A TRILHA DE AUDITORIA'
    $lg.Range('A1').Font.Bold = $true
    $lg.Range('A1').Font.Size = 13
    $lg.Range('A2').Value2 = 'Esta pagina existe porque quem audita nao conhece a estrutura interna do sistema.'
    $lg.Range('A2').Font.Italic = $true
    for ($i = 0; $i -lt $legenda.Count; $i++) {
        $lg.Cells.Item($i + 4, 1).Value2 = $legenda[$i][0]
        $lg.Cells.Item($i + 4, 2).Value2 = $legenda[$i][1]
        if ($legenda[$i][1] -eq '' -and $legenda[$i][0] -ne '') {
            $lg.Cells.Item($i + 4, 1).Font.Bold = $true
        }
    }
    $lg.Columns.Item(1).ColumnWidth = 42
    $lg.Columns.Item(2).ColumnWidth = 95
    $lg.Columns.Item(2).WrapText = $false
    $lg.Range('A1').Select() | Out-Null

    # ---- protecao QUE PERMITE FILTRAR E ORDENAR ----
    $au.Protect($SENHA, $true, $true, $true, $true, $true, $true, $true, $true, $false, $false, $false, $true, $true)
    $lg.Protect($SENHA, $true, $true, $true)

    $wb.Worksheets.Item('Painel').Activate()
    $wb.Save()

    "Audit_Log criada  : $($colunas.Count) colunas, tabela tblAuditoria, dados a partir da linha $AUDIT_R0"
    "  filtros         : nativos do Excel (AllowFiltering + AllowSorting)"
    "  data e hora     : colunas proprias, alem do timestamp completo"
    "  encadeamento    : SHA-256, schema 2"
    "Audit_Legenda     : $($legenda.Count) linhas de explicacao para quem audita"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
