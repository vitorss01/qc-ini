# criar_logs_db.ps1 - as duas tabelas de log dentro do DB_Resultados
#
# DESENHO DEFINIDO PELO GESTOR. O Audit_Log e o Event Store geral (append-only,
# encadeado por hash). Estas duas tabelas sao a estrutura FISICA desacoplada
# dentro do banco, preservando as referencias historicas ja estruturadas:
#
#   LOG_Resultados  <- exclusoes e nao conformidades originadas na aba Resultados
#   LOG_Registros   <- exclusoes originadas na aba Registros
#
# As duas tem AS MESMAS COLUNAS, de proposito: uma visao de auditoria consegue
# uni-las sem tratamento. E as duas carregam ID_Auditoria, que amarra cada linha
# ao evento correspondente no Event Store -- nenhuma das duas e fonte
# independente da verdade.
#
# ONDE FICAM. Em BLOCOS DE COLUNAS DESLOCADOS A DIREITA, nao abaixo dos dados:
# CarregarDB le o bloco A:G inteiro e UltimaLinhaBanco mede a coluna A. Uma
# tabela vizinha nessas colunas quebraria a leitura do banco -- que e o que da a
# performance atual. Deslocadas, a leitura fica intacta.
#
# O DB_Resultados e mais denso do que parece. Mapa levantado por varredura:
#   A:G    banco (rRUN, rData, rNivel, rLote, rAnalito, rValor, rStatus)
#   H      botao "Lancar corrida"
#   K:AZ   DADOS INTERFACEADOS -- Data, Nivel e os 40 analitos, com K1:AZ1 e
#          K2:AZ2 MESCLADAS
#   BA:BD  LoteCore, 1aOc, RunUnico -- com nomes definidos (rLote, rFirst,
#          rRunUnico) referenciados por formula
#
# Primeira coluna realmente livre: BE. Os blocos ficam depois dela, com folga:
#   LOG_Resultados : BG..BY  (59..77)
#   LOG_Registros  : CB..CT  (80..98)
#
# COMO LISTOBJECT, para o auditor filtrar sem VBA e para Power Query consumir
# pelo nome. Cada tabela tem seu proprio filtro, independente da outra.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_logs_db.ps1 -Workbook <copia_de_trabalho.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

$SENHA = 'qcini2025'
$LINHA_CAB = 3       # mesma linha de cabecalho do banco
$COL_LOGRES = 59     # BG
$COL_LOGREG = 80     # CB

# nome, largura. IDENTICAS nas duas tabelas para permitir uniao.
$colunas = @(
    @('ID_Auditoria', 22), @('DataHora', 18), @('TipoOperacao', 22), @('AbaOrigem', 12),
    @('RUN', 7), @('DataCorrida', 12), @('Nivel', 7), @('Analito', 12), @('Lote', 15),
    @('Resultado', 11), @('StatusAnterior', 15), @('StatusNovo', 15),
    @('ParecerTecnico', 46),
    @('UsuarioSistema', 16), @('UsuarioOffice', 18), @('UsuarioWindows', 18),
    @('Computador', 18), @('Arquivo', 22), @('VersaoSistema', 13)
)

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $ultimo = $_; Start-Sleep -Seconds ($tentativa * 2) }
    }
    throw "Excel COM nao subiu apos 6 tentativas: $($ultimo.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) {
    $wb.Close($false); $xl.Quit()
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}

try {
    $db = $wb.Worksheets.Item('DB_Resultados')
    $eraVisivel = $db.Visible
    $db.Visible = -1
    try { $db.Unprotect($SENHA) } catch { }

    $tabelas = @(
        @{ Nome = 'tblLogResultados'; Col = $COL_LOGRES; Titulo = 'LOG DE EXCLUSOES / NAO CONFORMIDADES - origem: aba Resultados' },
        @{ Nome = 'tblLogRegistros'; Col = $COL_LOGREG; Titulo = 'LOG DE EXCLUSOES - origem: aba Registros' }
    )

    foreach ($t in $tabelas) {
        # CONFERIR ANTES DE ESCREVER. O bloco so pode ser usado se estiver
        # vazio ou se ja for uma versao anterior DESTA tabela. Sem esta
        # verificacao, um erro de coordenada sobrescreveria em silencio a area
        # de dados interfaceados ou as colunas de apoio -- e o dano so
        # apareceria muito depois, como formula devolvendo vazio.
        $c0v = $t.Col
        $c1v = $t.Col + $colunas.Count - 1
        $ocupado = $false
        $jaNossa = $false
        for ($cc = $c0v; $cc -le $c1v; $cc++) {
            for ($rr = 1; $rr -le 6; $rr++) {
                $val = $db.Cells.Item($rr, $cc).Value2
                if ($val -ne $null -and "$val".Trim() -ne '') {
                    $ocupado = $true
                    if ("$val" -like '*LOG DE EXCLUSOES*' -or $colunas[0][0] -eq "$val" -or "$val" -eq 'ID_Auditoria') { $jaNossa = $true }
                }
            }
        }
        if ($ocupado -and -not $jaNossa) {
            throw "Bloco $c0v..$c1v do DB_Resultados NAO esta livre. Escolher outro intervalo em vez de sobrescrever."
        }

        # idempotencia: remove versao anterior da tabela e limpa o bloco
        foreach ($lo in @($db.ListObjects)) {
            if ($lo.Name -eq $t.Nome) { $lo.Unlist() | Out-Null }
        }
        $c0 = $t.Col
        $c1 = $t.Col + $colunas.Count - 1
        $db.Range($db.Cells.Item(1, $c0), $db.Cells.Item(500, $c1)).Clear() | Out-Null

        $db.Cells.Item($LINHA_CAB - 2, $c0).Value2 = $t.Titulo
        $db.Cells.Item($LINHA_CAB - 2, $c0).Font.Bold = $true

        for ($i = 0; $i -lt $colunas.Count; $i++) {
            $cel = $db.Cells.Item($LINHA_CAB, $c0 + $i)
            $cel.Value2 = $colunas[$i][0]
            $db.Columns.Item($c0 + $i).ColumnWidth = $colunas[$i][1]
        }
        # ID em texto: identificador nunca deve virar numero
        $db.Columns.Item($c0).NumberFormat = '@'

        $rng = $db.Range($db.Cells.Item($LINHA_CAB, $c0), $db.Cells.Item($LINHA_CAB + 1, $c1))
        $lo = $db.ListObjects.Add(1, $rng, $null, 1)      # xlSrcRange, xlYes
        $lo.Name = $t.Nome
        $lo.TableStyle = 'TableStyleLight9'

        "  $($t.Nome): colunas $c0..$c1, cabecalho na linha $LINHA_CAB"
    }

    # Protecao que PERMITE FILTRAR: sem isso o auditor nao usa as tabelas.
    $db.Protect($SENHA, $true, $true, $true, $true, $true, $true, $true, $true, $false, $false, $false, $true, $true)
    $db.Visible = $eraVisivel

    $wb.Save()
    "tabelas de log criadas no DB_Resultados: 2 x $($colunas.Count) colunas"
    "  bloco A:G do banco intacto (CarregarDB e UltimaLinhaBanco nao sao afetados)"
    "  ambas com filtro proprio e ligadas ao Event Store por ID_Auditoria"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
