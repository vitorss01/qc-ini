# criar_corridas.ps1 - registro de corridas + contador persistido do RUN
#
# Cria a aba Corridas e migra o historico. A migracao e ADITIVA: nenhum RUN
# existente muda de valor. Constroi o registro a partir dos triplos distintos
# (RUN, Data, nucleo do lote) ja gravados em DB_Resultados e posiciona o
# contador em max(RUN) + 1.
#
# Consequencia deliberada: dias que de fato tiveram duas corridas continuam
# colapsados num RUN so. Isso NAO e corrigido retroativamente -- inventar
# separacao de historico seria pior que registrar a limitacao. O modelo novo
# vale da migracao em diante.
#
# Layout:
#   D1            contador (nome definido proxRUN)
#   linha 3       cabecalhos
#   linha 4+      RUN | Data | Hora | Turno | Lote | Equipamento | Usuario |
#                 CriadoEm | CriadoPor
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_corridas.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$R0 = 4

$ErrorActionPreference = 'Stop'

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
$xl.AutomationSecurity = 3

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


# Somente leitura significa que OUTRA instancia do Excel ainda segura o arquivo.
# Sem esta guarda o script grava no vazio e reporta sucesso: DisplayAlerts=$false
# suprime o aviso do Excel, e o Save falha em silencio.
if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}


foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -eq 'Corridas') { $ws.Visible = -1; $ws.Delete(); "Corridas anterior removida"; break }
}

$cor = $wb.Worksheets.Add()
$cor.Name = 'Corridas'

$cor.Range('A1').Value2 = 'Registro de corridas'
$cor.Range('C1').Value2 = 'Proximo RUN:'
$cor.Range('A1:C1').Font.Bold = $true

$cab = @('RUN', 'Data', 'Hora', 'Turno', 'Lote', 'Equipamento', 'Usuario', 'CriadoEm', 'CriadoPor')
for ($k = 0; $k -lt $cab.Count; $k++) { $cor.Cells(3, $k + 1).Value2 = $cab[$k] }
$cor.Range($cor.Cells(3, 1), $cor.Cells(3, $cab.Count)).Font.Bold = $true

# ---- migracao: triplos distintos (RUN, Data, nucleo do lote) do banco ----
$db = $wb.Worksheets('DB_Resultados')
$ult = $db.Cells($db.Rows.Count, 1).End(-4162).Row
"linhas no banco: $($ult - 3)"

$vistos = @{}
$ordem = New-Object System.Collections.ArrayList
if ($ult -ge $R0) {
    $dados = $db.Range($db.Cells($R0, 1), $db.Cells($ult, 7)).Value2
    $nL = $dados.GetLength(0)
    for ($i = 1; $i -le $nL; $i++) {
        $run = $dados.GetValue($i, 1)
        if ($run -eq $null -or -not ($run -is [double])) { continue }
        $r = [int]$run
        if ($vistos.ContainsKey($r)) { continue }
        $serial = $dados.GetValue($i, 2)
        $lote = [string]$dados.GetValue($i, 4)
        $core = if ($lote.Length -ge 9) { $lote.Substring(3, 6) } else { '' }
        $vistos[$r] = $true
        [void]$ordem.Add([pscustomobject]@{ RUN = $r; Serial = $serial; Lote = $core })
    }
}

# Escrita celula a celula, com .Value e valor ja tipado.
# Tentei os dois atalhos e nenhum vinga no PowerShell 5.1: .Value2 com escalar
# numerico escolhe a sobrecarga de String e estoura InvalidCast, e a atribuicao
# de um object[,] a um Range multi-celula nao vincula (nao lanca erro e nao
# grava nada). Sao 25 corridas x 4 campos -- o custo e irrelevante e o
# comportamento e previsivel.
$lista = @($ordem | Sort-Object RUN)
if ($lista.Count -gt 0) {
    $cor.Range($cor.Cells($R0, 5), $cor.Cells($R0 + $lista.Count - 1, 5)).NumberFormat = '@'
    for ($k = 0; $k -lt $lista.Count; $k++) {
        $lin = $R0 + $k
        $vRun = [int]$lista[$k].RUN
        $cor.Cells($lin, 1).Formula = [string]$vRun
        if ($lista[$k].Serial -ne $null) {
            # .Formula com STRING, nao .Value/.Value2 com numero ou DateTime.
            # Neste ambiente o binding COM de valor numerico e erratico: dentro
            # de um .ps1 lanca InvalidCast, no mesmo codigo em bloco inline
            # grava normalmente. String sempre funcionou. O Excel interpreta o
            # texto como numero e o formato de data cuida da exibicao.
            # InvariantCulture e obrigatorio: em pt-BR a virgula decimal
            # transformaria o serial noutro numero.
            $vData = ([double]$lista[$k].Serial).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $cor.Cells($lin, 2).Formula = $vData
        }
        $cor.Cells($lin, 5).Value = [string]$lista[$k].Lote
        $cor.Cells($lin, 9).Value = 'migracao'

        # Confere a gravacao da linha. Este ambiente ja engoliu escrita COM tres
        # vezes sem lancar erro; sem a leitura de volta, a migracao mente.
        if ([string]$cor.Cells($lin, 1).Value2 -eq '') { throw "RUN nao gravado na linha $lin" }
        if ($lista[$k].Serial -ne $null -and [string]$cor.Cells($lin, 2).Value2 -eq '') {
            throw "Data nao gravada na linha $lin (RUN $vRun)"
        }
    }
    $cor.Range($cor.Cells($R0, 2), $cor.Cells($R0 + $lista.Count - 1, 2)).NumberFormat = 'dd/mm/yyyy'
}

$maior = 0
if ($ordem.Count -gt 0) { $maior = [int]($ordem | Measure-Object RUN -Maximum).Maximum }
$prox = [int]($maior + 1)
# .Value com valor ja tipado, nao .Value2 com expressao: o binding COM do
# PowerShell 5.1 escolhe a sobrecarga de String e estoura InvalidCast.
$cor.Range('D1').Formula = [string]$prox

$wb.Names.Add('proxRUN', '=Corridas!$D$1') | Out-Null

$cor.Columns('A:I').AutoFit() | Out-Null
$cor.Visible = 0        # oculta, mas nao very hidden: e registro auditavel

# Conferencia de que a gravacao ocorreu de fato. Sem isto, uma falha de binding
# COM deixa a aba vazia e o script segue relatando sucesso -- foi o que
# aconteceu na primeira versao, e so apareceu no teste de comportamento.
$ultGravada = $cor.Cells($cor.Rows.Count, 1).End(-4162).Row
$gravadas = $ultGravada - $R0 + 1
if ($gravadas -ne $ordem.Count) {
    throw "Migracao inconsistente: esperadas $($ordem.Count) corridas, gravadas $gravadas."
}
if ([int]$cor.Range('D1').Value2 -ne $prox) {
    throw "Contador proxRUN nao gravado: esperado $prox, encontrado '$($cor.Range('D1').Value2)'."
}

"corridas migradas : $($ordem.Count)  (conferido: $gravadas linhas)"
"maior RUN         : $maior"
"proxRUN           : $prox"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
