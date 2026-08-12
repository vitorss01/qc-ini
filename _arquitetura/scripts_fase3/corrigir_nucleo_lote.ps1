# corrigir_nucleo_lote.ps1 - o nucleo do lote deixa de assumir 6 digitos
#
# O DEFEITO
#
# CodigoLote(core, nivel) monta  "QC-" & core & Format(nivel,"00").
# O inverso, portanto, e  Mid(codigo, 4, Len(codigo) - 5): tira os 3 do prefixo
# "QC-" e os 2 do nivel, seja qual for o tamanho do nucleo.
#
# O sistema inteiro usava Mid(codigo, 4, 6) -- SEIS fixo -- em 13 lugares
# diferentes. Com nucleo de 6 digitos os dois dao no mesmo, e por isso passou.
# Com o lote real do laboratorio, "8974" (4 digitos), o codigo fica "QC-897401"
# e o Mid(4,6) devolve "897401": engole os dois digitos do NIVEL.
#
# A consequencia era muda e total:
#   Calc!B3  filtra por  (""&rLote)=(""&loteAtivo)
#   rLote    = DB_Resultados!BA = MID($D4,4,6) = "897401"
#   loteAtivo= Configuracao!C20                = "8974"
# Nunca casa. O Calc fica vazio, o Painel fica sem pontos e a aba Resultados nao
# carrega -- com os 555 registros intactos no banco, gravados corretamente. O
# dado entrava e nao saia.
#
# POR QUE 13 LUGARES, E NAO UM
#
# mEntrada.NucleoLote JA EXISTE e faz exatamente esta conta. Nenhum chamador a
# usava: cada modulo reescreveu o Mid inline. Conhecimento duplicado em 13
# copias e a mesma falha que o ADR-009 e o ADR-019 tratam -- corrigir uma nao
# corrige nada, e foi por isso que o defeito sobreviveu a tantas revisoes.
# Este script faz as 13 chamarem a funcao.
#
# NAO TOCA EM DADO. Nenhum registro do banco muda: o que muda e como o codigo
# do lote e LIDO. Os 555 resultados ja gravados passam a ser encontrados.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\corrigir_nucleo_lote.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    # ---------------------------------------------------------------- VBA ----
    # Troca o Mid inline pela chamada a NucleoLote, e corrige a propria
    # NucleoLote. A ordem importa: a definicao e corrigida DEPOIS das trocas,
    # senao a linha da definicao viraria uma chamada recursiva a si mesma.
    $trocasVba = 0
    $detalhesVba = @()
    $reMid = [regex]'Mid\$?\(\s*(.+?)\s*,\s*4\s*,\s*6\s*\)'

    foreach ($comp in $wb.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -lt 1) { continue }
        $texto = $cm.Lines(1, $cm.CountOfLines)
        if ($texto -notmatch 'Mid\$?\(\s*.+?\s*,\s*4\s*,\s*6\s*\)') { continue }

        $linhas = $texto -split "`r?`n"
        $mudou = $false
        for ($i = 0; $i -lt $linhas.Count; $i++) {
            $l = $linhas[$i]
            if ($l -match '^\s*''') { continue }                       # comentario
            if ($l -match 'Function\s+NucleoLote') { continue }        # a declaracao
            # A LINHA DE ATRIBUICAO DA PROPRIA FUNCAO.
            #
            # "NucleoLote = Mid$(Trim$(codigo), 4, 6)" casa com o padrao geral e
            # viraria "NucleoLote = NucleoLote(Trim$(codigo))" -- recursao
            # infinita, que em VBA e estouro de pilha e derruba o Excel. Esta
            # linha e corrigida no passo seguinte, que troca o 6 por Len-5.
            # (Pego pela execucao com -Simular, antes de gravar.)
            if ($l -match '^\s*NucleoLote\s*=') { continue }
            if ($l -notmatch 'Mid\$?\(') { continue }
            # so troca quando o Mid extrai de um CODIGO DE LOTE
            if ($l -notmatch 'COL_LOTE|lote|Lote|codigo') { continue }
            $novo = $reMid.Replace($l, 'NucleoLote($1)')
            if ($novo -ne $l) {
                $linhas[$i] = $novo
                $detalhesVba += "$($comp.Name):$($i + 1)"
                $trocasVba++
                $mudou = $true
            }
        }
        if ($mudou -and -not $Simular) {
            $cm.DeleteLines(1, $cm.CountOfLines)
            $cm.AddFromString(($linhas -join "`r`n"))
        }
    }

    # a definicao de NucleoLote
    $defOk = $false
    foreach ($comp in $wb.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -lt 1) { continue }
        $texto = $cm.Lines(1, $cm.CountOfLines)
        if ($texto -notmatch 'Function\s+NucleoLote') { continue }
        $linhas = $texto -split "`r?`n"
        for ($i = 0; $i -lt $linhas.Count; $i++) {
            if ($linhas[$i] -match '^\s*NucleoLote\s*=\s*Mid') {
                $linhas[$i] = '    ' + 'NucleoLote = Mid$(Trim$(codigo), 4, Len(Trim$(codigo)) - 5)'
                $defOk = $true
            }
        }
        if ($defOk -and -not $Simular) {
            $cm.DeleteLines(1, $cm.CountOfLines)
            $cm.AddFromString(($linhas -join "`r`n"))
        }
    }
    if (-not $defOk) { throw "definicao de NucleoLote nao encontrada -- nada foi alterado" }

    "VBA: $trocasVba chamada(s) inline trocadas por NucleoLote"
    $detalhesVba | ForEach-Object { "   $_" }
    "VBA: definicao de NucleoLote corrigida para Len(codigo) - 5"

    # ------------------------------------------------------------ formula ----
    # DB_Resultados!BA = LoteCore. Mesma conta, na camada de formula.
    $db = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'DB_Resultados') { $db = $w; break } }
    if ($db -eq $null) { throw 'aba DB_Resultados ausente' }
    if ($db.ProtectContents) { try { $db.Unprotect($SENHA) } catch { } }

    # ---- BA deixou de ser formula no ADR-025 ----
    #
    # Este bloco reescrevia a formula do nucleo do lote em BA4:BA15003. Depois do
    # ADR-025, BA/BB/BC sao VALOR mantido por mBanco.AtualizarFlagsBanco, e o
    # 15003 era o teto antigo de provisionamento (13,5 meses de dados).
    #
    # Rodar este script como estava DESFARIA o ADR-025: recriaria 15.000 formulas
    # e o COUNTIFS expansivo voltaria junto na proxima manutencao das colunas
    # vizinhas. Achado A3 da auditoria de 12/08/2026.
    #
    # A correcao do nucleo do lote continua necessaria -- e ela agora vive no
    # VBA, em NucleoLote, que este mesmo script ja corrige acima. Aqui basta
    # mandar o mBanco regravar os valores pela regra nova.
    $R0 = 4
    $temBanco = $false
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mBanco') { $temBanco = $true; break } }

    if ($temBanco) {
        $xl.Run('AtualizarFlagsBanco') | Out-Null
        $ultReal = $db.Cells($db.Rows.Count, 1).End(-4162).Row
        "coluna BA: regravada como VALOR por mBanco.AtualizarFlagsBanco (linhas 4..$ultReal)"
        "   BA4 agora = '$($db.Cells(4,53).Value2)'"
    }
    else {
        # Arquivo anterior ao ADR-025: mantem o comportamento de formula, mas o
        # limite vem da ULTIMA LINHA REAL, nao de um teto fixo.
        $ultReal = [Math]::Max($db.Cells($db.Rows.Count, 1).End(-4162).Row, $R0)
        $antes = $db.Cells($R0, 53).Formula
        $nova = '=IF($D{0}="","",MID($D{0},4,LEN($D{0})-5))' -f $R0
        if (-not $Simular) {
            $db.Range($db.Cells($R0, 53), $db.Cells($ultReal, 53)).Formula = $nova
        }
        "formula: BA$R0`:BA$ultReal  (sem mBanco no arquivo -- modo compatibilidade)"
        "   antes : $antes"
        "   agora : $nova"
    }

    if ($Simular) {
        "SIMULACAO -- nada foi gravado."
        $wb.Close($false); $xl.Quit(); exit 0
    }

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # --------------------------------------------------------- conferencia ---
    # CONFERE, NAO CONFIA: o unico criterio que importa e o Calc deixar de
    # estar vazio. Corrigir a formula e ver a mesma tela em branco seria
    # trocar um defeito por outro sem perceber.
    $erros = @()
    $la = "$($wb.Names.Item('loteAtivo').RefersToRange.Value2)".Trim()
    $ba = "$($db.Cells(4, 53).Value2)".Trim()
    if ($ba -ne $la) { $erros += "BA4='$ba' ainda nao casa com loteAtivo='$la'" }

    try {
        $n = $xl.Run('NucleoLote', 'QC-897401')
        if ("$n" -ne '8974') { $erros += "NucleoLote('QC-897401') devolveu '$n', esperado '8974'" }
    }
    catch { $erros += "NucleoLote nao respondeu: $($_.Exception.Message)" }

    $calc = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'Calc') { $calc = $w; break } }
    $comRun = 0
    for ($r = 3; $r -le 40; $r++) { if ("$($calc.Cells($r, 2).Value2)".Trim() -ne '') { $comRun++ } }
    if ($comRun -eq 0) { $erros += 'Calc continua sem nenhuma corrida na coluna B' }

    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Correcao rejeitada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }

    "conferencia: NucleoLote('QC-897401')='8974'; BA4='$ba' = loteAtivo; Calc com $comRun corrida(s)"

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
