# criar_abas_motor.ps1 - instala as abas que o motor estatistico exige
#
# Bioquimica e Imunologia nunca receberam a Fase 3: alem de nao terem
# mEstatistica, nao tem as duas abas de que o motor depende.
#
#   Cfg_Status        tabela de elegibilidade (ADR-006). O motor le daqui quais
#                     estados entram no calculo. Estado desconhecido NUNCA entra.
#                     Acrescentar um estado e digitar uma linha, sem tocar codigo.
#   Eventos_Westgard  historico auditavel de violacoes, reescrito a cada
#                     execucao de RegistrarEventosWestgard. E DERIVADA: pode ser
#                     apagada e reconstruida. Nao confundir com o Audit_Log, que
#                     e append-only e jamais pode ser reescrito.
#
# IDEMPOTENTE: se a aba ja existe, nao mexe. Rodar de novo nao destroi a
# configuracao de elegibilidade que o gestor tenha ajustado.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_abas_motor.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'

# Acentos por codigo, porque o arquivo e lido como ANSI e literal acentuado
# chegaria corrompido ao Excel.
#
# [string] em cada um NAO e redundante: em PowerShell, [char] + [char] e SOMA
# NUMERICA -- 0x00E7 + 0x00E3 da 458, nao "ca". O erro so aparece na
# concatenacao seguinte, como "nao e possivel converter System.Char em
# System.String", longe da causa.
$a_ = [string][char]0x00E1  # a agudo
$e_ = [string][char]0x00E9  # e agudo
$i_ = [string][char]0x00ED  # i agudo
$o_ = [string][char]0x00F3  # o agudo
$I_ = [string][char]0x00CD  # I agudo
$A_ = [string][char]0x00C3  # A TIL -- e o que 'NAO' precisa.
                            # Estava 0x00D5 (O til) e gravava 'NOO'.
$ac = [string][char]0x00E7 + [string][char]0x00E3 + 'o'      # "cao"

# CADA ELEMENTO ENTRE PARENTESES, e nao por estilo.
#
# Em PowerShell a VIRGULA liga MAIS FORTE que o '+'. Sem parenteses,
#   @('Exclu' + $i_ + 'do', 'N' + $A_ + 'O', 'desc')
# e lido como  'Exclu' + $i_ + ('do','N') + ...  -- soma de string com ARRAY --
# e o resultado colapsa: a linha vira UM elemento com tudo junto, ou se parte em
# cinco. Foi assim que a Cfg_Status da Bioquimica saiu com Status, Elegivel e
# Descricao fundidos numa celula so, e o combo de tipo de nao conformidade
# passou a mostrar "Excluido NAO Excluido logicamente pelo analista" ao usuario.
$STATUS = @(
    @( 'Ativo', 'SIM', 'Resultado valido, elegivel para media/DP/CV/Bias/Sigma/Westgard' ),
    @( ('Exclu' + $i_ + 'do'), ('N' + $A_ + 'O'), 'Excluido logicamente pelo analista' ),
    @( ('Inv' + $a_ + 'lido'), ('N' + $A_ + 'O'), 'Resultado invalido (erro analitico comprovado)' ),
    @( 'Rejeitado-Operacional', ('N' + $A_ + 'O'), 'Falha operacional (pipetagem, amostra, operador)' ),
    @( ('Calibra' + $ac), ('N' + $A_ + 'O'), 'Corrida associada a calibracao do equipamento' ),
    @( ('Manuten' + $ac), ('N' + $A_ + 'O'), 'Corrida associada a manutencao preventiva/corretiva' ),
    @( 'Troca de Lote', ('N' + $A_ + 'O'), 'Corrida de transicao entre lotes de controle' ),
    @( 'Treinamento', ('N' + $A_ + 'O'), 'Execucao para treinamento - nao compoe desempenho' )
)

# Conferir a forma do array ANTES de escrever. Se a precedencia voltar a
# quebrar, o build para aqui e nao entrega planilha corrompida ao laboratorio.
for ($k = 0; $k -lt $STATUS.Count; $k++) {
    if ($STATUS[$k].Count -ne 3) {
        throw "STATUS[$k] tem $($STATUS[$k].Count) elementos, esperado 3. Precedencia de virgula quebrou de novo: parentetize cada elemento."
    }
}

$CAB_EVENTOS = @( 'Data', 'RUN', 'Analito', ('N' + $i_ + 'vel'), 'Regra',
    ('Classifica' + $ac), 'Resultado', 'Z-Score' )
if ($CAB_EVENTOS.Count -ne 8) { throw "CAB_EVENTOS tem $($CAB_EVENTOS.Count) itens, esperado 8" }

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep -Seconds 5 } catch { }
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
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o segura): $Workbook"
}

$existentes = @($wb.Worksheets | ForEach-Object { $_.Name })
$criadas = @()

# ---------------- Cfg_Status ----------------
if ($existentes -contains 'Cfg_Status') {
    "Cfg_Status: ja existe, preservada"
}
else {
    $cfg = $wb.Worksheets.Add()
    $cfg.Name = 'Cfg_Status'
    $cfg.Range('B2').Formula = 'ELEGIBILIDADE ESTAT' + $I_ + 'STICA (CLSI EP05 / C24)'
    $cfg.Range('B2').Font.Bold = $true
    $cfg.Range('B3').Formula = 'Status'
    $cfg.Range('C3').Formula = 'Eleg' + $i_ + 'vel'
    $cfg.Range('D3').Formula = 'Descri' + $ac
    $cfg.Range('B3:D3').Font.Bold = $true
    for ($k = 0; $k -lt $STATUS.Count; $k++) {
        $lin = 4 + $k
        # .Value2 e nao .Formula: Formula faz o Excel INTERPRETAR o texto.
        $cfg.Cells($lin, 2).Value2 = $STATUS[$k][0]
        $cfg.Cells($lin, 3).Value2 = $STATUS[$k][1]
        $cfg.Cells($lin, 4).Value2 = $STATUS[$k][2]
    }
    $cfg.Cells(4 + $STATUS.Count, 2).Formula = 'Acrescente novos estados nas linhas abaixo. O Motor Estat' + $i_ + 'stico l' + $e_ + ' esta tabela - nenhuma rotina precisa ser alterada. Estado desconhecido nunca entra no c' + $a_ + 'lculo.'
    $cfg.Columns('B:D').AutoFit() | Out-Null
    $cfg.Visible = 0
    $criadas += 'Cfg_Status'
    "Cfg_Status: criada com $($STATUS.Count) estados (1 elegivel, $($STATUS.Count - 1) nao elegiveis)"
}

# ---------------- Eventos_Westgard ----------------
if ($existentes -contains 'Eventos_Westgard') {
    "Eventos_Westgard: ja existe, preservada"
}
else {
    $ev = $wb.Worksheets.Add()
    $ev.Name = 'Eventos_Westgard'
    $ev.Range('A1').Formula = 'EVENTOS DE WESTGARD - hist' + $o_ + 'rico audit' + $a_ + 'vel (gerado pelo Motor Estat' + $i_ + 'stico)'
    $ev.Range('A1').Font.Bold = $true
    $ev.Range('I2').Formula = 'Total de eventos:'
    for ($k = 0; $k -lt $CAB_EVENTOS.Count; $k++) {
        $ev.Cells(3, $k + 1).Formula = $CAB_EVENTOS[$k]
    }
    $ev.Range('A3:H3').Font.Bold = $true
    $ev.Columns('A:I').AutoFit() | Out-Null
    $ev.Visible = 0
    $criadas += 'Eventos_Westgard'
    "Eventos_Westgard: criada (cabecalho em A3:H3, dados a partir de A4)"
}

# ---------------- conferencia ----------------
$agora = @($wb.Worksheets | ForEach-Object { $_.Name })
$erros = @()
foreach ($n in @('Cfg_Status', 'Eventos_Westgard')) {
    if ($agora -notcontains $n) { $erros += "aba $n ausente apos a criacao" }
}
$cfgW = $wb.Worksheets('Cfg_Status')
$nElig = 0
for ($lin = 4; $lin -le 30; $lin++) {
    if ("$($cfgW.Cells($lin, 3).Value2)".Trim().ToUpper() -like 'SIM*') { $nElig++ }
}
if ($nElig -lt 1) { $erros += 'Cfg_Status sem nenhum estado elegivel - o motor nao calcularia nada' }

if ($erros.Count -gt 0) {
    $erros | ForEach-Object { "  FALHA: $_" }
    throw "Instalacao das abas do motor incompleta: $($erros.Count) verificacao(oes) falharam."
}
"conferencia: ok - $nElig estado(s) elegivel(is); abas criadas nesta execucao: $(if ($criadas.Count) { $criadas -join ', ' } else { 'nenhuma' })"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
