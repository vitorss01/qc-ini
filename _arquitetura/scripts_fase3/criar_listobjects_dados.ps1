# criar_listobjects_dados.ps1 - ADR-051: as duas abas viram tabela nomeada
#
# POR QUE
#
# O ETL da Fase 2 le o ListObject PELO NOME -- e assim que tblBI_Fato entra no
# Power Query, e e o que permite o banco crescer sem tocar na consulta.
# Eventos_Westgard e EQA_Base tem dado estruturado e nenhum nome: seriam lidas
# por faixa de celulas, que quebra na primeira linha a mais.
#
#   Eventos_Westgard   nenhum dos dois produtos tem ListObject
#   EQA_Base           so a Bioquimica tem tblEQA_Base
#
# O QUE ESTE SCRIPT NAO FAZ
#
# Nao toca no motor, nos eventos, nas regras de Westgard, no Bias, no CAP nem
# no Controllab. So expoe o que ja existe sob um nome estavel.
#
# OS DOIS CABECALHOS COM ACENTO ERRADO
#
# Eventos_Westgard nasceu com "NÍveis" e "Evidéncia" (certo: "Niveis" e
# "Evidencia"). Num cabecalho de tabela isso deixa de ser cosmetico: o nome da
# coluna do ListObject vira o nome do campo no Power Query e no modelo. Sao
# corrigidos aqui porque e exatamente a interface que este script padroniza --
# nao e melhoria oportunista, e o contrato sendo definido.
#
# IDEMPOTENTE: tabela que ja existe com o range certo nao e recriada.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso: .\criar_listobjects_dados.ps1 -Workbook <arquivo.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$SENHA = 'qcini2025'

# Eventos_Westgard: cabecalho na linha 3, colunas A..N (ADR-045)
$EV_ABA = 'Eventos_Westgard'; $EV_TAB = 'tblEventos_Westgard'
$EV_CAB = 3; $EV_NCOL = 14
# EQA_Base: cabecalho na linha 1, colunas A..U (as 21 do consolidado)
$EQ_ABA = 'EQA_Base'; $EQ_TAB = 'tblEQA_Base'
$EQ_CAB = 1; $EQ_NCOL = 21

$xlSrcRange = 1
$xlYes = 1
$xlUp = -4162

function Novo-Excel {
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { Start-Sleep -Seconds (1 + $t) }
    }
    throw 'nao consegui criar a instancia do Excel'
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvou = $false
try {
    $estrutura = [bool]$wb.ProtectStructure
    if ($estrutura) { $wb.Unprotect($SENHA) }

    foreach ($cfg in @(
        @{ Aba = $EV_ABA; Tab = $EV_TAB; Cab = $EV_CAB; NCol = $EV_NCOL },
        @{ Aba = $EQ_ABA; Tab = $EQ_TAB; Cab = $EQ_CAB; NCol = $EQ_NCOL })) {

        $ws = $null
        foreach ($w in $wb.Worksheets) { if ($w.Name -eq $cfg.Aba) { $ws = $w; break } }
        if ($ws -eq $null) { "  $($cfg.Aba): aba ausente neste produto, ignorada"; continue }

        $vis = $ws.Visible
        $ws.Visible = -1                       # visivel para poder mexer
        $prot = [bool]$ws.ProtectContents
        if ($prot) { try { $ws.Unprotect($SENHA) } catch { $ws.Unprotect() } }

        # Corrige os dois cabecalhos de acento errado ANTES de criar a tabela:
        # depois de criada, o nome da coluna passa a ser o contrato.
        if ($cfg.Aba -eq $EV_ABA) {
            $certo = @{
                4  = 'N' + [char]0x00ED + 'veis'
                12 = 'Evid' + [char]0x00EA + 'ncia'
            }
            foreach ($c in $certo.Keys) {
                $atual = [string]$ws.Cells($cfg.Cab, $c).Value2
                if ($atual -cne $certo[$c]) {
                    $ws.Cells($cfg.Cab, $c).Value2 = $certo[$c]
                    "  $($cfg.Aba): cabecalho da coluna $c corrigido para '$($certo[$c])'"
                }
            }
        }

        # Ultima linha COM DADO na coluna 1, nunca UsedRange: a aba guarda area
        # limpa muito abaixo do fim real, e a tabela herdaria milhares de linhas
        # vazias -- as "linhas fantasma" que o teste procura.
        $ult = $ws.Cells($ws.Rows.Count, 1).End($xlUp).Row
        if ($ult -lt ($cfg.Cab + 1)) { $ult = $cfg.Cab + 1 }   # tabela exige >= 1 linha

        $L = [char](64 + $cfg.NCol)
        $endereco = "A$($cfg.Cab):$L$ult"

        $lo = $null
        foreach ($t in $ws.ListObjects) { if ($t.Name -eq $cfg.Tab) { $lo = $t; break } }

        if ($lo -ne $null) {
            if ($lo.Range.Address -eq $ws.Range($endereco).Address) {
                "  $($cfg.Tab): ja existe em $endereco"
            }
            else {
                $lo.Resize($ws.Range($endereco))
                "  $($cfg.Tab): redimensionada para $endereco"
            }
        }
        else {
            # Outra tabela cobrindo a mesma area impediria a criacao.
            foreach ($t in @($ws.ListObjects)) {
                if ($xl.Intersect($t.Range, $ws.Range($endereco)) -ne $null) {
                    "  $($cfg.Aba): removendo tabela sobreposta '$($t.Name)'"
                    $t.Unlist()
                }
            }
            $novo = $ws.ListObjects.Add($xlSrcRange, $ws.Range($endereco), $null, $xlYes)
            $novo.Name = $cfg.Tab
            $novo.TableStyle = ''      # sem estilo: a aba e tecnica, nao vitrine
            "  $($cfg.Tab): criada em $endereco ($($ult - $cfg.Cab) linha(s) de dado)"
        }

        if ($prot) { try { $ws.Protect($SENHA) } catch { } }
        $ws.Visible = $vis
    }

    if ($estrutura) { try { $wb.Protect($SENHA, $true, $false) } catch { } }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { $wb.Close($salvou) } catch { }
    try { $xl.Quit() } catch { }
}
