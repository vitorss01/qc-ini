# snapshot_projeto.ps1 - ponto de restauracao completo de uma pasta .xlsm
#
# Extrai da pasta de trabalho tudo que e necessario para reconstrui-la:
#   vba/            todos os modulos (.bas), classes (.cls) e formularios (.frm/.frx)
#   nomes.csv       nomes definidos (escopo, referencia, visibilidade)
#   abas.csv        abas: nome, CodeName, visibilidade, protecao, UsedRange
#   objetos.csv     graficos, formas, tabelas (ListObjects), validacoes
#
# As formulas ficam de fora deste script de proposito: sao geradas por
# snapshot_formulas.ps1, que produz o inventario celula a celula usado no diff.
#
# Requer: "Confiar no acesso ao modelo de objeto do projeto do VBA" habilitado
#         (HKCU:\Software\Microsoft\Office\16.0\Excel\Security -> AccessVBOM = 1).
#         Sem isso, a exportacao do VBA falha e o script avisa.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\snapshot_projeto.ps1 -Workbook <caminho.xlsm> -OutDir <pasta_destino>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$OutDir
)

$vbaDir = Join-Path $OutDir 'vba'
New-Item -ItemType Directory -Force -Path $vbaDir | Out-Null

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
$xl.AutomationSecurity = 3          # nao executa macro ao abrir

$wb = $xl.Workbooks.Open($Workbook, 0, $true)

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




# ---------- VBA ----------
$nVba = 0
try {
    $proj = $wb.VBProject
    foreach ($comp in $proj.VBComponents) {
        $ext = switch ($comp.Type) {
            1 { '.bas' }    # vbext_ct_StdModule
            2 { '.cls' }    # vbext_ct_ClassModule
            3 { '.frm' }    # vbext_ct_MSForm
            100 { '.cls' }  # vbext_ct_Document (Planilha/EstaPastaDeTrabalho)
            default { '.txt' }
        }
        $destino = Join-Path $vbaDir ($comp.Name + $ext)
        $comp.Export($destino)
        $nVba++
        "  VBA  {0,-28} {1}" -f $comp.Name, $ext | Write-Output
    }
    "VBA exportado: $nVba componentes"
}
catch {
    "FALHA ao exportar VBA: $($_.Exception.Message)"
    "  Habilite: Arquivo > Opcoes > Central de Confiabilidade > Configuracoes de Macro"
    "            > Confiar no acesso ao modelo de objeto do projeto do VBA"
}

# ---------- nomes definidos ----------
$nomes = New-Object System.Collections.ArrayList
foreach ($nm in $wb.Names) {
    $ref = ''
    try { $ref = $nm.RefersTo } catch { $ref = '<erro>' }
    [void]$nomes.Add([pscustomobject]@{
            Nome    = $nm.Name
            RefersTo = $ref
            Visivel = $nm.Visible
        })
}
$nomes | Sort-Object Nome | Export-Csv (Join-Path $OutDir 'nomes.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
"Nomes definidos: $($nomes.Count)"

# ---------- abas ----------
$abas = New-Object System.Collections.ArrayList
$objs = New-Object System.Collections.ArrayList
foreach ($ws in $wb.Worksheets) {
    [void]$abas.Add([pscustomobject]@{
            Nome         = $ws.Name
            CodeName     = $ws.CodeName
            Indice       = $ws.Index
            Visivel      = $ws.Visible          # -1 visivel, 0 oculta, 2 muito oculta
            Protegida    = $ws.ProtectContents
            UsedRange    = $ws.UsedRange.Address()
            Celulas      = $ws.UsedRange.Cells.Count
        })
    foreach ($lo in $ws.ListObjects) {
        [void]$objs.Add([pscustomobject]@{ Aba = $ws.Name; Tipo = 'Tabela'; Nome = $lo.Name; Detalhe = $lo.Range.Address() })
    }
    foreach ($co in $ws.ChartObjects()) {
        [void]$objs.Add([pscustomobject]@{ Aba = $ws.Name; Tipo = 'Grafico'; Nome = $co.Name; Detalhe = $co.Chart.ChartType })
    }
    foreach ($sh in $ws.Shapes) {
        [void]$objs.Add([pscustomobject]@{ Aba = $ws.Name; Tipo = 'Forma'; Nome = $sh.Name; Detalhe = $sh.Type })
    }
}
$abas | Export-Csv (Join-Path $OutDir 'abas.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
$objs | Export-Csv (Join-Path $OutDir 'objetos.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
"Abas: $($abas.Count)   Objetos: $($objs.Count)"

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Snapshot em: $OutDir"
