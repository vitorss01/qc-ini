# aplicar_vba.ps1 - importa modulos VBA versionados para dentro do .xlsm de build
#
# Materializa a regra "VBA e fonte, .xlsm e artefato": o modulo e editado como
# arquivo de texto, versionado no Git, e daqui entra na pasta de trabalho.
# Nunca o contrario.
#
# Requer AccessVBOM = 1 (Central de Confiabilidade > Configuracoes de Macro >
# Confiar no acesso ao modelo de objeto do projeto do VBA).
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\aplicar_vba.ps1 -Workbook <build.xlsm> -Modulos <arq1.bas>,<arq2.bas>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string[]]$Modulos
)

# SEM ISTO O SCRIPT MENTE. Erro de COM em PowerShell e nao-terminante por
# padrao: o Import falha, a mensagem vai para o stderr, o laco continua, e a
# ultima linha imprime "VBA aplicado em: ..." como se tivesse dado certo. Foi
# exatamente assim que o motor corrigido deixou de entrar no .xlsm em duas
# maquinas diferentes, sem ninguem perceber -- o build reportava sucesso e o
# arquivo continuava com o mEstatistica de producao.
$ErrorActionPreference = 'Stop'

# Excel resolve caminho relativo contra o DIRETORIO PADRAO DELE, nao contra o
# diretorio do PowerShell. Um caminho relativo aqui abre outro arquivo ou nao
# abre nenhum.
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$Modulos = $Modulos | ForEach-Object { (Resolve-Path -LiteralPath $_).ProviderPath }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)
$proj = $wb.VBProject

foreach ($m in $Modulos) {
    $nome = [System.IO.Path]::GetFileNameWithoutExtension($m)

    $alvo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $nome) { $alvo = $c; break } }

    if ($alvo -ne $null) {
        if ($alvo.Type -eq 100) {
            # Documento (Planilha/EstaPastaDeTrabalho) nao pode ser removido:
            # substituir o codigo linha a linha.
            #
            # ATENCAO: AddFromFile insere o arquivo INTEIRO como codigo. Um .cls
            # exportado comeca com VERSION / BEGIN / END / Attribute VB_*, que
            # sao metadados do formato, nao VBA. Inseridos no corpo viram erro
            # de sintaxe e o PROJETO INTEIRO deixa de compilar -- inclusive
            # procedimentos de outros modulos, que passam a falhar no
            # Application.Run com 0x800A9C68. Por isso o cabecalho e removido
            # antes, e so o codigo entra.
            $bruto = [System.IO.File]::ReadAllLines($m, [System.Text.Encoding]::Default)
            $corpo = New-Object System.Collections.ArrayList
            $emCabecalho = $true
            foreach ($linha in $bruto) {
                if ($emCabecalho) {
                    if ($linha -match '^\s*(VERSION|BEGIN|END|MultiUse|Attribute)\b') { continue }
                    if ($linha.Trim() -eq '') { continue }
                    $emCabecalho = $false
                }
                [void]$corpo.Add($linha)
            }
            $tmp = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllLines($tmp, $corpo.ToArray(), [System.Text.Encoding]::Default)

            $alvo.CodeModule.DeleteLines(1, $alvo.CodeModule.CountOfLines)
            $alvo.CodeModule.AddFromFile($tmp)
            Remove-Item $tmp -Force
            "  substituido (documento): $nome  ($($corpo.Count) linhas de codigo)"
            continue
        }
        $proj.VBComponents.Remove($alvo)
        "  removido: $nome"
    }
    $proj.VBComponents.Import($m) | Out-Null

    # Conferir, nao confiar. Import pode falhar sem excecao util, e um modulo
    # ausente so aparece muito depois, como erro numa rotina sem relacao.
    $novo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $nome) { $novo = $c; break } }
    if ($novo -eq $null) { throw "Import de $nome nao criou o componente" }

    $linhasArquivo = ([System.IO.File]::ReadAllLines($m, [System.Text.Encoding]::Default)).Count
    $linhasModulo = $novo.CodeModule.CountOfLines
    if ($linhasModulo -lt 1) { throw "$nome entrou vazio" }
    "  importado: $nome  ($linhasModulo linhas no modulo, $linhasArquivo no arquivo)"
}

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"VBA aplicado em: $Workbook"
