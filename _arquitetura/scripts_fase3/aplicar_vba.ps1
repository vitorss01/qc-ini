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
            $alvo.CodeModule.DeleteLines(1, $alvo.CodeModule.CountOfLines)
            $alvo.CodeModule.AddFromFile($m)
            "  substituido (documento): $nome"
            continue
        }
        $proj.VBComponents.Remove($alvo)
        "  removido: $nome"
    }
    $proj.VBComponents.Import($m) | Out-Null
    "  importado: $nome"
}

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"VBA aplicado em: $Workbook"
