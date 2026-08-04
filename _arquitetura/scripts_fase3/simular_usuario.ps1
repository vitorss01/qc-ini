# simular_usuario.ps1 - AGENTE USUARIO / TESTADOR DE INTERFACE
#
# Simula o uso diario de um analista sobre os UserForms do produto indicado.
# Nao le codigo: CARREGA os formularios de verdade, preenche controles e
# observa a reacao. Erro de execucao no Initialize, combo vazio, evento que nao
# dispara ou validacao que nao valida aparecem aqui e nao no laboratorio.
#
# POR QUE NAO CLICAR NOS BOTOES DE SALVAR. Os manipuladores de clique abrem
# MsgBox de confirmacao. Dialogo modal em processo sem interface prende o
# Excel indefinidamente -- foi assim que a caixa de senha travou um build por
# horas nesta sessao. O que os botoes CHAMAM (MarcarNaoConforme,
# UpsertResultados, ExcluirLogico) ja e coberto pela suite; aqui se testa o que
# so existe na interface: carga, povoamento, eventos e feedback ao usuario.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\simular_usuario.ps1 -Produto Bioquimica
#   .\simular_usuario.ps1 -Produto Bioquimica -Workbook <caminho.xlsm>

param(
    [string]$Produto = 'Bioquimica',
    [string]$Workbook = ''
)

$ErrorActionPreference = 'Stop'

$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$raiz = Split-Path -Parent $arq

if ($Workbook -eq '') {
    $Workbook = Join-Path $raiz "_entregas\QC_${Produto}_hardening1.xlsm"
}
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

# Trabalha sobre COPIA: a simulacao preenche campos e dispara eventos; o
# entregavel nao pode sair alterado por um teste.
$bd = Join-Path $env:USERPROFILE "QCINI_sim_$Produto"
New-Item -ItemType Directory -Force -Path $bd | Out-Null
$copia = Join-Path $bd "QC_$Produto.xlsm"
Copy-Item $Workbook $copia -Force

function Novo-Excel {
    $ultimo = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

$itens = New-Object System.Collections.ArrayList
function Anotar {
    param([string]$Form, [string]$Caso, [bool]$Ok, [string]$Evidencia)
    [void]$itens.Add([pscustomobject]@{ Form = $Form; Caso = $Caso; Ok = $Ok; Evidencia = $Evidencia })
    $m = if ($Ok) { 'OK   ' } else { 'FALHA' }
    "  [$m] $Form :: $Caso"
    if ($Evidencia) { "          $Evidencia" }
}

""
"================================================================"
"  SIMULACAO DE USO - $Produto"
"  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
"================================================================"
"  copia de trabalho: $copia"

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $true          # o usuario real dispara eventos
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($copia)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $copia" }

try {
    # ---- login, como o analista faz ao abrir ----
    # A senha nao pode ser chutada: o hash e de mao unica e cada produto pode ter
    # a sua. O harness tenta as conhecidas e RELATA qual valeu -- se nenhuma
    # valer, o achado e que ninguem consegue entrar no sistema.
    $senhas = @('HEMATO123', 'BIOQ123', 'IMUNO123', 'QCINI123')
    $usr = ''
    $usada = ''
    foreach ($sn in $senhas) {
        $wb.Names('loginUser').RefersToRange.Value2 = 'QCINI'
        $wb.Names('loginPass').RefersToRange.Value2 = $sn
        try { $xl.Run('DoLogin') } catch { }
        $usr = [string]$wb.Names('currentUser').RefersToRange.Value2
        if ($usr -ne '') { $usada = $sn; break }
    }
    Anotar 'Login' 'autenticacao inicial' ($usr -ne '') "currentUser = '$usr' (senha: $usada)"
    if ($usada -ne '' -and $usada -notlike "*$($Produto.Substring(0,4).ToUpper())*") {
        Anotar 'Login' 'produto usa credencial propria' $false `
            "entrou com '$usada' -- o produto $Produto esta sendo entregue com a senha de outro setor"
    }

    # O harness vive num modulo temporario: manipulador Private de formulario
    # nao e alcancavel de fora, mas o objeto do formulario e.
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'mSimUsuario'
    $codigo = [System.IO.File]::ReadAllText((Join-Path $arq 'src_hardening1\SimUsuario.txt'),
        (New-Object System.Text.UTF8Encoding($false)))
    $mod.CodeModule.AddFromString($codigo)
    Start-Sleep -Milliseconds 400

    $casos = @(
        @('frmCorrida', 'S_Corrida_Carga', 'carrega e popula sem erro'),
        @('frmCorrida', 'S_Corrida_Niveis', 'combo de nivel com os niveis do produto'),
        @('frmCorrida', 'S_Corrida_LoteAtivo', 'lote ativo vem pre-selecionado'),
        @('frmCorrida', 'S_Corrida_CamposAnalito', 'um campo por analito cadastrado'),
        @('frmCorrida', 'S_Corrida_RunPrevisto', 'RUN previsto aparece ao escolher data e lote'),
        @('frmCorrida', 'S_Corrida_DataInvalida', 'data invalida limpa o RUN em vez de travar'),
        @('frmExcluir', 'S_Excluir_Carga', 'carrega e popula sem erro'),
        @('frmExcluir', 'S_Excluir_Runs', 'combo de RUN traz as corridas do lote'),
        @('frmExcluir', 'S_Excluir_Analitos', 'lista de analitos do nivel selecionado'),
        @('frmExcluir', 'S_Excluir_Todos', 'botao Todos seleciona a lista inteira'),
        @('frmMassa', 'S_Massa_Carga', 'carrega e popula sem erro'),
        @('frmMassa', 'S_Massa_Validacao', 'colagem com nivel invalido e recusada'),
        @('frmNC', 'S_NC_Carga', 'carrega e popula sem erro'),
        @('frmNC', 'S_NC_Tipos', 'tipos vem do Cfg_Status, nenhum em codigo'),
        @('frmNC', 'S_NC_RunPreenche', 'informar o RUN preenche data, lote e resultados'),
        @('frmNC', 'S_NC_RunInexistente', 'RUN inexistente avisa em vez de travar'),
        @('frmNC', 'S_NC_ContagemParecer', 'contador de palavras reage ao que se digita'),
        @('frmNC', 'S_NC_Todos', 'botao Todos seleciona a lista inteira'),
        @('frmExcluirNC', 'S_ExNC_Carga', 'carrega e popula sem erro'),
        @('frmExcluirNC', 'S_ExNC_ContagemParecer', 'contador de palavras reage ao que se digita'),
        @('frmAssinar', 'S_Assinar_Carga', 'carrega sem erro e mascara a senha'),
        @('frmDev', 'S_Dev_Carga', 'carrega sem erro e mascara a senha')
    )

    ""
    "-- FORMULARIOS ------------------------------------------------"
    foreach ($c in $casos) {
        $ret = ''
        $falhou = $false
        try { $ret = [string]$xl.Run($c[1]) }
        catch { $falhou = $true; $ret = $_.Exception.Message }
        $ok = (-not $falhou) -and ($ret -like 'OK*')
        Anotar $c[0] $c[2] $ok $ret
    }

    $wb.VBProject.VBComponents.Remove($mod)
}
finally {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    Get-Process EXCEL -EA SilentlyContinue | ForEach-Object { try { $_.Kill() } catch { } }
}

$falhas = @($itens | Where-Object { -not $_.Ok })
""
"================================================================"
"  $($itens.Count - $falhas.Count) de $($itens.Count) casos passaram"
if ($falhas.Count -gt 0) {
    "  FALHAS:"
    foreach ($f in $falhas) { "    - $($f.Form) :: $($f.Caso)"; "        $($f.Evidencia)" }
}
"================================================================"

if ($falhas.Count -gt 0) { exit 1 }
exit 0
