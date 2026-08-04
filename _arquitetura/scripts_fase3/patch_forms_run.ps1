# patch_forms_run.ps1 - migra os formularios para a nova API do RUN
#
# Edita o codigo pelo CodeModule, linha a linha. NAO reimporta o .frm: um
# UserForm tem um .frx binario companheiro (layout, imagens), e remover e
# reimportar o componente arrisca perder ou dessincronizar esse par. Trocar
# linhas dentro do modulo nao toca no .frx.
#
# O que muda e por que:
#
#   frmCorrida.AtualizarRun    NovoRUN -> PreverRUN
#       E preview: roda a cada tecla digitada na data. Nao pode alocar numero.
#
#   frmCorrida.btnSalvar_Click NovoRUN -> ObterOuCriarRUN
#       E gravacao efetiva: aqui a corrida passa a existir.
#
#   frmMassa.Validar           NovoRUN -> PreverRUN
#       Validar roda tambem no botao Validar, sem gravar nada.
#
#   frmMassa.btnSalvar_Click   insere remapeamento antes de gravar
#       Corrige um defeito antigo: numa importacao com DUAS datas, nenhuma
#       esta no banco ainda, entao ambas recebiam o mesmo numero. O remap
#       resolve a identidade por (Data, nucleo do lote) uma vez por corrida,
#       no momento da gravacao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\patch_forms_run.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

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

$proj = $wb.VBProject

function Get-Comp($proj, $nome) {
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $nome) { return $c } }
    return $null
}

$alteracoes = 0

# ---------- frmCorrida ----------
$fc = Get-Comp $proj 'frmCorrida'
if ($fc -eq $null) { throw 'frmCorrida nao encontrado' }
$cm = $fc.CodeModule
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    $t = $cm.Lines($i, 1)
    if ($t -match 'lblRun\.Caption\s*=.*NovoRUN') {
        $cm.ReplaceLine($i, ($t -replace 'NovoRUN', 'PreverRUN'))
        "  frmCorrida L$i  preview  -> PreverRUN"
        $alteracoes++
    }
    elseif ($t -match '^\s*run\s*=\s*NovoRUN') {
        $cm.ReplaceLine($i, ($t -replace 'NovoRUN', 'ObterOuCriarRUN'))
        "  frmCorrida L$i  gravacao -> ObterOuCriarRUN"
        $alteracoes++
    }
}

# ---------- frmMassa ----------
$fm = Get-Comp $proj 'frmMassa'
if ($fm -eq $null) { throw 'frmMassa nao encontrado' }
$cm = $fm.CodeModule
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    $t = $cm.Lines($i, 1)
    if ($t -match '^\s*run\s*=\s*NovoRUN') {
        $cm.ReplaceLine($i, ($t -replace 'NovoRUN', 'PreverRUN'))
        "  frmMassa   L$i  validacao -> PreverRUN (provisorio)"
        $alteracoes++
        break
    }
}

# remapeamento definitivo antes de gravar
$remap = @(
    '',
    '    '' Identidade definitiva da corrida, uma vez por (Data + nucleo do lote).',
    '    '' Validar produz RUN apenas PROVISORIO: durante a validacao nenhuma das',
    '    '' datas novas existe no banco, entao duas datas diferentes receberiam o',
    '    '' mesmo numero. Aqui a corrida passa a existir de fato e o RUN definitivo',
    '    '' substitui o provisorio.',
    '    Dim mapaRun As Object, kRun As String, iR As Long',
    '    Set mapaRun = CreateObject("Scripting.Dictionary")',
    '    For iR = 1 To n',
    '        kRun = CStr(CLng(CDate(regs(iR, 2)))) & "|" & NucleoLote(CStr(regs(iR, 4)))',
    '        If Not mapaRun.Exists(kRun) Then',
    '            mapaRun.Add kRun, ObterOuCriarRUN(CDate(regs(iR, 2)), NucleoLote(CStr(regs(iR, 4))))',
    '        End If',
    '        regs(iR, 1) = mapaRun(kRun)',
    '    Next iR',
    ''
)

# Idempotencia: se o remapeamento ja esta la, nao inserir de novo. Sem isto,
# rodar o script duas vezes sobre o mesmo artefato duplica o bloco -- e a
# segunda copia sobrescreve regs com o mesmo resultado, mas o codigo fica
# ilegivel e a proxima manutencao tropeca.
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    if ($cm.Lines($i, 1) -match 'mapaRun') {
        "  frmMassa   remapeamento ja presente (L$i) - nada a fazer"
        $inserido = $true
        break
    }
}

if (-not $inserido) {
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    if ($cm.Lines($i, 1) -match '^\s*ReDim\s+final_\(1\s+To\s+n') {
        $cm.InsertLines($i, ($remap -join "`r`n"))
        "  frmMassa   L$i  remapeamento de RUN inserido ($($remap.Count) linhas)"
        $alteracoes++
        $inserido = $true
        break
    }
}
}
if (-not $inserido) { throw 'Nao encontrei ReDim final_ em frmMassa para inserir o remapeamento' }

"alteracoes: $alteracoes"

# ---------- conferencia: nenhuma chamada a NovoRUN sobrou ----------
$restantes = 0
foreach ($c in $proj.VBComponents) {
    $m = $c.CodeModule
    for ($i = 1; $i -le $m.CountOfLines; $i++) {
        $t = $m.Lines($i, 1)
        if ($t -match 'NovoRUN' -and $t -notmatch '^\s*''' -and $t -notmatch 'Function NovoRUN' -and $t -notmatch 'NovoRUN\s*=') {
            "  RESTOU: $($c.Name) L$i  $($t.Trim())"
            $restantes++
        }
    }
}
if ($restantes -eq 0) { "conferencia: nenhuma chamada a NovoRUN restante" }

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
