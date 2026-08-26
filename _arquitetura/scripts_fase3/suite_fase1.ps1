# suite_fase1.ps1 - a bateria que fecha a Fase 1, num comando so
#
# POR QUE EXISTE
#
# A Fase 1 acumulou seis verificacoes independentes, cada uma com seu script e
# sua forma de reportar. Roda-las a mao, uma a uma, nos dois produtos, e como o
# resultado de "todas passaram" foi ficando dependente de eu lembrar de rodar
# todas -- que e o tipo de garantia que se perde no primeiro dia corrido.
#
# Aqui a bateria e uma coisa so, com codigo de saida: 0 quando tudo passa,
# diferente de 0 quando qualquer teste reprova. O que nao roda aparece como
# PULADO e conta como falha, para "nao rodou" nunca se parecer com "passou".
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\suite_fase1.ps1
#   .\suite_fase1.ps1 -Produto Hematologia

param(
    [ValidateSet('Ambos', 'Bioquimica', 'Hematologia')][string]$Produto = 'Ambos'
)

$ErrorActionPreference = 'Continue'
$py = 'C:\Users\vitor.santos\AppData\Local\Programs\Python\Python312\python.exe'
$s = $PSScriptRoot

$artefatos = @{
    'Bioquimica'  = 'C:\Users\vitor.santos\QCINI_build_hardening1_Bioquimica\QC_Bioquimica.xlsm'
    'Hematologia' = 'C:\Users\vitor.santos\QCINI_build_hardening1_Hematologia\QC_Hematologia.xlsm'
}

# nome ; script ; so neste produto (vazio = os dois)
$TESTES = @(
    @('Westgard por modulo',       'testar_westgard_por_modulo.py',  ''),
    @('Cobertura x contrato',      'testar_cobertura_westgard.py',   ''),
    @('Eventos (grao de evento)',  'testar_eventos_westgard.py',     ''),
    @('Pior nivel com 3 niveis',   'testar_pior_nivel_3niveis.py',   'Hematologia'),
    @('Painel exibe o que o motor calcula', 'qa_painel_westgard.py',  '')
)

$linhas = @()
$falhou = 0

$alvos = if ($Produto -eq 'Ambos') { @('Bioquimica', 'Hematologia') } else { @($Produto) }

foreach ($p in $alvos) {
    $arq = $artefatos[$p]
    "=" * 78
    "SUITE FASE 1 -- $p"
    "=" * 78
    if (-not (Test-Path $arq)) {
        "  artefato ausente: $arq"
        $linhas += , @($p, 'artefato', 'AUSENTE', 0)
        $falhou++
        continue
    }

    foreach ($t in $TESTES) {
        $nome, $script, $so = $t
        if ($so -and $so -ne $p) { continue }
        $caminho = Join-Path $s $script
        if (-not (Test-Path $caminho)) {
            "  {0,-40} PULADO (script ausente)" -f $nome
            $linhas += , @($p, $nome, 'PULADO', 0)
            $falhou++
            continue
        }
        $saida = & $py -u $caminho $arq 2>&1
        $codigo = $LASTEXITCODE
        # A ultima linha "TOTAL:" e o placar de cada script.
        $placar = ($saida | Where-Object { $_ -match '^TOTAL:' } | Select-Object -Last 1)
        if ($null -eq $placar) { $placar = '(sem placar)' }
        $estado = if ($codigo -eq 0) { 'PASS' } else { 'FAIL' }
        if ($codigo -ne 0) { $falhou++ }
        "  {0,-40} {1,-6} {2}" -f $nome, $estado, $placar
        if ($codigo -ne 0) {
            $saida | Where-Object { $_ -match 'FAIL|Error|Traceback|erro' } |
                Select-Object -First 6 | ForEach-Object { "        $_" }
        }
        $linhas += , @($p, $nome, $estado, $codigo)
    }

    # A projecao Cfg_Westgard_Escopo tem seu proprio modo de conferencia.
    $saida = & $py -u (Join-Path $s 'materializar_cfg_westgard.py') $arq '--conferir' 2>&1
    $codigo = $LASTEXITCODE
    $estado = if ($codigo -eq 0) { 'PASS' } else { 'FAIL' }
    if ($codigo -ne 0) { $falhou++ }
    "  {0,-40} {1,-6} {2}" -f 'Cfg_Westgard_Escopo sincronizada', $estado,
        (($saida | Select-Object -Last 1))
    $linhas += , @($p, 'Cfg_Westgard_Escopo', $estado, $codigo)
    ""
}

"=" * 78
"PLACAR DA FASE 1"
"=" * 78
foreach ($L in $linhas) {
    "  {0,-13} {1,-40} {2}" -f $L[0], $L[1], $L[2]
}
""
if ($falhou -eq 0) {
    "TUDO VERDE: $($linhas.Count) verificacoes, 0 falhas"
    exit 0
}
"FALHAS: $falhou de $($linhas.Count) verificacoes"
exit 1
