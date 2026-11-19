$env:PYTHONIOENCODING = 'utf-8'
$R = '<QC_RAIZ>'
$env:QC_RAIZ = $R
$Q = "$env:TEMP\qcw"
$D = "$Q\regtotal"
$log = "$D\run.log"
"retomada $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (Bio a partir de qa_desempenho.py; as 12 anteriores ja OK)" | Out-File -Append -Encoding utf8 $log
Set-Location "$Q\clean46\qc-ini\_arquitetura\integracao_dbseac\testes"
Get-ChildItem "$D\Bioquimica" -Directory -Filter 'teste_*' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
foreach ($p in 'Bioquimica', 'Hematologia') {
    $arq = "QC_$p.xlsm"
    $h0 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    New-Item -ItemType Directory -Force "$D\$p\base", "$D\$p\original" | Out-Null
    Copy-Item "$R\$arq" "$D\$p\base\$arq" -Force
    Copy-Item "$R\$arq" "$D\$p\original\$arq" -Force
    "## $p producao $h0" | Out-File -Append -Encoding utf8 $log
    $extra = @()
    if ($p -eq 'Bioquimica') { $extra = @('--a-partir', 'qa_desempenho.py') }
    python -u -W ignore rodar_suites.py $p "$D\$p\base\$arq" "$D\$p" --original "$D\$p\original\$arq" @extra *>&1 | Out-File -Append -Encoding utf8 $log
    $h1 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    "   producao depois $h1 (intacta: $($h0 -eq $h1))" | Out-File -Append -Encoding utf8 $log
}
'FIM' | Out-File -Append -Encoding utf8 $log
