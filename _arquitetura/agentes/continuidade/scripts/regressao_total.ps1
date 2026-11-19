$env:PYTHONIOENCODING = 'utf-8'
$R = '<QC_RAIZ>'
$env:QC_RAIZ = $R
$Q = "$env:TEMP\qcw"
$D = "$Q\regtotal"
$log = "$D\run.log"
New-Item -ItemType Directory -Force $D | Out-Null
if (Test-Path "$Q\clean46") { Remove-Item -Recurse -Force "$Q\clean46" }
New-Item -ItemType Directory "$Q\clean46\qc-ini" | Out-Null
git -C $R archive --format=tar -o "$Q\clean46.tar" HEAD
& "$env:SystemRoot\System32\tar.exe" -xf "$Q\clean46.tar" -C "$Q\clean46\qc-ini"
"inicio $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') commit $(git -C $R rev-parse --short HEAD)" | Out-File -Encoding utf8 $log
Set-Location "$Q\clean46\qc-ini\_arquitetura\integracao_dbseac\testes"
foreach ($p in 'Bioquimica', 'Hematologia') {
    $arq = "QC_$p.xlsm"
    $h0 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    New-Item -ItemType Directory -Force "$D\$p\base", "$D\$p\original" | Out-Null
    Copy-Item "$R\$arq" "$D\$p\base\$arq" -Force
    Copy-Item "$R\$arq" "$D\$p\original\$arq" -Force
    "## $p producao $h0" | Out-File -Append -Encoding utf8 $log
    python -u -W ignore rodar_suites.py $p "$D\$p\base\$arq" "$D\$p" --original "$D\$p\original\$arq" *>&1 | Out-File -Append -Encoding utf8 $log
    $h1 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    "   producao depois $h1 (intacta: $($h0 -eq $h1))" | Out-File -Append -Encoding utf8 $log
}
'FIM' | Out-File -Append -Encoding utf8 $log
