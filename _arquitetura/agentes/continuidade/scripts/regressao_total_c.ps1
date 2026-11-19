$env:PYTHONIOENCODING = 'utf-8'
$R = '<QC_RAIZ>'
$env:QC_RAIZ = $R
$Q = "$env:TEMP\qcw"
$D = "$Q\regtotal"
$log = "$D\run.log"
"retomada $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (Hema a partir de qa_manual_substitui.py; as 26 anteriores ja OK)" | Out-File -Append -Encoding utf8 $log
Set-Location "$Q\clean46\qc-ini\_arquitetura\integracao_dbseac\testes"
$p = 'Hematologia'; $arq = "QC_$p.xlsm"
$h0 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
Remove-Item -Recurse -Force "$D\$p\teste_qa_manual_substitui" -ErrorAction SilentlyContinue
Copy-Item "$R\$arq" "$D\$p\base\$arq" -Force
Copy-Item "$R\$arq" "$D\$p\original\$arq" -Force
python -u -W ignore rodar_suites.py $p "$D\$p\base\$arq" "$D\$p" --original "$D\$p\original\$arq" --a-partir qa_manual_substitui.py *>&1 | Out-File -Append -Encoding utf8 $log
$h1 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
"   producao depois $h1 (intacta: $($h0 -eq $h1))" | Out-File -Append -Encoding utf8 $log
'FIM' | Out-File -Append -Encoding utf8 $log
