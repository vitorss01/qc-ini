$env:PYTHONIOENCODING = 'utf-8'
$R = '<QC_RAIZ>'
$env:QC_RAIZ = $R
$D = "$env:TEMP\qcw\t46"
New-Item -ItemType Directory -Force "$D\bio", "$D\hema" | Out-Null
$log = "$D\run.log"
"inicio $(Get-Date -Format 'HH:mm:ss')" | Out-File -Encoding utf8 $log
Set-Location "$R\_arquitetura\integracao_dbseac\testes"
foreach ($c in @(@('Bioquimica', 'QC_Bioquimica.xlsm', 'bio', 'qa_adr097.py'), @('Hematologia', 'QC_Hematologia.xlsm', 'hema', 'qa_manual_substitui.py'))) {
    $p, $arq, $sub, $s = $c
    $h0 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    Copy-Item "$R\$arq" "$D\$sub\$arq" -Force
    python -u -W ignore $s $p "$D\$sub\$arq" "$D\$sub\$s.json" *>&1 | Out-File -Encoding utf8 "$D\$sub\$s.log"
    $h1 = (Get-FileHash "$R\$arq").Hash.Substring(0, 16).ToLower()
    $fim = (Get-Content "$D\$sub\$s.log" | Select-String '===' | Select-Object -Last 1).Line
    "$p $s | $fim | producao $h0 -> $h1" | Out-File -Append -Encoding utf8 $log
}
'FIM' | Out-File -Append -Encoding utf8 $log
