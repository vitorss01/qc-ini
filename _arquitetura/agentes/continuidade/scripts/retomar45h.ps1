$env:PYTHONIOENCODING = 'utf-8'
$env:QC_RAIZ = '<QC_RAIZ>'
$E = Join-Path $env:QC_RAIZ '_entrega_2026-11-18_0047\Hematologia'
$log = "$env:TEMP\qcw\retomar45h.log"
Set-Location "$env:TEMP\qcw\clean45\qc-ini\_arquitetura\integracao_dbseac\testes"
python -u -W ignore rodar_suites.py Hematologia "$E\QC_Hematologia.xlsm" "$E\retomada" --original "$E\producao_original\QC_Hematologia.xlsm" --suites qa_manual_substitui.py,qa_adr086.py,qa_adr088.py *>&1 | Out-File -Encoding utf8 $log
'FIM_RETOMADA' | Out-File -Append -Encoding utf8 $log
