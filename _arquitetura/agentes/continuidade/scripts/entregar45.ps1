$env:PYTHONIOENCODING = 'utf-8'
$log = "$env:TEMP\qcw\entrega45.log"
$origem = '<CAMINHO_COMPLETO_DO_DB_SEAC.xlsm>'   # preencher SO na maquina do trabalho (copiar do entregar42.ps1 local); NUNCA versionar
if (-not (Test-Path -LiteralPath $origem)) { "ORIGEM_NAO_ENCONTRADA" | Out-File -Encoding utf8 $log; 'FIM_ENTREGA' | Out-File -Append -Encoding utf8 $log; exit 1 }
$env:QC_RAIZ = '<QC_RAIZ>'
Set-Location "$env:TEMP\qcw\clean45\qc-ini\_arquitetura\integracao_dbseac"
python -u -W ignore entregar.py Bioquimica Hematologia --modo SEAC --caminho $origem --direcionada --com qa_manual_substitui.py,qa_casos_extremos.py,qa_lotes_auto.py,qa_linhas_alvo.py,qa_inativar.py,qa_adr086.py,qa_adr088.py *>&1 | Out-File -Encoding utf8 $log
'FIM_ENTREGA' | Out-File -Append -Encoding utf8 $log
