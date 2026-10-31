# Tester/QA: modo REPRODUZIR, ciclo 2026-10-31_ensaio (ensaio de integração)

Nada foi executado: não abri o Excel, não copiei arquivos e não rodei suítes. Os testes ficam em NAO_EXECUTADO (motivo: ensaio).

No ciclo real, o Tester faria isto:
1. Copiaria `QC_Bioquimica.xlsm` (com `shutil.copy2`) para `%TEMP%\qcw\auditoria\2026-10-31_ensaio\`, registraria o sha256 curto da produção e não tocaria nos arquivos da raiz.
2. Ocorrência 1 (Entrada CIQ, Alto): escreveria `ferramentas/repro_<id>.py` com `qa_final.QA`, lançaria valores sintéticos marcados `TESTE-AUDITORIA` e compararia o `Value2` gravado com o esperado.
3. Ocorrência 2 (Estatística/Sigma, Crítico): recalcularia bias, CV, ET e Sigma pelo método dos ADR (035/064/067/068/069) e compararia com uma tolerância relativa de 1e-9.
4. Ampliaria as duas ocorrências para outros analitos, níveis e lotes, e para a Hematologia.
5. Usaria um Excel por vez e encerraria só o `EXCEL_PID` do meu script. Não capturaria a tela.
6. Daria o veredito REPRODUZIDA, NAO_REPRODUZIDA ou PARCIAL com a evidência, e proporia uma suíte `qa_*.py` e o ID do caso.

| Ocorrência | Veredito | Estado do teste |
|---|---|---|
| [ENSAIO] Entrada CIQ (Alto) | NAO_EXECUTADO | ensaio |
| [ENSAIO] Cálculo Sigma (Crítico) | NAO_EXECUTADO | ensaio |
