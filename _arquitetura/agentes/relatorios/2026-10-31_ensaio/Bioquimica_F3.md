# Bioquimica - F3 (ensaio 2026-10-31_ensaio)

Ensaio de integracao: nada foi executado. No ciclo real eu faria:
1. Copiar QC_Bioquimica.xlsm com shutil.copy2 para %TEMP%\qcw\auditoria\2026-10-31_ensaio\ (sem abrir Excel nesta frente).
2. Ler a copia com openpyxl: tblCQ_Final e EQA_Base (brutos) e Estatistica/Painel/Eng_Saida/Calc com data_only=True.
3. Confirmar em mCEQ.BiasEQ e ADR-035/064/068/069 a definicao de bias, CV pooled (n>=20, 6 meses), ET e Sigma.
4. Escolher 5 analitos por risco (2 lotes, ETp de fonte diferente, sem ETp, CEQ em varias rodadas, rotina estavel).
5. Recalcular com _arquitetura/agentes/ferramentas/calc_independente.py e montar a tabela do Relatorio B.
6. Conferir sinal/escala %, provedor/ano/rodada (sem misturar CAP e Controllab), ETp do cadastro e "-" para ETp ausente.
7. Desenhar o caminho CEQ -> bias -> ET -> Sigma -> Painel e CIQ -> CV; a verificacao de inclusao/correcao fica com F1/F2.

Ocorrencia: apenas a ficticia [ENSAIO] exigida pelo roteiro.
