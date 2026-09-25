#!/usr/bin/env bash
# validar_tudo.sh -- instala nos dois produtos (a partir das copias de producao)
# e roda TODAS as provas: UX, suite T0-T9 (Bioquimica) e cinco anos (os dois).
# Uso: bash validar_tudo.sh   (log em $TMP/qcw/final/validar.log)
set -u
RAIZ="/c/Users/vitor/OneDrive - MSFT/Desktop/QC_INI"
AQUI="$RAIZ/_arquitetura/etapa1_multilote"
F="$TMP/qcw/final"
mkdir -p "$F"
export PYTHONIOENCODING=utf-8
cd "$AQUI" || exit 1
echo "== copias INTOCADAS (backup pre-entrega): instalar sempre a partir delas"
ORIG="${ORIG:-$RAIZ/_backup_pre_5anos_2026-09-19_1808}"
cp "$ORIG/QC_Bioquimica.xlsm" "$F/QC_Bioquimica.xlsm"
cp "$ORIG/QC_Hematologia.xlsm" "$F/QC_Hematologia.xlsm"
echo "== instalar Bioquimica"
python aplicar_etapa1.py "$F/QC_Bioquimica.xlsm" | tail -3
echo "== portar Hematologia"
python portar_hema.py "$F/QC_Hematologia.xlsm" "$AQUI/ref/QC_Hematologia_build_h1.xlsm" "$F/QC_Bioquimica.xlsm" | tail -3
echo "== desenho do Painel (ADR-054)"
python testes/prova_painel.py "$F/QC_Bioquimica.xlsm" bio | tail -3
python testes/prova_painel.py "$F/QC_Hematologia.xlsm" hema | tail -3
echo "== UX"
python testes/prova_ux.py "$F/QC_Bioquimica.xlsm" bio | tail -30
python testes/prova_ux.py "$F/QC_Hematologia.xlsm" hema | tail -30
echo "== suite T0-T9 Bioquimica"
python testes/suite_etapa1.py base "$F/QC_Bioquimica.xlsm" "$F/base_bio.xlsm" | tail -1
python testes/suite_etapa1.py todos "$F/base_bio.xlsm" | tail -3
echo "== cinco anos"
python testes/t_cinco_anos.py bio "$F/QC_Bioquimica.xlsm" | tail -22
python testes/t_cinco_anos.py hema "$F/QC_Hematologia.xlsm" | tail -22
echo FIM_VALIDACAO
