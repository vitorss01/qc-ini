# -*- coding: utf-8 -*-
"""qa_sync.py -- ADR-079: gate da camada de dados em VBA (MOTOR_DADOS = VBA).

Uso: python qa_sync.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

Os cenarios estao em estudos/teste_incremental.py (incremental com resultado atrasado, idempotencia, origem
inacessivel sem gravar nada, trava contra a Principal encolher, recuperacao e a reserva MOTOR_DADOS = PQ dando a
mesma tabela). Sai com 1 se algum falhar.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'estudos'))
from teste_incremental import main  # noqa: E402

if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
