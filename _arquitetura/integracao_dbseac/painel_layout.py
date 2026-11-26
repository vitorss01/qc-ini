# -*- coding: utf-8 -*-
"""painel_layout.py -- ADR-103 (entrega 49, 26/11/2026): onde ficam, no topo do Painel, o status por nivel e a classe Sigma.

Pedido do usuario: os status por nivel (antes J7:J9) saem da interface e a "Classe Sigma" (antes K6:K9) ocupa o lugar
deles. As formulas de status NAO mudam: foram RECORTADAS para a coluna tecnica oculta AB (mesmas linhas, entao as
referencias relativas e o Q4 -- COUNTIF do status -- continuam certos; nome definido stNivel). Os instaladores que
regravam essas celulas (instalar_adr078/084/092/101) e as suites perguntam aqui o endereco: layout antigo (J = status,
K = classe) ou novo (AB = status oculto, J = classe). O layout novo e reconhecido pelo cabecalho J6 = "Classe Sigma".
"""
COL_STATUS_NOVA = 'AB'
CAB_CLASSE = 'Classe Sigma'


def layout_novo(p):
    """p: a aba Painel (COM)."""
    return str(p.Range('J6').Value or '').strip() == CAB_CLASSE


def col_status(p):
    return COL_STATUS_NOVA if layout_novo(p) else 'J'


def col_classe(p):
    return 'J' if layout_novo(p) else 'K'


def status_nivel(p, i):
    """Texto do status do nivel i (1..3), onde quer que esteja."""
    return p.Range(f'{col_status(p)}{6 + i}').Value


def status_celula(p, i):
    return p.Range(f'{col_status(p)}{6 + i}')
