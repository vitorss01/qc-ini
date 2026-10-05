# -*- coding: utf-8 -*-
"""instalar_adr062.py -- ADR-062: troca de analito quase imediata no Painel.

Uso:  python instalar_adr062.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

  1. todo o VBA atual (codigo_atual.py): motor aquecido no login e ao entrar no Painel, um redesenho
     por clique, lote resolvido numa leitura, protecao rapida no caminho do clique (sem o ciclo de
     senha de ~150 ms quando UserInterfaceOnly esta em vigor), lista C3 com uma recalculacao;
  2. Spinner 1 limitado aos analitos cadastrados (ia ate 40; acima do ultimo caia em posicao vazia);
  3. fumaca que CHAMA as rotinas novas;
  4. devolve a protecao de cada aba como estava e salva.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402


def log(*a):
    print(*a, flush=True)


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)

        log('1. VBA (todo o codigo atual das fontes)')
        codigo_atual.instalar(ex, wb, produto, log)

        log('2. Spinner 1 limitado aos analitos cadastrados')
        p = wb.Worksheets('Painel')
        sp = p.Spinners('Spinner 1')
        n = int(ex.xl.WorksheetFunction.CountA(wb.Worksheets('Analitos').Range('A4:A43')))
        sp.Min = 1
        sp.Max = max(1, n)
        if not (1 <= int(p.Range('B3').Value or 0) <= n):
            p.Range('B3').Value = 1
        log(f'   Min 1, Max {int(sp.Max)} ({n} analitos)')

        log('3. fumaca (chama as rotinas novas)')
        run = lambda m, *a: ex.run("'" + wb.Name + "'!" + m, *a, teto=300)
        run('mEstatistica.AquecerMotor')
        run('mEstatistica.AquecerMotor')                 # quente: sai na hora
        run('mUI.SpinnerLimites')
        r = run('mEstatistica.GarantirMotorDoPainel')
        if int(sp.Max) != n:
            raise SystemExit(f'SpinnerLimites deixou Max={sp.Max}, esperado {n}')
        log(f'   AquecerMotor 2x, SpinnerLimites, GarantirMotorDoPainel -> {r}')

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
