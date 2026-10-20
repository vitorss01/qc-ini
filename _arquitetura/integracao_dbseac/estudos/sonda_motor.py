# -*- coding: utf-8 -*-
"""Perfil do MOTOR do ATUALIZAR (graficos, Westgard, Estatistica, incerteza) etapa por etapa. Roda numa COPIA.

Uso: python sonda_motor.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--n 3]

Repete, dentro do Excel e com o mesmo envelope do botao (mApp.InicioUsuario/fim: tela, eventos e calculo
desligados), a sequencia do motor de mIntegracao.AtualizarDadosCore e mede cada chamada com o Timer do VBA.
"""
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
from qa_final import QA  # noqa: E402

MODULO = r'''
Option Explicit

Private Function Rodar(ByVal nome As String) As Boolean
    On Error GoTo nao
    Application.Run nome
    Rodar = True
nao:
End Function

Public Function PerfilMotor() As String
    Dim t As Double, s As String, t0 As Double
    mApp.InicioUsuario "perfil do motor"
    t0 = Timer
    t = Timer: mEstatistica.InvalidarCache: s = s & "InvalidarCache=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: Rodar "InvalidarCacheEstat": s = s & "InvalidarCacheEstat=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mDados.AtualizarListasAno: s = s & "AtualizarListasAno=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mLotes.AtualizarListaLiberacao: s = s & "AtualizarListaLiberacao=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mEstatistica.AtualizarEstatistica: s = s & "AtualizarEstatistica=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: Rodar "RecalcularEstatPeriodo": s = s & "RecalcularEstatPeriodo=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mIncerteza.RecalcularIncerteza: s = s & "RecalcularIncerteza=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mApp.fim: s = s & "fim(calculo automatico)=" & Format$(Timer - t, "0.000") & ";"
    PerfilMotor = s & "TOTAL=" & Format$(Timer - t0, "0.000")
End Function

Public Function PerfilEstatistica() As String
    ' por dentro do AtualizarEstatistica: calcular a pasta inteira (o que o Application.Calculate custa sozinho)
    Dim t As Double
    mApp.InicioUsuario "perfil"
    t = Timer: Application.CalculateFull: PerfilEstatistica = "CalculateFull=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: Application.Calculate: PerfilEstatistica = PerfilEstatistica & "Calculate(nada sujo)=" & Format$(Timer - t, "0.000")
    mApp.fim
End Function

Public Function PerfilFino() As String
    ' por dentro do AtualizarEstatistica e da incerteza, com o snapshot do ADR-079
    Dim t As Double, s As String, t0 As Double
    mApp.InicioUsuario "perfil fino"
    t0 = Timer
    mDados.IniciarSnapshot
    t = Timer: mEstatistica.InvalidarCache: Rodar "InvalidarCacheEstat": s = s & "invalidar=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mDados.AtualizarListasAno: s = s & "ListasAno=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mLotes.AtualizarListaLiberacao: s = s & "ListaLiberacao=" & Format$(Timer - t, "0.000") & ";"
    Application.Calculation = xlCalculationManual
    t = Timer: mEstatistica.AtualizarCalc: s = s & "AtualizarCalc=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mEstatistica.RegistrarEventosWestgard: s = s & "Westgard=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mEstatistica.AtualizarPainelEng: s = s & "PainelEng=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mEstatistica.AtualizarEstatisticaAba: s = s & "EstatisticaAba(+Westgard)=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: Application.Calculate: s = s & "Calculate1=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: Rodar "RecalcularEstatPeriodo": s = s & "EstatPeriodo=" & Format$(Timer - t, "0.000") & ";"
    t = Timer: mIncerteza.RecalcularIncerteza: s = s & "Incerteza(Calculate)=" & Format$(Timer - t, "0.000") & ";"
    mDados.EncerrarSnapshot
    t = Timer: mApp.fim: s = s & "fim=" & Format$(Timer - t, "0.000") & ";"
    PerfilFino = s & "TOTAL=" & Format$(Timer - t0, "0.000")
End Function

Public Function PerfilCarregarDB() As String
    Dim t As Double, v As Variant
    t = Timer: v = mDados.CarregarDB(): PerfilCarregarDB = "CarregarDB=" & Format$(Timer - t, "0.000")
End Function
'''


def main(produto, arq, saida, n=3):
    res = {'produto': produto, 'perfis': []}
    q = QA(produto, arq)
    try:
        q.run('mSeguranca.ReprotectAll')
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mSondaMotor'
        comp.CodeModule.AddFromString(MODULO)
        for i in range(n):
            r = str(q.run('mSondaMotor.PerfilMotor', teto=900))
            res['perfis'].append(r)
            print('motor', i + 1, r, flush=True)
        for fn in ('PerfilFino', 'PerfilFino', 'PerfilCarregarDB', 'PerfilEstatistica'):
            r = str(q.run('mSondaMotor.' + fn, teto=900))
            res[fn] = r
            print(fn, r, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], int(a[a.index('--n') + 1]) if '--n' in a else 3)
