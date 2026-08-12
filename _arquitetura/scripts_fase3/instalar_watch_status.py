# -*- coding: utf-8 -*-
"""instalar_watch_status.py - as flags voltam a ser auto-corretivas (ADR-025)

O BURACO QUE ISTO FECHA

Enquanto BA:BC eram formula, qualquer mudanca de Status recalculava sozinha. Ao
virarem VALOR, passaram a depender de alguem chamar AtualizarFlagsBanco --
UpsertResultados e ExcluirLogico chamam, mas EDITAR A CELULA DIRETO nao passa por
nenhum dos dois. As flags ficavam obsoletas em silencio, que e exatamente o modo
de falha que o ADR-025 existe para eliminar.

Foi o que derrubou a prova 4.2: a suite exclui logicamente uma linha e depois
restaura o Status escrevendo direto na celula
(verificar_tudo.ps1: $db.Cells.Item($linhaTeste, 7).Value2 = $statusOriginal).
A exclusao recalculava as flags; a restauracao nao. A Liberacao ficava listando
as corridas do estado intermediario, e 52 celulas divergiam da referencia.

A GUARDA CONTRA RECURSAO

AtualizarFlagsBanco escreve em BA:BC (53..55), nunca na coluna G, entao o
Intersect ja bastaria. Mesmo assim o evento desliga EnableEvents e usa uma
trava de modulo: em VBA, evento que se redispara nao da erro -- da pilha
estourada minutos depois, longe da causa.

O bloco de restauracao roda SEMPRE, inclusive em excecao. Deixar EnableEvents
em False seria pior que o defeito original: a planilha inteira para de reagir.

CUSTO: AtualizarFlagsBanco leva 0,74 s com 93.000 registros. E o preco de uma
edicao manual de Status, que e operacao rara e deliberada.

Uso: python instalar_watch_status.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'

DECL = [
    "' ADR-025: trava de reentrancia do vigia de Status. Declarada aqui, antes",
    "' do primeiro procedimento -- Const/Dim de modulo depois de um Sub nao",
    "' existe para o compilador.",
    "Private mBancoAtualizandoFlags As Boolean",
    "",
]

HANDLER = [
    "",
    "' Vigia do Status (ADR-025).",
    "'",
    "' BA:BC sao VALOR, nao formula. Editar Status direto na celula nao passa por",
    "' UpsertResultados nem por ExcluirLogico, e sem este vigia as flags ficariam",
    "' obsoletas em silencio -- rFirst/rRunUnico apontariam para a linha errada e a",
    "' Liberacao, o Calc e a Estatistica passariam a contar corridas que nao",
    "' existem mais. Foi o que derrubou a prova 4.2.",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    "    Dim inter As Range",
    "    If mBancoAtualizandoFlags Then Exit Sub",
    "    On Error Resume Next",
    "    Set inter = Application.Intersect(Target, Me.Columns(COL_STATUS))",
    "    On Error GoTo 0",
    "    If inter Is Nothing Then Exit Sub",
    "",
    "    mBancoAtualizandoFlags = True",
    "    Application.EnableEvents = False",
    "    On Error GoTo fim",
    "    AtualizarFlagsBanco",
    "fim:",
    "    ' restaura SEMPRE: EnableEvents preso em False deixaria a pasta inteira",
    "    ' surda, o que e pior do que a falha que este vigia corrige.",
    "    Application.EnableEvents = True",
    "    mBancoAtualizandoFlags = False",
    "End Sub",
]


def novo_excel():
    for t in range(1, 8):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t == 2:
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(6)
            time.sleep(2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura')
    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)

        comp = None
        for c in wb.VBProject.VBComponents:
            try:
                if c.Type == 100 and c.Properties('Name').Value == 'DB_Resultados':
                    comp = c
            except Exception:
                pass
        if comp is None:
            raise SystemExit('modulo de planilha do DB_Resultados nao encontrado')
        cm = comp.CodeModule
        txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ''

        if 'mBancoAtualizandoFlags' in txt:
            print('vigia de Status ja instalado em %s -- nada a fazer' % comp.Name)
        else:
            # SO PODE HAVER UM Worksheet_Change por modulo de planilha. Dois
            # fazem o VBA rejeitar o modulo inteiro com "nome ambiguo", e o erro
            # aparece na primeira rotina chamada, longe da causa.
            if 'Private Sub Worksheet_Change' in txt or 'Sub Worksheet_Change' in txt:
                raise SystemExit(
                    'ja existe um Worksheet_Change no modulo do DB_Resultados -- '
                    'a chamada tem de ser INSERIDA nele, nao um segundo evento')
            # declaracao antes do primeiro procedimento
            ini = cm.CountOfDeclarationLines
            cm.InsertLines(ini + 1, '\r\n'.join(DECL))
            cm.InsertLines(cm.CountOfLines + 1, '\r\n'.join(HANDLER))
            print('vigia de Status instalado no modulo %s' % comp.Name)

        # conferencia: exatamente um Worksheet_Change
        txt = cm.Lines(1, cm.CountOfLines)
        n = txt.count('Sub Worksheet_Change')
        if n != 1:
            raise SystemExit('Worksheet_Change aparece %d vez(es) -- esperado 1' % n)
        if 'mBancoAtualizandoFlags As Boolean' not in txt:
            raise SystemExit('trava de reentrancia nao entrou')
        print('conferido: 1 Worksheet_Change, trava de reentrancia presente')

        if estrutura and not wb.ProtectStructure:
            wb.Protect(SENHA, True, False)
        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    main(sys.argv[1])
