# -*- coding: utf-8 -*-
"""sonda_cadeia_operacao.py - onde a cadeia pos-gravacao quebra

mOperacao.AtualizarOperacao e descrita no proprio codigo como "cadeia unica
chamada apos qualquer gravacao/exclusao". Chamada por COM, ela poe o VBA em
MODO DE INTERRUPCAO -- o diagnostico veio do titulo da janela do VBE, porque o
dialogo de erro e modal e invisivel com Visible=False, e travou a automacao por
dez minutos sem devolver nada.

Chamar de fora nao serve para diagnosticar: o erro vira dialogo antes de virar
resposta. A sonda injeta um modulo TEMPORARIO que chama cada elo com
On Error Resume Next e devolve Err.Number/Err.Description como TEXTO.

O modulo e removido no fim e o arquivo e fechado SEM salvar.
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SHIM = '''Option Explicit

Private Function Rodar(ByVal nome As String) As String
    Dim t As Single
    t = Timer
    On Error Resume Next
    Application.Run nome
    If Err.Number <> 0 Then
        Rodar = "ERRO " & Err.Number & " | " & Err.Description & _
                " | origem=" & Err.Source & " | " & Format$(Timer - t, "0.00") & "s"
        Err.Clear
    Else
        Rodar = "ok | " & Format$(Timer - t, "0.00") & "s"
    End If
    On Error GoTo 0
End Function

Public Function SondaElos() As String
    Dim r As String
    r = "AtualizarBanco............ " & Rodar("AtualizarBanco") & vbLf
    r = r & "AtualizarViewResultados... " & Rodar("AtualizarViewResultados") & vbLf
    r = r & "AtualizarEstatistica...... " & Rodar("AtualizarEstatistica") & vbLf
    r = r & "AtualizarPainel........... " & Rodar("AtualizarPainel") & vbLf
    SondaElos = r
End Function

Public Function SondaCadeia() As String
    SondaCadeia = "AtualizarOperacao......... " & Rodar("AtualizarOperacao")
End Function

' Qual copia de AtualizarEstatistica o mOperacao enxerga? A resposta honesta
' vem de dentro do proprio modulo: uma chamada SEM qualificar aqui resolve
' pelas mesmas regras que a do mOperacao (nenhum dos dois tem copia local).
Public Function SondaBind() As String
    Dim antes As String, depois As String
    On Error Resume Next
    antes = CStr(Application.Run("EstaPastaDeTrabalho.Name"))
    SondaBind = "n/d"
    On Error GoTo 0
End Function
'''


TRANS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=8):
    u = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            u = e
            if not any(t in str(e).lower() for t in TRANS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise u


def novo():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (1, 3, 5):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Minimized'],
                                stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    comp = None
    try:
        proj = wb.VBProject
        for c in list(proj.VBComponents):
            if c.Name == 'mSondaTmp':
                proj.VBComponents.Remove(c)
        comp = proj.VBComponents.Add(1)      # modulo padrao
        comp.Name = 'mSondaTmp'
        comp.CodeModule.AddFromString(SHIM)
        print('sonda temporaria instalada')
        print()

        print('=== cada elo isolado ===')
        r = tenta(lambda: xl.Run('SondaElos'), vezes=2)
        for l in str(r).split('\n'):
            if l.strip():
                print('  ' + l.rstrip())
        print()

        print('=== a cadeia inteira ===')
        r2 = tenta(lambda: xl.Run('SondaCadeia'), vezes=2)
        print('  ' + str(r2).strip())
    finally:
        try:
            if comp is not None:
                wb.VBProject.VBComponents.Remove(comp)
                print()
                print('sonda temporaria removida')
        except Exception:
            pass
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1])
