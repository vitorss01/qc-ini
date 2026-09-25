# -*- coding: utf-8 -*-
"""xlh.py -- utilitario COM do Excel para a Etapa 1.

- Proxy com retry em 'chamada rejeitada' (Excel ocupado recalculando).
- Watchdog: toda chamada longa pode ter teto de tempo; estourou, o processo do
  Excel e morto (um erro de compilacao VBA vira modal invisivel e travaria o
  script para sempre).
"""
import os
import threading
import time

import pythoncom
import win32com.client as w
import win32process
import win32api
import win32con

REJ = (-2147418111, -2147417846, -2146777998, -2147352573)   # rejeitada, ocupado, membro nao encontrado (transitorio)


def _retry(fn, *a, **k):
    t0 = time.time()
    while True:
        try:
            return fn(*a, **k)
        except pythoncom.com_error as e:
            if e.args[0] in REJ and time.time() - t0 < 180:
                time.sleep(0.2)
                continue
            raise


def _wrap(v):
    if hasattr(v, '_oleobj_'):
        return P(v)
    if callable(v) and not isinstance(v, type):
        # metodo COM vinculado (ex.: ws.Range): o RESULTADO tambem precisa do proxy
        def chamada(*a, **k):
            a = [x._o if isinstance(x, P) else x for x in a]
            return _wrap(_retry(v, *a, **k))
        return chamada
    return v


class P:
    __slots__ = ('_o',)

    def __init__(self, o):
        object.__setattr__(self, '_o', o)

    def __getattr__(self, n):
        return _wrap(_retry(getattr, self._o, n))

    def __setattr__(self, n, v):
        if isinstance(v, P):
            v = v._o
        _retry(setattr, self._o, n, v)

    def __call__(self, *a, **k):
        a = [x._o if isinstance(x, P) else x for x in a]
        return _wrap(_retry(self._o, *a, **k))

    def __iter__(self):
        for x in _retry(iter, self._o):
            yield _wrap(x)

    @property
    def raw(self):
        return self._o


class Excel:
    def __init__(self, visivel=False, eventos=False):
        pythoncom.CoInitialize()
        self.xl = P(w.DispatchEx('Excel.Application'))
        self.xl.Visible = visivel
        self.xl.DisplayAlerts = False
        self.xl.EnableEvents = eventos
        self.xl.AutomationSecurity = 1
        self.xl.Interactive = False          # mLotes.Avisar nao abre modal
        _, self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)

    def abrir(self, caminho):
        self.wb = self.xl.Workbooks.Open(caminho, 0)
        self.esperar()
        return self.wb

    def esperar(self, teto=900):
        """Espera o Excel parar de recalcular (chamadas sao rejeitadas enquanto isso)."""
        t0 = time.time()
        while time.time() - t0 < teto:
            try:
                x = self.xl.raw
                # em MANUAL o estado fica "pendente" ate alguem calcular: basta responder
                if x.CalculationState == 0 or x.Calculation == -4135:
                    return time.time() - t0
            except pythoncom.com_error:
                pass
            time.sleep(0.5)
        raise TimeoutError('Excel nao ficou pronto')

    def run(self, macro, *args, teto=300):
        """Application.Run com watchdog: estourou o teto, mata o Excel."""
        feito = threading.Event()
        estourou = []

        def cao():
            if not feito.wait(teto):
                estourou.append(True)
                self.matar()
        th = threading.Thread(target=cao, daemon=True)
        th.start()
        try:
            return self.xl.Run(macro, *args)
        except Exception:
            if estourou:
                raise TimeoutError(f'{macro} passou de {teto}s (modal/compilacao?) -- Excel encerrado')
            raise
        finally:
            feito.set()

    def matar(self):
        import subprocess
        try:
            subprocess.run(['taskkill', '/F', '/PID', str(self.pid)], capture_output=True, timeout=30)
        except Exception:
            pass

    def fechar(self, salvar=False):
        try:
            if getattr(self, 'wb', None) is not None:
                if salvar:
                    self.wb.Save()
                self.wb.Close(False)
        except Exception:
            pass
        try:
            self.xl.Quit()
        except Exception:
            pass
        time.sleep(1.5)
        self.matar()   # garante que nao sobra zumbi segurando o arquivo
