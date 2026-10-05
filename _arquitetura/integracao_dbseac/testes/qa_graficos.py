# -*- coding: utf-8 -*-
"""QA dos GRAFICOS DO PAINEL -- todos os niveis a vista (ADR-063) e vigia do zoom (ADR-061).
Excel VISIVEL, numa COPIA instalada. Nada e salvo no arquivo testado.

Uso: python qa_graficos.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Pedido do usuario (03/10/2026): "sao 2 graficos de Levey-Jennings na Bioquimica e 3 na Hematologia;
nao pode exibir apenas um na tela; precisa garantir boa visualizacao de todos". A prova de "inteiro
na tela" e do proprio Excel (VisibleRange, que inclui celula parcialmente visivel): a legenda do
lote (linha 39) so esta inteira se a linha 40 aparece, e o cabecalho A:U so esta inteiro se a coluna V
aparece. Os graficos acabam antes da linha 39. A janela e fotografada (PrintWindow: nao precisa
estar na frente) e a foto vai junto da evidencia.

O Excel nao tem evento de zoom: G02 muda o zoom e NAO clica em nada -- a largura tem de se
reajustar sozinha (vigia). G07 fecha com o vigia ligado e confere que o Excel NAO reabre o arquivo.
"""
import ctypes
import json
import os
import sys
import time

import win32con
import win32gui
import win32ui
from PIL import Image

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)      # pixels reais (tela com escala, ex. 125%)
except Exception:
    pass

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402

RES = []
NIVEIS = {'Bioquimica': 2, 'Hematologia': 3}
ALT_MIN = 130.0          # mUI.GRAF_ALT_MIN
LINHA_LEGENDA = 39       # mUI.PAINEL_FAIXA_FIM + 1


def reg(teste, ok, evid):
    RES.append({'teste': teste, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': evid})
    print(('PASS ' if ok else 'FAIL ') + teste + ' -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:600], flush=True)


def esperar(cond, teto=4.0):
    t0 = time.time()
    while time.time() - t0 < teto:
        try:
            if cond():
                return time.time() - t0
        except Exception:
            pass
        time.sleep(0.1)
    return None


def px_por_pt():
    try:
        return ctypes.windll.user32.GetDpiForSystem() / 72.0
    except Exception:
        return 96 / 72.0


_CAB = {}


def cabecalhos_pt(xl, w):
    """Cabecalhos de linha/coluna em pontos de tela (UsableWidth/Height os incluem) -- medido aqui,
    independente do VBA: liga/desliga com a tela congelada; guardado por zoom."""
    if not w.DisplayHeadings:
        return 0.0, 0.0
    z = float(w.Zoom)
    if z in _CAB:
        return _CAB[z]
    salva = bool(xl.ActiveWorkbook.Saved)            # ligar/desligar cabecalho marca a pasta como alterada
    xl.ScreenUpdating = False
    try:
        x1, y1 = w.PointsToScreenPixelsX(0), w.PointsToScreenPixelsY(0)
        w.DisplayHeadings = False
        x0, y0 = w.PointsToScreenPixelsX(0), w.PointsToScreenPixelsY(0)
        w.DisplayHeadings = True
    finally:
        xl.ScreenUpdating = True
        if salva:
            xl.ActiveWorkbook.Saved = True
    k = px_por_pt()
    _CAB[z] = ((x1 - x0) / k, (y1 - y0) / k)
    return _CAB[z]


def largura_visivel(xl, w):
    lx, _ = cabecalhos_pt(xl, w)
    return (w.UsableWidth - lx) * 100.0 / float(w.Zoom)


def geo(p):
    return [(round(co.Top, 1), round(co.Height, 1), bool(co.Visible)) for co in p.ChartObjects()]


def encaixe(xl, p, w, n):
    """Todos os n graficos visiveis, inteiros na tela, sem sobrepor, mesma altura, legenda e cabecalho inteiros."""
    vr = w.VisibleRange
    ult_lin = vr.Row + vr.Rows.Count - 1
    ult_col = vr.Column + vr.Columns.Count - 1
    leg = p.Rows(LINHA_LEGENDA).Top
    vis = largura_visivel(xl, w)
    cos = sorted(p.ChartObjects(), key=lambda c: c.Top)
    g = [(round(c.Top, 1), round(c.Top + c.Height, 1), round(c.Height, 1), round(c.Left + c.Width, 1), bool(c.Visible))
         for c in cos]
    alturas = [x[2] for x in g]
    ev = {'zoom': w.Zoom, 'visivel': vr.Address.replace('$', ''), 'graficos_topo_fundo_altura_direita': [x[:4] for x in g],
          'linha_legenda_topo': round(leg, 1), 'largura_visivel_pt': round(vis, 1)}
    ok = {
        'quantos': len(g) == n,
        'todos_visiveis': all(x[4] for x in g),
        'dentro_vertical': all(x[0] >= vr.Top - 0.5 and x[1] <= leg + 0.5 for x in g),
        'legenda_do_lote_inteira': ult_lin >= LINHA_LEGENDA + 1,
        'cabecalho_A_U_inteiro': ult_col >= 22,
        'dentro_horizontal': all(x[3] <= vis + 0.5 for x in g),
        'sem_sobrepor': all(g[i][1] <= g[i + 1][0] + 0.5 for i in range(len(g) - 1)),
        'mesma_altura_legivel': min(alturas) >= ALT_MIN - 0.5 and max(alturas) - min(alturas) <= 1.0,
        'rolagem_no_topo': w.ScrollRow == 1 and w.ScrollColumn == 1,
    }
    ev['criterios'] = ok
    return all(ok.values()), ev


def foto(hwnd, arq):
    """Foto da janela pelo PrintWindow (funciona mesmo atras de outra janela)."""
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w, h = max(1, r - l), max(1, b - t)
        hdc = win32gui.GetWindowDC(hwnd)
        mdc = win32ui.CreateDCFromHandle(hdc)
        sdc = mdc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mdc, w, h)
        sdc.SelectObject(bmp)
        ctypes.windll.user32.PrintWindow(hwnd, sdc.GetSafeHdc(), 2)
        info = bmp.GetInfo()
        Image.frombuffer('RGB', (info['bmWidth'], info['bmHeight']), bmp.GetBitmapBits(True), 'raw', 'BGRX', 0, 1).save(arq)
        win32gui.DeleteObject(bmp.GetHandle())
        sdc.DeleteDC()
        mdc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hdc)
        return arq
    except Exception as e:
        return f'sem foto: {e}'


def primeiro_plano(hwnd):
    """Traz o Excel para a frente sem mandar tecla (so para a foto ficar igual a tela; PrintWindow
    nao depende disso)."""
    try:
        import win32com.client
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32com.client.Dispatch('WScript.Shell').AppActivate(win32gui.GetWindowText(hwnd))
    except Exception:
        pass
    time.sleep(0.3)
    return win32gui.GetForegroundWindow() == hwnd


def abrir(caminho, eventos=True):
    ex = xlh.Excel(visivel=True, eventos=False)
    wb = ex.abrir(caminho)
    ex.xl.WindowState = -4137                       # maximizada: area util real
    try:
        wb.Unprotect('qcini2025')
    except Exception:
        pass
    for ws in wb.Worksheets:
        if ws.Visible != -1:
            ws.Visible = -1
    ex.xl.Interactive = True
    ex.xl.EnableEvents = eventos
    return ex, wb


def executar(produto, caminho, saida):
    ex, wb = abrir(caminho)
    run = lambda m, *a: ex.run("'" + wb.Name + "'!" + m, *a, teto=60)
    try:
        return _executar(produto, caminho, saida, ex, wb, run)
    finally:
        ex.fechar()                                   # nunca deixa Excel visivel para tras


def _executar(produto, caminho, saida, ex, wb, run):
    n = NIVEIS['Bioquimica' if produto.startswith('Bio') else 'Hematologia']
    base = os.path.splitext(saida)[0] if saida else os.path.splitext(caminho)[0]
    try:
        run('mSeguranca.ReprotectAll')               # sessao de usuario (UserInterfaceOnly)
        run('mUI.SystemLook', True)                  # como o usuario ve depois do login
        wb.Worksheets('Início').Activate()
        primeiro_plano(ex.xl.Hwnd)
        p = wb.Worksheets('Painel')
        p.Activate()                                 # Worksheet_Activate -> encaixe + vigia
        time.sleep(1.0)
        w = ex.xl.ActiveWindow

        # ------------------------------------------------------------ G01 todos a vista ao entrar
        ok, ev = encaixe(ex.xl, p, w, n)
        ev['foto'] = foto(ex.xl.Hwnd, base + '_painel.png')
        reg(f'G01 Entrar no Painel: os {n} gráficos de Levey-Jennings visíveis e INTEIROS na tela, lado a lado na '
            'vertical, mesma altura (≥ 130 pt), cabeçalho A:U e legenda do lote (linha 39) inteiros', ok, ev)
        geo0 = geo(p)

        est = str(run('mUI.GrafVigiaEstado'))
        reg('G02 Abrir o Painel liga o vigia de zoom (só com gente olhando: Interactive e eventos ligados)',
            not est.startswith('0|'), {'estado': est})

        # ------------------------------------------------------------ G03 zoom manual sem clique
        res = {}
        for z in (130, 70, 160):
            wb.Saved = True
            w.Zoom = z
            vis = largura_visivel(ex.xl, w)
            dt = esperar(lambda: all(abs(co.Left + co.Width - (vis - 6)) <= 3 for co in p.ChartObjects()), 4.0)
            res[z] = {'reagiu_em_s': None if dt is None else round(dt, 2), 'zoom_mantido': w.Zoom == z,
                      'larguras': [round(co.Width) for co in p.ChartObjects()],
                      'largura_visivel': round(vis), 'pasta_continua_salva': bool(wb.Saved),
                      'graficos_inalterados': geo(p) == geo0}
        reg('G03 Zoom manual (130% → 70% → 160%) SEM clique: o zoom do usuário vale, a largura dos gráficos acompanha '
            'a janela em até ~1 s, nenhum gráfico some nem muda de altura, a pasta não fica "suja"',
            all(r['reagiu_em_s'] is not None and r['reagiu_em_s'] <= 2.5 and r['zoom_mantido'] and r['pasta_continua_salva']
                and r['graficos_inalterados'] for r in res.values()), res)

        # ------------------------------------------------------------ G04 botao
        wb.Saved = True
        r = str(run('mUI.PainelEncaixar'))           # o que o botao "Ver todos os graficos" faz
        ok, ev = encaixe(ex.xl, p, w, n)
        ev['retorno'] = r
        ev['pasta_continua_salva'] = bool(wb.Saved)
        ev['botao'] = {'texto': p.Shapes('btnGrafTodos').TextFrame2.TextRange.Text,
                       'macro': p.Shapes('btnGrafTodos').OnAction}
        reg('G04 Botão "Ver todos os gráficos" depois de um zoom manual: volta a mostrar TODOS inteiros, sem sujar a pasta',
            ok and r.startswith('OK|') and ev['pasta_continua_salva'] and ev['botao']['texto'] == 'Ver todos os gráficos'
            and ev['botao']['macro'].endswith('PainelVerTodos'), ev)

        # ------------------------------------------------------------ G05 janela de outro tamanho
        res = {}
        ex.xl.WindowState = -4143                    # normal
        ex.xl.Width = 900
        ex.xl.Height = 520
        dt = esperar(lambda: encaixe(ex.xl, p, w, n)[0], 4.0)
        ok1, ev1 = encaixe(ex.xl, p, w, n)
        ev1['reagiu_em_s'] = None if dt is None else round(dt, 2)
        ev1['foto'] = foto(ex.xl.Hwnd, base + '_painel_janela_menor.png')
        res['janela_menor_900x520pt'] = ev1
        ex.xl.WindowState = -4137                    # maximizada de novo
        dt = esperar(lambda: encaixe(ex.xl, p, w, n)[0], 4.0)
        ok2, ev2 = encaixe(ex.xl, p, w, n)
        ev2['reagiu_em_s'] = None if dt is None else round(dt, 2)
        res['maximizada'] = ev2
        reg('G05 Janela redimensionada (menor e depois maximizada): o Painel se encaixa de novo sozinho, todos inteiros',
            ok1 and ok2, res)

        # ------------------------------------------------------------ G06 troca de analito
        sp = p.Spinners('Spinner 1')
        geo1 = geo(p)
        z1 = w.Zoom
        trocas = []
        for _ in range(3):
            v = int(sp.Value) + 1 if int(sp.Value) < int(sp.Max) else 1
            sp.Value = v
            run('mEstatistica.PainelMudou')          # OnAction do spinner
            trocas.append(str(p.Range('C3').Value))
        ok, ev = encaixe(ex.xl, p, w, n)
        ev['analitos'] = trocas
        ev['geometria_igual'] = geo(p) == geo1
        ev['zoom_igual'] = w.Zoom == z1
        reg('G06 Trocar de analito pelo spinner mantém todos os gráficos inteiros na tela (nada muda de lugar)',
            ok and ev['geometria_igual'] and ev['zoom_igual'], ev)

        # ------------------------------------------------------------ G07 sair do Painel
        wb.Worksheets('Início').Activate()
        est = str(run('mUI.GrafVigiaEstado'))
        visiveis = all(co.Visible for co in p.ChartObjects())
        reg('G07 Sair do Painel desliga o vigia; nenhum gráfico fica oculto',
            est.startswith('0|') and visiveis, {'vigia': est, 'todos_visiveis': visiveis})
    finally:
        try:
            run('mUI.SystemLook', False)
        except Exception:
            pass

    # ---------------------------------------------------------------- fechar com o vigia ligado: nao reabre
    # Fecha por AUTOMACAO (wb.Close), com outra pasta aberta (Excel de pe, como Bio e Hema juntas): e o
    # caminho dificil -- medido em 03/10/2026, nele o Excel ignorava o Application.Wait do BeforeClose e
    # o tique vencido REABRIA o arquivo ~5 s depois. A prova antiga olhava 4 s depois, antes de a
    # reabertura aparecer (carregar o arquivo leva ~4,7 s): falso positivo. Agora olha 10 s depois.
    # (Pelo X do usuario a espera ja funcionava: 2 de 2 rodadas com clique real, sem reabrir.)
    outra = ex.xl.Workbooks.Add()
    wb.Activate()
    wb.Worksheets('Painel').Activate()
    run('mUI.GrafVigiaLigar')
    time.sleep(1.6)                                   # pelo menos um tique ja rodou e reagendou
    ligado = not str(run('mUI.GrafVigiaEstado')).startswith('0|')
    nome = wb.Name
    t0 = time.time()
    wb.Close(False)
    fechou_em = round(time.time() - t0, 2)
    time.sleep(10)                                    # reabrir leva ~5 s: 10 s pega com folga
    try:
        abertos = [b.Name for b in ex.xl.Workbooks]
    except Exception as e:                            # caixa modal (ex.: "Quer salvar?") bloqueia a automacao
        abertos = [f'Excel ocupado/modal: {e}']
        import win32process

        def cb(h, _):                                 # fecha a caixa (= Cancelar) para o Excel poder ser encerrado
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) in ('NUIDialog', '#32770') \
                    and win32process.GetWindowThreadProcessId(h)[1] == ex.pid:
                abertos.append('caixa: ' + win32gui.GetWindowText(h))
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            return True
        win32gui.EnumWindows(cb, None)
    reg('G08 Fechar o arquivo com o vigia ligado (outra pasta aberta): o Excel NÃO reabre o arquivo (10 s depois)',
        # o que se prova e "nao reabriu": a outra pasta pode ja nao estar la (lista vazia quebrava a prova)
        ligado and nome not in abertos and not any(str(a).startswith('Excel ocupado') for a in abertos),
        {'vigia_estava_ligado': ligado, 'close_levou_s': fechou_em, 'pastas_abertas_10s_depois': abertos})

    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} gráficos: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
