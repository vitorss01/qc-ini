# -*- coding: utf-8 -*-
"""QA dos LOTES AUTOMATICOS e do aviso de VALIDADE (ADR-060) -- roda numa COPIA com ADR-060 instalado
SEM o --registrar (os lotes ja recebidos ainda fora do cadastro, como no primeiro dia).

Uso: python qa_lotes_auto.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Pelo caminho real: ATUALIZAR DADOS (nucleo do botao), escrita na celula da validade com eventos
ligados, e as CAIXAS DE MENSAGEM DE VERDADE -- uma thread acha o dialogo do Excel, le o texto e
clica Sim / Nao / OK como o usuario. Por fim fecha, reabre e faz login: o aviso tem de voltar.
Nada e salvo no arquivo testado (a reabertura usa uma copia propria).
"""
import datetime as dtm
import hashlib
import json
import os
import shutil
import sys
import threading
import time

import win32con
import win32gui
import win32process

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES  # noqa: E402

CFG = 'Configuração'
SENHA_ADM = 'QaAdmLotes#2026'


class Dialogos:
    """Responde as caixas de mensagem do Excel (pid) na ordem dada, guardando o texto de cada uma."""

    def __init__(self, pid, botoes, teto=60):
        self.pid, self.botoes, self.teto = pid, list(botoes), teto
        self.vistos = []
        self.th = threading.Thread(target=self._rodar, daemon=True)
        self.th.start()

    def _janelas(self):
        achou = []

        def cb(h, _):
            try:
                _, p = win32process.GetWindowThreadProcessId(h)
            except Exception:
                return True
            if p == self.pid and win32gui.GetClassName(h) == '#32770' and win32gui.IsWindowVisible(h):
                achou.append(h)
            return True
        win32gui.EnumWindows(cb, None)
        return achou

    def _rodar(self):
        t0 = time.time()
        while self.botoes and time.time() - t0 < self.teto:
            for h in self._janelas():
                textos, botoes = [], {}

                def cb2(c, _):
                    t = win32gui.GetWindowText(c)
                    if win32gui.GetClassName(c) == 'Button':
                        botoes[t.replace('&', '').strip().upper()] = c
                    elif t:
                        textos.append(t)
                    return True
                win32gui.EnumChildWindows(h, cb2, None)
                alvo = self.botoes[0].upper()
                if alvo in botoes:
                    self.vistos.append({'titulo': win32gui.GetWindowText(h), 'texto': '\n'.join(textos),
                                        'botoes': sorted(botoes), 'clicou': self.botoes[0]})
                    self.botoes.pop(0)
                    win32gui.PostMessage(botoes[alvo], win32con.BM_CLICK, 0, 0)
                    time.sleep(0.6)
                    break
            time.sleep(0.2)

    def esperar(self):
        self.th.join(self.teto + 5)
        return self.vistos


def sha(t):
    return hashlib.sha256(t.encode('latin-1')).hexdigest()


def cadastro(q):
    ws = q.ws(CFG)
    v = ws.Range('C26:E125').Value
    return [(26 + i, str(r[0]).strip() if r[0] not in (None, '') else '', r[1], r[2]) for i, r in enumerate(v)]


def esperados(q):
    """Lotes que o sistema DEVE registrar: interfaceamento + analito cadastrado, fora do cadastro, por 1o resultado."""
    cad = {c[1].upper() for c in cadastro(q) if c[1]}
    prim = {}
    for r in q.ler('tblCQ_Final'):
        if str(r.get('ORIGEM_RESULTADO')) == 'INTERFACEAMENTO' and str(r.get('ANALITO_CADASTRADO')) == 'SIM':
            lote = str(r.get('LOTE') or '').strip()
            d = r.get('DATA_HORA')
            if lote and d is not None and lote.upper() not in cad:
                prim[lote] = min(prim.get(lote, d), d)
    return [k for k, _ in sorted(prim.items(), key=lambda kv: kv[1])]


def eventos(q, desde, acao=None):
    al = q.ws('Audit_Log')
    ult = al.Cells(al.Rows.Count, 1).End(-4162).Row
    out = []
    for r in range(desde + 1, ult + 1):
        a = al.Cells(r, 7).Value
        if acao is None or a == acao:
            out.append({'acao': a, 'categoria': al.Cells(r, 6).Value, 'lote': al.Cells(r, 13).Value,
                        'antes': al.Cells(r, 17).Value, 'depois': al.Cells(r, 18).Value, 'parecer': al.Cells(r, 24).Value})
    return out


def marca(q):
    al = q.ws('Audit_Log')
    return al.Cells(al.Rows.Count, 1).End(-4162).Row


def criar_adm(q):
    us = q.ws('Usuarios')
    us.Unprotect('qcini2025')
    i = next(i for i in range(4, 54) if not us.Cells(i, 1).Value)
    us.Cells(i, 1).Value = 'QA_ADM_LOTES'
    us.Cells(i, 2).Value = 'QA'
    us.Cells(i, 3).Value = 'ADM'
    us.Cells(i, 4).Value = sha(SENHA_ADM)


def executar(produto, caminho, saida):
    q = QA(produto, caminho)
    try:
        q.run('mSeguranca.ReprotectAll')          # sessao como a do login
        ws = q.ws(CFG)
        # A entrega ja registra os lotes recebidos (instalar_adr060 --registrar). Para provar o
        # REGISTRO pelo ATUALIZAR DADOS, esta copia volta ao "primeiro dia": as linhas de origem
        # automatica (sempre as ultimas, anexadas) sao limpas -- so nesta copia, nunca salva.
        pre = [c for c in cadastro(q) if str(c[3] or '').startswith('automático em')]
        if pre:
            ws.Unprotect('qcini2025')
            q.ex.xl.EnableEvents = False
            for c in pre:
                ws.Range(f'C{c[0]}:E{c[0]}').ClearContents()
                ws.Range(f'AD{c[0]}').ClearContents()
            q.ex.xl.EnableEvents = True
            q.run('mSeguranca.ReprotectAll')
        reg('A00 A entrega chega com os lotes recebidos já registrados (origem automática), sem validade',
            True if not pre else all(c[2] in (None, '') for c in pre),
            {'lotes_ja_registrados_na_entrega': [c[1] for c in pre]})
        esp = esperados(q)
        ult0 = max([c[0] for c in cadastro(q) if c[1]] or [25])
        sem_val0 = str(q.run('mLotes.LotesSemValidade'))

        # ------------------------------------------------------------ A01: ATUALIZAR DADOS registra
        m = marca(q)
        resumo, t = q.atualizar()
        cad = cadastro(q)
        novos = [c for c in cad if c[0] > ult0 and c[1]]
        lotes_novos = [c[1] for c in novos]
        ev = eventos(q, m, 'LOTE_CADASTRADO')
        qa = q.ler('tblQA_Integracao')
        i03 = [x for x in qa if x['CODIGO'] == 'I03']
        reg('A01 ATUALIZAR DADOS registra sozinho os lotes recebidos (interfaceamento + analito cadastrado), '
            'anexados no fim, em ordem de chegada, sem validade, com origem; trilha CONFIGURACAO; QA sem I03',
            lotes_novos == esp and len(esp) > 0 and all(c[2] in (None, '') for c in novos)
            and all(str(c[3]).startswith('automático em') for c in novos)
            and [c[0] for c in novos] == list(range(ult0 + 1, ult0 + 1 + len(esp)))
            and len(ev) == len(esp) and all(e['categoria'] == 'CONFIGURACAO' and e['lote'] for e in ev)
            and not i03 and 'Lotes novos registrados automaticamente' in resumo,
            {'esperados': esp, 'registrados': lotes_novos, 'linhas': [c[0] for c in novos], 'eventos': len(ev),
             'I03_restantes': len(i03), 'origem_exemplo': novos[0][3] if novos else None, 'tempo_atualizar_s': round(t, 1),
             'linha_resumo': [l for l in resumo.split('\n') if 'Lotes novos' in l]})

        # ------------------------------------------------------------ A02: idempotente e rapido
        t0 = time.perf_counter()
        r2 = str(q.run('mLotes.RegistrarLotesRecebidos'))
        dt = time.perf_counter() - t0
        reg('A02 Segunda passada não registra nada (idempotente) e é rápida',
            r2 == 'OK|0||' and dt < 2.0, {'retorno': r2, 'tempo_s': round(dt, 3), 'linhas_tblCQ_Final': len(q.ler('tblCQ_Final'))})

        # ------------------------------------------------------------ A03: manual NAO vira lote
        q.escrever_linha('tblResultados_Manuais', {'DATA': dtm.datetime(2026, 10, 1), 'HORA': '03:07', 'LOTE': 'QA7777', 'NIVEL': 1,
                                                   'ANALITO': q.ws('Painel').Range('C3').Value, 'RESULTADO': 1})
        q.atualizar()
        reg('A03 Lote digitado em resultado MANUAL (QA7777) não entra no cadastro (lote fantasma)',
            'QA7777' not in [c[1] for c in cadastro(q)], {'cadastro_tem_QA7777': 'QA7777' in [c[1] for c in cadastro(q)]})

        # ------------------------------------------------------------ A04: pendencias
        sv = str(q.run('mLotes.LotesSemValidade'))
        n_sv, lista_sv = int(sv.split('|')[0]), sv.split('|')[1].split(';') if sv.split('|')[1] else []
        reg('A04 Lotes sem validade = novos + os que já estavam sem validade',
            set(esp) <= set(lista_sv) and n_sv == len(lista_sv), {'antes': sem_val0, 'agora': sv})

        # ------------------------------------------------------------ A05/A06: digitar a validade (evento)
        alvo = novos[0]
        m = marca(q)
        ws.Range(f'D{alvo[0]}').Value = dtm.datetime(2026, 9, 30)   # o usuario digita a data
        sv2 = str(q.run('mLotes.LotesSemValidade'))
        ev = eventos(q, m, 'VALIDADE_LOTE_ALTERADA')
        sombra = ws.Range(f'AD{alvo[0]}').Value
        reg('A05 Validade digitada: sai da lista de pendentes, fica na trilha com antes/depois, sombra atualizada',
            int(sv2.split('|')[0]) == n_sv - 1 and ev and ev[-1]['lote'] == alvo[1] and ev[-1]['depois'] == '30/09/2026'
            and sombra is not None and ws.Range(f'D{alvo[0]}').Text == '30/09/2026',
            {'pendentes': sv2.split('|')[0], 'evento': ev[-1:] if ev else None, 'texto_celula': ws.Range(f'D{alvo[0]}').Text})
        m = marca(q)
        ws.Range(f'D{novos[1][0] if len(novos) > 1 else alvo[0]}').Value = 'amanha'
        cel = ws.Range(f'D{novos[1][0] if len(novos) > 1 else alvo[0]}')
        reg('A06 Texto colado na validade (passa por cima da validação de dados) é recusado e desfeito, sem trilha',
            str(q.run('mLotes.LotesSemValidade')) == sv2 and cel.Text != 'amanha' and not eventos(q, m, 'VALIDADE_LOTE_ALTERADA'),
            {'celula_depois': cel.Text})

        # ------------------------------------------------------------ A07: automacao nunca abre dialogo
        t0 = time.perf_counter()
        q.run('mLotes.AvisarLotesSemValidade')
        reg('A07 Em automação (Interactive=False) o aviso não abre caixa nenhuma (não trava o Excel)',
            time.perf_counter() - t0 < 5, {'tempo_s': round(time.perf_counter() - t0, 2)})

        # ------------------------------------------------------------ A08: caixa real -> SIM
        q.ex.xl.Interactive = True
        d = Dialogos(q.ex.pid, ['Sim'])
        q.run('mLotes.AvisarLotesSemValidade')
        vistos = d.esperar()
        q.ex.xl.Interactive = False
        ativa = q.ex.xl.ActiveSheet.Name
        cel_ativa = q.ex.xl.ActiveCell.Address
        lin_ativa = q.ex.xl.ActiveCell.Row
        prim_pend = next((c[0] for c in cadastro(q) if c[1] and not ws.Range(f'D{c[0]}').Value), None)
        txt = vistos[0]['texto'] if vistos else ''
        reg('A08 Caixa de verdade: "detectei lote SEM DATA DE VALIDADE ... RECOMENDADO: vamos inserir agora?" '
            '-> Sim leva à célula da validade do 1º lote pendente',
            vistos and 'SEM DATA DE VALIDADE' in txt and 'RECOMENDADO: vamos inserir agora?' in txt
            and 'NÃO é recomendado' in txt and ativa == CFG and cel_ativa == f'$D${prim_pend}',
            {'dialogo': vistos, 'aba_ativa': ativa, 'celula_ativa': cel_ativa, 'primeira_pendente': prim_pend})

        # ------------------------------------------------------------ A09: caixa real -> NAO
        m = marca(q)
        q.ex.xl.Interactive = True
        d = Dialogos(q.ex.pid, ['Não', 'OK'])
        q.run('mLotes.AvisarLotesSemValidade')
        vistos = d.esperar()
        q.ex.xl.Interactive = False
        ev = eventos(q, m, 'VALIDADE_LOTE_ADIADA')
        seg = vistos[1]['texto'] if len(vistos) > 1 else ''
        reg('A09 Não -> segunda caixa: "NÃO é recomendado ... vai aparecer de novo toda vez que o arquivo for aberto"; '
            'adiamento na trilha; pendências continuam',
            len(vistos) == 2 and 'NÃO é recomendado' in seg and 'toda vez que o arquivo for aberto' in seg and ev
            and str(q.run('mLotes.LotesSemValidade')) == sv2,
            {'dialogos': vistos, 'evento': ev[-1:] if ev else None})

        # ------------------------------------------------------------ A10: sem buraco, cadastro cheio
        snap = ws.Range('C26:E125').Value
        ws.Unprotect('qcini2025')
        q.ex.xl.EnableEvents = False
        a, b = novos[-1], novos[-2] if len(novos) > 1 else None
        ws.Range(f'C{a[0]}').Value = ''                          # dois lotes "somem" do cadastro (buracos)
        if b:
            ws.Range(f'C{b[0]}').Value = ''
        livre = [r for r in range(26, 126) if not ws.Range(f'C{r}').Value]
        enchidas = []
        for r in livre:
            if r < 125 and r > a[0]:
                ws.Range(f'C{r}').Value = f'QA{r:04d}'
                enchidas.append(r)
        q.ex.xl.EnableEvents = True
        q.run('mSeguranca.ReprotectAll')
        r10 = str(q.run('mLotes.RegistrarLotesRecebidos'))
        c125 = str(ws.Range('C125').Value or '')
        buraco_reusado = bool(ws.Range(f'C{a[0]}').Value) or (b and bool(ws.Range(f'C{b[0]}').Value))
        p = r10.split('|')
        reg('A10 Lote novo é anexado no FIM (nunca reaproveita buraco: a posição endereça a Liberação); '
            'cadastro cheio registra o que cabe e informa o resto, sem erro',
            p[0] == 'OK' and p[1] == '1' and c125 and not buraco_reusado and len(p) > 3 and (p[3] != '' if b else True),
            {'retorno': r10, 'C125': c125, 'buraco_reusado': buraco_reusado})
        # desfaz (so nesta copia)
        ws.Unprotect('qcini2025')
        q.ex.xl.EnableEvents = False
        ws.Range('C26:E125').Value = snap
        q.ex.xl.EnableEvents = True
        q.run('mSeguranca.ReprotectAll')

        # ------------------------------------------------------------ A11: C21 enxerga os 100 lotes
        ws.Unprotect('qcini2025')
        q.ex.xl.EnableEvents = False
        c20 = ws.Range('C20').Value
        ws.Range('C60').NumberFormat = '@'
        ws.Range('C60').Value = 'QA6060'
        ws.Range('D60').Value = dtm.datetime(2027, 8, 15)
        ws.Range('C20').Value = 'QA6060'
        q.ex.xl.Calculate()
        c21 = ws.Range('C21').Text
        ws.Range('C20').Value = c20
        ws.Range('C60:D60').ClearContents()
        q.ex.xl.EnableEvents = True
        q.run('mSeguranca.ReprotectAll')
        reg('A11 Validade do lote em uso (C21) enxerga as 100 linhas do cadastro (a Bioquímica olhava só 25)',
            c21 == '15/08/2027', {'C21': c21})

        # ------------------------------------------------------------ A12: travas da Configuracao
        trav = {a_: ws.Range(a_).Locked for a_ in ('B26', 'C26', 'C21', 'AB2', 'E26', 'C9', 'D26', 'D125', 'C20')}
        reg('A12 Configuração travada; editáveis só identificação/constantes (C5:C16), lote em uso (C20) e validade (D26:D125)',
            trav['B26'] and trav['C26'] and trav['C21'] and trav['AB2'] and trav['E26']
            and trav['C9'] is False and trav['D26'] is False and trav['D125'] is False and trav['C20'] is False
            and bool(ws.ProtectContents) and ws.Columns(30).Hidden,
            {'locked': trav, 'protegida': bool(ws.ProtectContents), 'sombra_oculta': ws.Columns(30).Hidden})

        criar_adm(q)
        q.ws('Usuarios').Protect('qcini2025', False, True, True, True)
    finally:
        # copia para a reabertura ANTES de fechar sem salvar: o estado (lotes registrados, ADM de teste) vai junto
        copia = caminho.replace('.xlsm', '_reabrir_lotes.xlsm')
        try:
            q.wb.SaveCopyAs(copia)
        except Exception:
            copia = None
        q.fechar(salvar=False)

    # ---------------------------------------------------------------- A13: fechar, abrir, login -> aviso volta
    if copia:
        q2 = QA(produto, copia)
        try:
            q2.run('mSeguranca.LockApp')            # o que o Workbook_Open faz
            q2.nome('loginUser').Value = 'QA_ADM_LOTES'
            q2.nome('loginPass').Value = SENHA_ADM
            q2.ex.xl.Interactive = True
            d = Dialogos(q2.ex.pid, ['Não', 'OK'])
            q2.run('mSeguranca.DoLogin')
            vistos = d.esperar()
            q2.ex.xl.Interactive = False
            sess = str(q2.nome('currentUser').Value or '')
            reg('A13 Fechou sem registrar a validade, abriu de novo, fez login: o MESMO aviso aparece (e o "Não" repete o alerta)',
                sess == 'QA_ADM_LOTES' and len(vistos) == 2 and 'SEM DATA DE VALIDADE' in vistos[0]['texto']
                and 'RECOMENDADO: vamos inserir agora?' in vistos[0]['texto'] and 'NÃO é recomendado' in vistos[1]['texto'],
                {'sessao': sess, 'dialogos': vistos})
        finally:
            q2.fechar(salvar=False)
            try:
                os.remove(copia)
            except Exception:
                pass

    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} lotes automáticos: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
