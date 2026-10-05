# -*- coding: utf-8 -*-
"""QA de SEGURANCA (ADR-059 + QUALITY_GATE 3.3/3.5) -- roda numa COPIA ja instalada e blindada.

Uso: python qa_seguranca.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Exercita o caminho real: a aba Login (loginUser/loginPass + DoLogin), a area de cadastro
da aba Usuarios (cadLogin/cadNome/cadSenha/cadPapel + CadastrarUsuario), Logout e o nucleo
do Modo Desenvolvedor. Os usuarios de teste (QA_*) sao criados so nesta copia, com hash
calculado aqui (SHA-256 da senha ASCII, o mesmo que o SHA256Hex do VBA). Nada e salvo.

A pergunta de cada teste e a do achado CRITICO da FASE3A: alguem consegue se promover,
assumir a identidade de outro ou agir sem deixar rastro?
"""
import hashlib
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import blindar_entrega  # noqa: E402

RES = []
SENHA_PROT = 'qcini2025'
SENHAS = {'QA_ADM': 'QaAdm#2026x', 'QA_ANA': 'QaAna#2026y', 'QA_TEC': 'QaTec#2026z'}
NOVAS = {'QA_TEC2': 'QaTec2#2026w', 'QA_ANA_NOVA': 'QaAnaNova#2026v', 'QA_ADM_RESET': 'QaReset#2026u',
         'QA_ADM_PROPRIA': 'QaAdmPropria#2026t', 'QA_FORJA': 'QaForja#2026s'}


def reg(teste, ok, evid):
    RES.append({'teste': teste, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': evid})
    print(('PASS ' if ok else 'FAIL ') + teste + ' -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:500], flush=True)


def sha(t):
    return hashlib.sha256(t.encode('latin-1')).hexdigest()


class Seg:
    def __init__(self, caminho):
        self.ex = xlh.Excel()
        self.wb = self.ex.abrir(caminho)
        self.us = self.wb.Worksheets('Usuarios')

    def run(self, macro, *a):
        return self.ex.run("'" + self.wb.Name + "'!" + macro, *a, teto=300)

    def nome(self, n):
        return self.wb.Names(n).RefersToRange

    def val(self, n):
        v = self.nome(n).Value
        return '' if v is None else str(v)

    def linha(self, login):
        for i in range(4, 54):
            if str(self.us.Cells(i, 1).Value or '').strip().upper() == login.upper():
                return i
        return 0

    def usuario(self, login):
        i = self.linha(login)
        if not i:
            return None
        return {'papel': self.us.Cells(i, 3).Value, 'hash': str(self.us.Cells(i, 4).Value or '').lower()}

    def login(self, u, p):
        self.nome('loginUser').Value = u
        self.nome('loginPass').Value = p
        self.run('mSeguranca.DoLogin')
        return self.val('currentUser'), self.val('currentPapel'), self.val('loginMsg')

    def cadastrar(self, login, senha, papel='', nome_='QA'):
        self.nome('cadLogin').Value = login
        self.nome('cadNome').Value = nome_
        self.nome('cadSenha').Value = senha
        self.nome('cadPapel').Value = papel
        self.run('mSeguranca.CadastrarUsuario')
        return self.val('cadMsg')

    def log_linhas(self):
        al = self.wb.Worksheets('Audit_Log')
        ult = al.Cells(al.Rows.Count, 1).End(-4162).Row
        return ult

    def eventos(self, desde):
        al = self.wb.Worksheets('Audit_Log')
        ate = self.log_linhas()
        out = []
        for r in range(desde + 1, ate + 1):
            out.append({'acao': al.Cells(r, 7).Value, 'categoria': al.Cells(r, 6).Value, 'usuario': al.Cells(r, 25).Value,
                        'papel_sessao': al.Cells(r, 26).Value, 'st_ant': al.Cells(r, 21).Value, 'st_novo': al.Cells(r, 22).Value,
                        'parecer': al.Cells(r, 24).Value})
        return out

    def texto_log(self, desde):
        al = self.wb.Worksheets('Audit_Log')
        ate = self.log_linhas()
        v = al.Range(al.Cells(desde + 1, 1), al.Cells(ate, 33)).Value
        if not isinstance(v, tuple):
            v = ((v,),)
        return ' '.join(str(c) for linha in v for c in linha if c is not None)

    def fechar(self):
        self.ex.fechar(salvar=False)


def executar(produto, caminho, saida):
    # ------------------------------------------------------------ S01: o arquivo SALVO, lido como zip
    reg('S01 Arquivo salvo travado (XML): só Login visível, toda aba com <sheetProtection>, estrutura travada',
        blindar_entrega.verificar(caminho), {'arquivo': os.path.basename(caminho)})

    q = Seg(caminho)
    try:
        us = q.us
        # ------------------------------------------------------------ S02: travas da aba Usuarios
        travas = {a: us.Range(a).Locked for a in ('N1:N2', 'A4:E53')}
        cad = {n: q.nome(n).Locked for n in ('cadLogin', 'cadNome', 'cadSenha', 'cadPapel')}
        reg('S02 Usuarios: sessão (N1:N2) e tabela (A4:E53) travadas; só o cadastro editável; hash oculto; sessão vazia no arquivo',
            all(v is True for v in travas.values()) and all(v is False for v in cad.values()) and us.Columns(4).Hidden
            and q.val('currentUser') == '' and q.val('currentPapel') == '',
            {'travas': travas, 'cadastro_destravado': cad, 'hash_oculto': us.Columns(4).Hidden,
             'sessao': [q.val('currentUser'), q.val('currentPapel')]})

        # ------------------------------------------------------------ preparo: usuarios QA_* (so nesta copia)
        try:
            q.wb.Unprotect(SENHA_PROT)
        except Exception:
            pass
        us.Unprotect(SENHA_PROT)
        livres = [i for i in range(4, 54) if not us.Cells(i, 1).Value][:3]
        for i, (lg, pp) in zip(livres, (('QA_ADM', 'ADM'), ('QA_ANA', 'ANALISTA'), ('QA_TEC', 'TÉCNICO'))):
            us.Cells(i, 1).Value = lg
            us.Cells(i, 2).Value = 'QA ' + pp
            us.Cells(i, 3).Value = pp
            us.Cells(i, 4).Value = sha(SENHAS[lg])
        q.run('mSeguranca.LockApp')        # o mesmo que Workbook_Open: tela de login, tudo protegido
        marca0 = q.log_linhas()

        # ------------------------------------------------------------ S03: login falho
        u, p, msg = q.login('QA_ANA', 'senha-errada')
        ev = [e for e in q.eventos(marca0) if e['acao'] == 'LOGIN_FALHOU']
        reg('S03 Senha errada: não entra, sessão vazia, LOGIN_FALHOU na trilha (com o login tentado, sem a senha)',
            u == '' and p == '' and 'invalidos' in msg and ev and 'QA_ANA' in str(ev[-1]['parecer']),
            {'sessao': [u, p], 'msg': msg, 'evento': ev[-1:] if ev else None})

        # ------------------------------------------------------------ S04: login ANALISTA
        m = q.log_linhas()
        u, p, msg = q.login('QA_ANA', SENHAS['QA_ANA'])
        ev = [e for e in q.eventos(m) if e['acao'] == 'LOGIN']
        calc_oculto = q.wb.Worksheets('Calc').Visible == 2
        audit_filtra = q.wb.Worksheets('Audit_Log').Protection.AllowFiltering
        reg('S04 Login ANALISTA: sessão = tabela (QA_ANA/ANALISTA), LOGIN na trilha, Calc oculta, Audit_Log filtrável',
            u == 'QA_ANA' and p == 'ANALISTA' and ev and ev[-1]['usuario'] == 'QA_ANA' and calc_oculto and audit_filtra,
            {'sessao': [u, p], 'evento': ev[-1:] if ev else None, 'calc_veryHidden': calc_oculto,
             'audit_AllowFiltering': audit_filtra})

        # ------------------------------------------------------------ S05..S09: SOMENTE O ADM (03/10/2026)
        hashes0 = {i: str(us.Cells(i, 4).Value or '').lower() for i in range(4, 54)}
        h_adm = q.usuario('QA_ADM')['hash']
        m = q.log_linhas()
        msg5 = q.cadastrar('QA_NOVO_ADM', 'x1#Abcdef', 'ADM')
        msg6 = q.cadastrar('QA_ADM', NOVAS['QA_ADM_RESET'], 'ADM')
        msg7 = q.cadastrar('QA_ANA', SENHAS['QA_ANA'], 'ADM')
        msg8 = q.cadastrar('QA_TEC2', NOVAS['QA_TEC2'], 'Tecnico')
        msg9 = q.cadastrar('QA_ANA', NOVAS['QA_ANA_NOVA'], '')
        rec = [e for e in q.eventos(m) if e['acao'] == 'CADASTRO_RECUSADO']
        so_adm = 'Somente o ADM'
        reg('S05 ANALISTA cria um ADM: recusado, ninguém criado, recusa na trilha',
            q.usuario('QA_NOVO_ADM') is None and so_adm in msg5 and len(rec) >= 1, {'msg': msg5})
        reg('S06 ANALISTA regrava a senha do ADM: recusado, hash do ADM intacto (achado CRÍTICO da FASE3A)',
            q.usuario('QA_ADM')['hash'] == h_adm and so_adm in msg6 and len(rec) >= 2, {'msg': msg6})
        reg('S07 ANALISTA muda a própria função para ADM: recusado, continua ANALISTA',
            q.usuario('QA_ANA')['papel'] == 'ANALISTA' and so_adm in msg7 and len(rec) >= 3, {'msg': msg7})
        reg('S08 ANALISTA não cria usuário (definir a senha inicial é definir senha): recusado, ninguém criado, formulário sem senha',
            q.usuario('QA_TEC2') is None and so_adm in msg8 and q.val('cadSenha') == '' and len(rec) >= 4, {'msg': msg8})
        reg('S09 ANALISTA não troca nem a própria senha: recusado, hash intacto, recusa na trilha',
            q.usuario('QA_ANA')['hash'] == sha(SENHAS['QA_ANA']) and so_adm in msg9 and len(rec) >= 5, {'msg': msg9})

        # ------------------------------------------------------------ S10: sessao forjada (celula e login da sessao)
        # o que barra a DIGITACAO e celula travada em aba protegida. (Escrita por COM nao
        # serve de prova: com UserInterfaceOnly a protecao deixa passar automacao.)
        bloqueado = q.nome('currentPapel').Locked is True and bool(us.ProtectContents)
        # defesa em profundidade 1: papel forcado na celula -> o papel vem da tabela
        us.Unprotect(SENHA_PROT)
        q.nome('currentPapel').Value = 'ADM'
        q.run('mSeguranca.ReprotectAll')
        msg10a = q.cadastrar('QA_FORJADO', 'x2#Abcdef', 'ADM')
        # defesa em profundidade 2: LOGIN da sessao forcado para o ADM -> a ancora em memoria e do DoLogin
        h_ana = q.usuario('QA_ANA')['hash']
        us.Unprotect(SENHA_PROT)
        q.nome('currentUser').Value = 'QA_ADM'
        q.nome('currentPapel').Value = 'ADM'
        q.run('mSeguranca.ReprotectAll')
        msg10b = q.cadastrar('QA_ANA', NOVAS['QA_FORJA'], 'ANALISTA')
        reg('S10 Sessão forjada: célula travada; papel forçado → papel vem da tabela; login forçado para o ADM → "sessão inválida", hash intacto',
            bloqueado and q.usuario('QA_FORJADO') is None and so_adm in msg10a
            and 'Sessao invalida' in msg10b and q.usuario('QA_ANA')['hash'] == h_ana,
            {'celula_bloqueada': bloqueado, 'msg_papel': msg10a, 'msg_login': msg10b})

        # ------------------------------------------------------------ S11/S12: logout, TECNICO e Alt+F8
        m = q.log_linhas()
        q.run('mSeguranca.Logout')
        sess = [q.val('currentUser'), q.val('currentPapel')]
        ev = [e for e in q.eventos(m) if e['acao'] == 'LOGOUT']
        # sem login, UnlockApp pela lista de macros (Alt+F8) nao revela nada
        q.run('mSeguranca.UnlockApp')
        visiveis = [ws.Name for ws in q.wb.Worksheets if ws.Visible == -1]
        reg('S11 Logout limpa a sessão (LOGOUT na trilha); UnlockApp sem login (Alt+F8) não revela nenhuma aba',
            sess == ['', ''] and ev and visiveis == ['Login'], {'sessao': sess, 'visiveis_apos_UnlockApp': visiveis})
        q.login('QA_TEC', SENHAS['QA_TEC'])
        h_tec = q.usuario('QA_TEC')['hash']
        msg12a = q.cadastrar('QA_DO_TECNICO', 'x3#Abcdef', 'TÉCNICO')
        msg12b = q.cadastrar('QA_TEC', 'x4#Abcdef', '')
        q.run('mSeguranca.UnprotectAll')                 # Alt+F8 de um TECNICO
        ainda = bool(us.ProtectContents) and bool(q.wb.Worksheets('Configuração').ProtectContents)
        hashes1 = {i: str(us.Cells(i, 4).Value or '').lower() for i in range(4, 54)}
        reg('S12 TÉCNICO: não cria usuário, não troca a própria senha, UnprotectAll (Alt+F8) não desprotege nada; '
            'nenhum hash mudou sob sessão TÉCNICO/ANALISTA',
            q.usuario('QA_DO_TECNICO') is None and so_adm in msg12a and so_adm in msg12b
            and q.usuario('QA_TEC')['hash'] == h_tec and ainda and hashes1 == hashes0,
            {'msg_criar': msg12a, 'msg_propria': msg12b, 'abas_continuam_protegidas': ainda,
             'hashes_inalterados': hashes1 == hashes0})
        q.run('mSeguranca.Logout')

        # ------------------------------------------------------------ S13/S14: ADM
        u, p, _ = q.login('QA_ADM', SENHAS['QA_ADM'])
        m = q.log_linhas()
        msg13a = q.cadastrar('QA_TEC2', NOVAS['QA_TEC2'], 'Tecnico')
        t2 = q.usuario('QA_TEC2')
        msg13b = q.cadastrar('QA_TEC2', NOVAS['QA_ADM_RESET'], 'ANALISTA')
        msg13c = q.cadastrar('QA_ADM', NOVAS['QA_ADM_PROPRIA'], '')
        evs = q.eventos(m)
        criado = [e for e in evs if e['acao'] == 'USUARIO_CRIADO']
        alterado = [e for e in evs if e['acao'] == 'USUARIO_ALTERADO']
        propria = [e for e in evs if e['acao'] == 'SENHA_PROPRIA_ALTERADA']
        reg('S13 ADM cria TÉCNICO ("Tecnico" normalizado), redefine senha e função de outro e troca a própria: tudo aceito e na trilha',
            p == 'ADM' and t2 is not None and t2['papel'] == 'TÉCNICO' and t2['hash'] == sha(NOVAS['QA_TEC2'])
            and q.usuario('QA_TEC2')['papel'] == 'ANALISTA' and q.usuario('QA_TEC2')['hash'] == sha(NOVAS['QA_ADM_RESET'])
            and q.usuario('QA_ADM')['hash'] == sha(NOVAS['QA_ADM_PROPRIA']) and q.val('cadSenha') == ''
            and criado and criado[-1]['st_novo'] == 'TÉCNICO'
            and alterado and alterado[-1]['st_ant'] == 'TÉCNICO' and alterado[-1]['st_novo'] == 'ANALISTA' and propria,
            {'msgs': [msg13a, msg13b, msg13c]})
        # unico ADM: nesta copia, os outros ADM viram ANALISTA (nada e salvo)
        us.Unprotect(SENHA_PROT)
        outros = []
        for i in range(4, 54):
            if us.Cells(i, 1).Value and str(us.Cells(i, 1).Value).upper() != 'QA_ADM' and str(us.Cells(i, 3).Value).upper() == 'ADM':
                us.Cells(i, 3).Value = 'ANALISTA'
                outros.append(i)
        q.run('mSeguranca.ReprotectAll')
        msg14 = q.cadastrar('QA_ADM', NOVAS['QA_ADM_PROPRIA'], 'ANALISTA')
        reg('S14 Último ADM não se rebaixa (o sistema ficaria sem administrador)',
            q.usuario('QA_ADM')['papel'] == 'ADM' and 'unico ADM' in msg14, {'msg': msg14, 'outros_adm_na_copia': len(outros)})

        # ------------------------------------------------------------ S15: Modo Desenvolvedor
        m = q.log_linhas()
        r15 = q.run('mSeguranca.AtivarModoDesenvolvedor', 'nao-e-a-senha')
        ev = [e for e in q.eventos(m) if e['acao'] == 'MODO_DESENVOLVEDOR_RECUSADO']
        reg('S15 Modo Desenvolvedor com senha errada: não abre e deixa marca na trilha',
            r15 is False and bool(ev),
            {'retorno': r15, 'evento': ev[-1:] if ev else None})

        # ------------------------------------------------------------ S16/S17: trilha
        txt = q.texto_log(marca0)
        vazou = [s for s in list(SENHAS.values()) + list(NOVAS.values()) if s in txt or sha(s) in txt.lower()]
        seg = [e['categoria'] for e in q.eventos(marca0)]
        reg('S16 Nenhuma senha nem hash de senha na trilha; todos os eventos de acesso na categoria SEGURANCA',
            not vazou and seg and all(c == 'SEGURANCA' for c in seg), {'vazamentos': vazou, 'eventos': len(seg)})
        integ = str(q.run('mAuditoria.VerificarIntegridadeLog'))
        reg('S17 Cadeia de hash da Audit_Log íntegra depois de todos os eventos de segurança', integ.startswith('OK|'),
            {'verificacao': integ})
    finally:
        q.fechar()
    gravar_em_sessao(caminho)
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} segurança: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


def gravar_em_sessao(caminho):
    """S18-S20 (auditoria 04/10/2026): Ctrl+S no meio de uma sessao ADM. Numa copia DESCARTAVEL -- o arquivo
    testado nao e regravado (a bateria roda sobre os bytes que vao para a producao)."""
    import shutil
    import tempfile
    pasta = tempfile.mkdtemp(prefix='qa_grava_')
    tmp = os.path.join(pasta, os.path.basename(caminho))
    shutil.copy2(caminho, tmp)
    q = Seg(tmp)
    try:
        us = q.us
        try:
            q.wb.Unprotect(SENHA_PROT)
        except Exception:
            pass
        us.Unprotect(SENHA_PROT)
        i = [i for i in range(4, 54) if not us.Cells(i, 1).Value][0]
        us.Cells(i, 1).Value = 'QA_ADM'
        us.Cells(i, 2).Value = 'QA ADM'
        us.Cells(i, 3).Value = 'ADM'
        us.Cells(i, 4).Value = sha(SENHAS['QA_ADM'])
        q.run('mSeguranca.LockApp')
        sess0 = q.run('mSeguranca.SessaoAtiva')
        u, p, _ = q.login('QA_ADM', SENHAS['QA_ADM'])
        sess1 = q.run('mSeguranca.SessaoAtiva')
        audit_prot = bool(q.wb.Worksheets('Audit_Log').ProtectContents)
        cfg_aberta = not q.wb.Worksheets('Configuração').ProtectContents
        reg('S18 Sessão ADM: abas técnicas desprotegidas, mas a trilha (Audit_Log) continua protegida -- antes o ADM a '
            'editava sem rastro; SessaoAtiva (guarda das macros que alteram dados, ex.: NovoLote) só depois do login',
            p == 'ADM' and audit_prot and cfg_aberta and sess0 is False and sess1 is True,
            {'papel': p, 'audit_protegida': audit_prot, 'configuracao_aberta': cfg_aberta,
             'sessao_antes_depois_login': [sess0, sess1]})
        q.wb.Worksheets('Painel').Activate()
        q.ex.xl.EnableEvents = True                  # o Ctrl+S do usuario: BeforeSave/AfterSave rodam
        q.wb.Save()
        q.ex.xl.EnableEvents = False
        vis = [ws.Name for ws in q.wb.Worksheets if ws.Visible == -1]
        ativa = q.wb.ActiveSheet.Name
        sess2 = q.run('mSeguranca.SessaoAtiva')
        reg('S19 Depois do Ctrl+S a sessão continua como estava: mesma aba ativa, abas do ADM visíveis e abertas, '
            'sessão válida, pasta marcada como salva',
            q.val('currentUser') == 'QA_ADM' and 'Painel' in vis and 'Login' not in vis and ativa == 'Painel'
            and bool(q.wb.Saved) and sess2 is True and not q.wb.Worksheets('Configuração').ProtectContents,
            {'aba_ativa': ativa, 'visiveis': len(vis), 'sessao': q.val('currentUser'), 'saved': q.wb.Saved,
             'sessao_ativa': sess2})
    finally:
        q.fechar()
    trav = blindar_entrega.verificar(tmp)
    r = Seg(tmp)                                     # sem eventos: o Workbook_Open nao limpa nada
    try:
        gravado = [r.val('currentUser'), r.val('currentPapel')]
    finally:
        r.fechar()
    shutil.rmtree(pasta, ignore_errors=True)
    reg('S20 O arquivo gravado NO MEIO da sessão ADM está TRAVADO no disco (XML: só Login visível, toda aba com '
        '<sheetProtection>, estrutura travada) e sem a identidade da sessão -- com macros desabilitadas não abre nada',
        trav and gravado == ['', ''], {'xml_travado': trav, 'sessao_gravada': gravado})


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
