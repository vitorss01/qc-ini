# -*- coding: utf-8 -*-
"""QA do ADR-084 -- cadastro central de profissionais, PBKDF2, login e assinatura nos QCs. Roda em COPIAS: o QC instalado
e uma copia do cadastro central instalado (instalar_cadastro_central.py) numa pasta de teste. Usuarios de TESTE (QA_*);
a unica linha real usada e a do ADMIN da copia, para o primeiro acesso -- o login dele nunca e impresso. Nada e salvo no
QC; a copia do central e salva (o QC le o arquivo), e e descartavel.

Uso: python qa_usuarios_adr084.py <Bioquimica|Hematologia> <copia_QC.xlsm> <copia_central.xlsm> <saida.json>

  S01 primeiro acesso: o ADMIN da tabela define a propria senha (sem outro ADMIN com senha); hash pbkdf2_sha256 que o
      hashlib do Python confere; depois disso, o primeiro acesso nao vale mais
  S02 recusas sem gravar: senhas diferentes, senha fraca, cargo fora da lista, ADMIN com senha errada, nao-ADMIN autorizando
  S03 cadastro de TECNICO e ANALISTA de teste pelo ADMIN: linha nova (E vazia, G ATIVO); login repetido (outra caixa)
      NAO cria linha (vira redefinicao de senha); nenhuma senha em texto em celula alguma do arquivo nem no LOG_CADASTRO
  S04 QC: Sincronizar le o central (copia Usuarios_Central = o central, papel pelo cargo, hash igual)
  S05 login de usuario central: senha certa entra com o papel do cargo; errada nao entra
  S06 usuario INATIVO no central: nao entra nem assina
  S07 central indisponivel: ERRO, a copia anterior fica, o login continua funcionando, Audit_Log registra
  S08 usuario LOCAL com hash SHA-256 antigo: entra e o hash vira PBKDF2 (SENHA_MIGRADA_PBKDF2); entra de novo
  S09 assinatura (Liberação) de usuario central com imagem: texto com nome, cargo, conselho, data e login; imagem
      qcAssin_ na celula; Audit_Log com a identidade. Outro usuario sem imagem na linha seguinte: so texto, nenhuma
      imagem nessa linha
  S10 FINALIZADO por TECNICO: recusado; senha errada: recusada
  S11 o cadastro muda depois (nome): a liberacao ja feita continua com o nome da hora da assinatura
  S00 arquivo QC intacto
"""
import hashlib
import json
import os
import shutil
import struct
import sys
import time
import traceback
import zlib

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

SENHA_ADM = 'QaAdmin#2026x'                 # valores de TESTE: so existem nas copias deste teste
SENHA_TEC = 'QaTecnico#2026'
SENHA_ANA = 'QaAnalista#2026'
SENHA_LOC = 'QaLocal#2026'
TEC, ANA, LOC = 'QA_TECNICO_084', 'QA_ANALISTA_084', 'QA_LOCAL_084'


def png(caminho, cor=(31, 78, 121), w=60, h=18):
    raw = b''.join(b'\x00' + bytes(cor) * w for _ in range(h))

    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    with open(caminho, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
                + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def pbkdf2_ok(senha, h):
    p = str(h).split('$')
    return (len(p) == 4 and p[0] == 'pbkdf2_sha256' and
            hashlib.pbkdf2_hmac('sha256', senha.encode('utf-8'), bytes.fromhex(p[2]), int(p[1])).hex() == p[3])


def executar(produto, caminho, central, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    xl = q.ex.xl
    wc = None
    try:
        wc = xl.Workbooks.Open(central, 0)
        nc = wc.Name

        def cen(macro, *a):
            return str(q.ex.run(f"'{nc}'!{macro}", *a, teto=300))

        def qc(macro, *a, teto=300):
            return str(q.run(macro, *a, teto=teto))
        lo = wc.Worksheets('REGISTRO_DE_USUÁRIOS').ListObjects('tb_registro_usuários')

        def linhas():
            return [list(r) for r in lo.DataBodyRange.Value]
        adm = next(r[1] for r in linhas() if str(r[2]).upper() == 'ADMIN')
        n0 = len(linhas())

        # ------------------------------------------------------------ S01 primeiro acesso
        r_out = cen('mCadastro.CadastrarOuRedefinir', '', adm, '', SENHA_ADM, SENHA_ADM, '', TEC, 'x')
        r1 = cen('mCadastro.CadastrarOuRedefinir', '', adm, '', SENHA_ADM, SENHA_ADM, '', adm, '')
        h_adm = next(r[3] for r in linhas() if r[1] == adm)
        r1b = cen('mCadastro.CadastrarOuRedefinir', '', adm, '', SENHA_ADM + 'y', SENHA_ADM + 'y', '', adm, '')
        reg('S01 Primeiro acesso: só o próprio ADMIN define a 1ª senha; hash pbkdf2_sha256 conferido pelo hashlib; '
            'depois disso o primeiro acesso não vale mais',
            r_out.startswith('ERRO|') and r1.startswith('OK|') and pbkdf2_ok(SENHA_ADM, h_adm) and r1b.startswith('ERRO|'),
            {'outro_antes': r_out, 'primeiro': r1, 'formato': str(h_adm).split('$')[:2], 'de_novo_sem_senha': r1b})

        # ------------------------------------------------------------ S02 recusas
        cargo_t = 'Técnico de Laboratório'
        rec = {
            'senhas_diferentes': cen('mCadastro.CadastrarOuRedefinir', 'QA Técnico', TEC, cargo_t, SENHA_TEC, SENHA_TEC + '1', 'CRBM 0', adm, SENHA_ADM),
            'senha_fraca': cen('mCadastro.CadastrarOuRedefinir', 'QA Técnico', TEC, cargo_t, 'abc', 'abc', 'CRBM 0', adm, SENHA_ADM),
            'cargo_fora': cen('mCadastro.CadastrarOuRedefinir', 'QA Técnico', TEC, 'Chefe', SENHA_TEC, SENHA_TEC, '', adm, SENHA_ADM),
            'adm_senha_errada': cen('mCadastro.CadastrarOuRedefinir', 'QA Técnico', TEC, cargo_t, SENHA_TEC, SENHA_TEC, '', adm, 'errada123'),
        }
        reg('S02 Recusas sem gravar: senhas diferentes, senha fraca, cargo fora da lista, ADMIN com senha errada',
            all(v.startswith('ERRO|') for v in rec.values()) and len(linhas()) == n0, rec)

        # ------------------------------------------------------------ S03 cadastro
        r3a = cen('mCadastro.CadastrarOuRedefinir', 'QA Técnico de Teste', TEC, cargo_t, SENHA_TEC, SENHA_TEC, 'CRBM 99999', adm, SENHA_ADM)
        r3b = cen('mCadastro.CadastrarOuRedefinir', 'QA Analista de Teste', ANA, 'Analista de Laboratório', SENHA_ANA, SENHA_ANA, 'CRF 88888', adm, SENHA_ADM)
        n3 = len(linhas())
        r3c = cen('mCadastro.CadastrarOuRedefinir', 'Outro Nome', TEC.lower(), cargo_t, SENHA_TEC, SENHA_TEC, '', adm, SENHA_ADM)
        r3d = cen('mCadastro.CadastrarOuRedefinir', '', TEC, '', SENHA_TEC, SENHA_TEC, '', TEC, SENHA_TEC)   # nao-ADMIN
        lt = next(r for r in linhas() if r[1] == TEC)
        texto = []
        for w in wc.Worksheets:
            v = w.UsedRange.Value
            for row in (v if isinstance(v, tuple) else ((v,),)):
                texto.extend(str(x) for x in row if x is not None)
        vaza = [s for s in (SENHA_ADM, SENHA_TEC, SENHA_ANA) if any(s in t for t in texto)]
        reg('S03 Cadastro pelo ADMIN: linhas novas (ASSINATURA vazia, SITUACAO ATIVO); login repetido em outra caixa '
            'não cria linha; não-ADMIN não autoriza; nenhuma senha em texto em nenhuma célula (inclusive LOG_CADASTRO)',
            r3a.startswith('OK|CRIADO') and r3b.startswith('OK|CRIADO') and n3 == n0 + 2 and r3c.startswith('OK|ATUALIZADO')
            and len(linhas()) == n3 and r3d.startswith('ERRO|') and not lt[4] and lt[6] == 'ATIVO' and pbkdf2_ok(SENHA_TEC, lt[3])
            and not vaza,
            {'tecnico': r3a, 'analista': r3b, 'repetido': r3c, 'nao_admin': r3d, 'linhas': [n0, n3, len(linhas())],
             'senhas_em_texto': vaza})
        # imagem de assinatura so do ANALISTA
        pasta = os.path.join(os.path.dirname(central), 'assinaturas')
        os.makedirs(pasta, exist_ok=True)
        png(os.path.join(pasta, 'qa_analista_084.png'))
        rr = next(i for i, r in enumerate(linhas(), 1) if r[1] == ANA)
        lo.DataBodyRange.Cells(rr, 5).Value = 'qa_analista_084.png'
        wc.Save()

        # ------------------------------------------------------------ S04 sincronizar
        q.cfg('CAMINHO_USUARIOS', central) if _tem_cfg(q, 'CAMINHO_USUARIOS') else None
        r4 = qc('mUsuariosCentral.Sincronizar', True)
        uc = q.ws('Usuarios_Central')
        ult = uc.Cells(uc.Rows.Count, 1).End(-4162).Row
        copia = {str(r[0]): r for r in uc.Range(f'A4:I{max(ult, 4)}').Value if r[0]}
        ok4 = (r4.startswith('OK|') and TEC in copia and ANA in copia and copia[TEC][3] == 'TÉCNICO'
               and copia[ANA][3] == 'ANALISTA' and copia[TEC][4] == lt[3] and copia[TEC][7] == 'ATIVO')
        reg('S04 Sincronizar: a cópia Usuarios_Central é o central (papel pelo cargo, hash igual, ATIVO)', ok4,
            {'retorno': r4, 'usuarios': len(copia)})

        # ------------------------------------------------------------ S05 login
        def login(lg, senha):
            q.nome('loginUser').Value = lg
            q.nome('loginPass').Value = senha
            qc('mSeguranca.DoLogin', teto=600)
            return str(q.nome('currentUser').Value or ''), str(q.nome('currentPapel').Value or ''), \
                str(q.nome('loginMsg').Value or '')
        e5a = login(TEC, SENHA_TEC + 'x')
        e5b = login(TEC, SENHA_TEC)
        reg('S05 Login de usuário central: senha errada não entra; certa entra com o papel do cargo',
            e5a[0] == '' and e5b[0] == TEC and e5b[1] == 'TÉCNICO', {'errada': e5a, 'certa': e5b[:2]})

        # ------------------------------------------------------------ S07 central indisponivel
        cam_ok = q.cfg('CAMINHO_USUARIOS')
        q.cfg('CAMINHO_USUARIOS', central + '.nao_existe')
        r7 = qc('mUsuariosCentral.Sincronizar', True)
        copia7 = {str(r[0]) for r in uc.Range(f'A4:I{max(ult, 4)}').Value if r[0]}
        e7 = login(ANA, SENHA_ANA)
        q.cfg('CAMINHO_USUARIOS', cam_ok)
        aud = q.ws('Audit_Log')
        ua = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        acoes7 = [str(aud.Cells(i, 7).Value) for i in range(max(ua - 40, 2), ua + 1)]
        reg('S07 Central indisponível: ERRO, a cópia anterior fica, o login continua; Audit_Log registra',
            r7.startswith('ERRO|') and copia7 == set(copia) and e7[0] == ANA and 'USUARIOS_CENTRAL_INDISPONIVEL' in acoes7,
            {'retorno': r7[:160], 'copia_mantida': copia7 == set(copia), 'login': e7[:2]})

        # ------------------------------------------------------------ S09 / S10 assinatura
        lib = q.ws('Liberação')
        lib.Unprotect(ii.SENHA)
        r_a, r_b = 4, 5
        for r_ in (r_a, r_b):
            lib.Range(f'C{r_}:F{r_}').ClearContents()
        s9 = qc('mSeguranca.AssinarCom', lib.Range(f'E{r_a}'), ANA, SENHA_ANA, 'FINALIZADO')
        txt9 = str(lib.Range(f'F{r_a}').Value or '')
        figs = [sh.Name for sh in lib.Shapes if sh.Name.startswith('qcAssin_')]
        s9b = qc('mSeguranca.AssinarCom', lib.Range(f'C{r_b}'), TEC, SENHA_TEC, 'INICIADO')
        txt9b = str(lib.Range(f'D{r_b}').Value or '')
        figs_b = [sh.Name for sh in lib.Shapes if sh.Name.startswith('qcAssin_')]
        ua9 = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        det9 = ' '.join(str(aud.Cells(i, 24).Value) for i in range(max(ua9 - 6, 2), ua9 + 1))
        reg('S09 Assinatura: texto com nome, cargo, conselho, data/hora e login; imagem só na célula de quem tem imagem '
            'cadastrada; a linha do outro usuário (sem imagem) fica só com o texto; Audit_Log com a identidade',
            s9 == 'OK' and 'QA Analista de Teste' in txt9 and 'Analista de Laboratório' in txt9 and 'CRF 88888' in txt9
            and f'login {ANA}' in txt9 and figs == [f'qcAssin_E{r_a}'] and s9b == 'OK' and f'login {TEC}' in txt9b
            and figs_b == [f'qcAssin_E{r_a}'] and 'QA Analista de Teste' in det9,
            {'assinou': s9, 'texto': txt9, 'imagens': figs, 'outro': s9b, 'texto_outro': txt9b, 'imagens_depois': figs_b})
        s10a = qc('mSeguranca.AssinarCom', lib.Range(f'E{r_b}'), TEC, SENHA_TEC, 'FINALIZADO')
        s10b = qc('mSeguranca.AssinarCom', lib.Range(f'E{r_b}'), ANA, 'errada', 'FINALIZADO')
        reg('S10 FINALIZADO por TÉCNICO e senha errada: recusados, nada carimbado',
            s10a != 'OK' and s10b != 'OK' and not lib.Range(f'E{r_b}').Value,
            {'tecnico_finalizando': s10a, 'senha_errada': s10b})

        # ------------------------------------------------------------ S11 cadastro muda depois
        rr = next(i for i, r in enumerate(linhas(), 1) if r[1] == ANA)
        lo.DataBodyRange.Cells(rr, 1).Value = 'Nome Trocado Depois'
        wc.Save()
        qc('mUsuariosCentral.Sincronizar', True)
        qc('mSeguranca.RedesenharAssinaturas')
        reg('S11 O cadastro muda depois: a liberação já feita continua com o nome da hora da assinatura',
            'QA Analista de Teste' in str(lib.Range(f'F{r_a}').Value) and
            [sh.Name for sh in lib.Shapes if sh.Name.startswith('qcAssin_')] == [f'qcAssin_E{r_a}'],
            {'texto': str(lib.Range(f'F{r_a}').Value)})

        # ------------------------------------------------------------ S06 inativo
        r6 = cen('mCadastro.DefinirSituacao', ANA, False, adm, SENHA_ADM)
        wc.Save()
        qc('mUsuariosCentral.Sincronizar', True)
        e6 = login(ANA, SENHA_ANA)
        s6 = qc('mSeguranca.AssinarCom', lib.Range(f'C{r_b + 1}'), ANA, SENHA_ANA, 'INICIADO')
        reg('S06 Usuário INATIVO no central: não entra nem assina', r6 == 'OK|INATIVO' and e6[0] == '' and s6 != 'OK',
            {'inativar': r6, 'login': e6, 'assinar': s6})

        # ------------------------------------------------------------ S08 usuario local legado
        us = q.ws('Usuarios')
        us.Unprotect(ii.SENHA)
        i8 = next(i for i in range(4, 54) if not us.Cells(i, 1).Value)
        us.Cells(i8, 1).Value = LOC
        us.Cells(i8, 2).Value = 'QA local'
        us.Cells(i8, 3).Value = 'ANALISTA'
        us.Cells(i8, 4).Value = hashlib.sha256(SENHA_LOC.encode()).hexdigest()
        e8a = login(LOC, SENHA_LOC)
        h8 = str(us.Cells(i8, 4).Value)
        e8b = login(LOC, SENHA_LOC)
        e8c = login(LOC, SENHA_LOC + 'x')
        reg('S08 Usuário local com SHA-256 antigo: entra, o hash vira PBKDF2 e continua entrando; senha errada não',
            e8a[0] == LOC and pbkdf2_ok(SENHA_LOC, h8) and e8b[0] == LOC and e8c[0] == '',
            {'1o': e8a[:2], 'hash_novo': h8.split('$')[:2], '2o': e8b[:2], 'errada': e8c[0]})
    except Exception as ex_geral:                    # noqa: BLE001
        reg('S99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        try:
            if wc is not None:
                wc.Close(False)
        except Exception:                            # noqa: BLE001
            pass
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('S00 Segurança do teste: arquivo QC intacto (SHA-256 igual antes e depois)', hash0 == hash1,
        {'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} usuários/assinatura (ADR-084): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


def _tem_cfg(q, chave):
    return chave in [str(r[0] or '').strip().upper() for r in q.lo('tblConfigIntegracao').DataBodyRange.Value]


def central_de_teste(copia_qc):
    """Sem o 3o argumento (entregar.py passa produto, copia, json): copia o cadastro central da raiz do projeto para
    uma pasta ao lado da copia do QC e instala o ADR-084 nele (instalar_cadastro_central). Descartavel."""
    sys.path.insert(0, os.path.join(AQUI, '..'))
    import instalar_cadastro_central as icc
    raiz = os.environ.get('QC_RAIZ') or os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
    origem = os.path.join(raiz, 'QC_INI_cadastrar usuários.xlsm')
    if not os.path.exists(origem):
        raise SystemExit(f'cadastro central nao encontrado para o teste: {origem}')
    pasta = os.path.join(os.path.dirname(copia_qc), 'central_teste')
    shutil.rmtree(pasta, ignore_errors=True)
    os.makedirs(pasta)
    destino = os.path.join(pasta, os.path.basename(origem))
    shutil.copy2(origem, destino)
    icc.main(destino)
    return destino


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) == 3:                          # produto, copia, saida (como as outras suites no entregar.py)
        a = [a[0], a[1], central_de_teste(os.path.abspath(a[1])), a[2]]
    if len(a) < 4:
        raise SystemExit(__doc__)
    falhas = executar(a[0], os.path.abspath(a[1]), os.path.abspath(a[2]), a[3])
    sys.exit(1 if falhas else 0)
