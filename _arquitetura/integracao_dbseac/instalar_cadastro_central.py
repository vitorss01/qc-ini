# -*- coding: utf-8 -*-
"""instalar_cadastro_central.py -- ADR-084: o arquivo QC_INI_cadastrar usuarios.xlsm vira o cadastro central.

Uso:  python instalar_cadastro_central.py <arquivo.xlsm> [--sem-salvar]

Idempotente. Passos:
  1. confere a aba REGISTRO_DE_USUARIOS e a tabela tb_registro_usuarios (A FUNCIONARIOS, B LOGIN, C CARGO);
  2. cabecalhos: D SENHA -> SENHA_HASH (a coluna guarda so hash) e G SITUACAO (nova; ATIVO nas linhas existentes);
     a ordem A..F nao muda (os QCs leem por posicao). Se a coluna D tiver algo que nao e hash PBKDF2, PARA: nao ha
     como saber se e senha em texto (nesta versao do arquivo ela esta vazia);
  3. VBA: mSenha (o mesmo dos QCs) e mCadastro; o frmCadastroUsuario (declarado aqui). Saem o UserForm1/UserForm2 (procuravam a aba RESTRITO/tblUsuarios, que nao existe -- nao cadastravam
     nada; senha mestre fixa no codigo) e os Modulo4/Modulo5 (SHA-256 sem salt pelo PowerShell, com a senha montada
     na linha de comando). Nenhum outro modulo e tocado. EXIGE compilacao;
  4. aba LOG_CADASTRO (trilha do cadastro, sem senha) e o botao "CADASTRAR / REDEFINIR SENHA" na aba do cadastro;
  5. pasta "assinaturas" ao lado do arquivo (para as imagens; a coluna ASSINATURA recebe o nome do arquivo).
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import instalar_frm_inativar as fi  # noqa: E402

ABA = 'REGISTRO_DE_USUÁRIOS'
TABELA = 'tb_registro_usuários'
SRC = os.path.join(AQUI, 'src')
SAI = ['UserForm1', 'UserForm2', 'Módulo4', 'Módulo5']
FONTE_UI = 'Segoe UI'
MUDANCAS = []

# formulario: 2 colunas de rotulo + campo; medidas em multiplos de 3 pt
FORM = {'Caption': 'Cadastro de profissionais — QC_INI', 'Width': 390, 'Height': 381, 'StartUpPosition': 1}


def _campo(nome, rot, top, tab, senha=False, lista=False):
    lbl = ('Forms.Label.1', 'lbl' + nome, {'Left': 12, 'Top': top + 3, 'Width': 129, 'Height': 15, 'Caption': rot,
                                           'Font.Name': FONTE_UI, 'Font.Size': 9})
    props = {'Left': 144, 'Top': top, 'Width': 222, 'Height': 21, 'Font.Name': FONTE_UI, 'Font.Size': 9,
             'TabIndex': tab}
    if senha:
        props['PasswordChar'] = '*'
    if lista:
        props['Style'] = 2
        return [lbl, ('Forms.ComboBox.1', 'cbo' + nome, props)]
    return [lbl, ('Forms.TextBox.1', 'txt' + nome, props)]


CONTROLES = ([('Forms.Label.1', 'lblTitulo', {'Left': 12, 'Top': 9, 'Width': 354, 'Height': 18,
                                              'Caption': 'Profissional', 'Font.Name': FONTE_UI, 'Font.Size': 11,
                                              'Font.Bold': True})]
             + _campo('Nome', 'Nome completo', 33, 0) + _campo('Login', 'Login', 60, 1)
             + _campo('Cargo', 'Cargo', 87, 2, lista=True) + _campo('Conselho', 'Conselho e nº (se houver)', 114, 3)
             + _campo('Senha', 'Senha (mín. 8, letras e números)', 141, 4, senha=True)
             + _campo('Confirma', 'Confirmar a senha', 168, 5, senha=True)
             + [('Forms.Label.1', 'lblAutoriza', {'Left': 12, 'Top': 198, 'Width': 354, 'Height': 18,
                                                  'Caption': 'Autorização do administrador', 'Font.Name': FONTE_UI,
                                                  'Font.Size': 10, 'Font.Bold': True})]
             + _campo('AdmLogin', 'Login do ADMIN', 219, 6) + _campo('AdmSenha', 'Senha do ADMIN', 246, 7, senha=True)
             + [('Forms.Label.1', 'lblMsg', {'Left': 12, 'Top': 273, 'Width': 354, 'Height': 33, 'Caption': '',
                                             'WordWrap': True, 'ForeColor': 0x606060, 'Font.Name': FONTE_UI,
                                             'Font.Size': 8}),
                ('Forms.CommandButton.1', 'btnSalvar', {'Left': 12, 'Top': 312, 'Width': 114, 'Height': 27,
                                                        'Caption': 'Salvar', 'Font.Name': FONTE_UI, 'Font.Size': 9,
                                                        'Font.Bold': True, 'TabIndex': 8}),
                ('Forms.CommandButton.1', 'btnInativar', {'Left': 132, 'Top': 312, 'Width': 114, 'Height': 27,
                                                          'Caption': 'Inativar login', 'Font.Name': FONTE_UI,
                                                          'Font.Size': 9, 'TabIndex': 9}),
                ('Forms.CommandButton.1', 'btnCancelar', {'Left': 252, 'Top': 312, 'Width': 114, 'Height': 27,
                                                          'Caption': 'Fechar', 'Cancel': True, 'Font.Name': FONTE_UI,
                                                          'Font.Size': 9, 'TabIndex': 10})])


def log(m):
    print(m, flush=True)


def main(caminho, salvar=True):
    import codigo_atual
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA: {caminho}')
        ws = wb.Worksheets(ABA)
        lo = ws.ListObjects(TABELA)
        cab = [c.Name for c in lo.ListColumns]
        if [x.upper() for x in cab[:3]] != ['FUNCIONÁRIOS', 'LOGIN', 'CARGO']:
            raise SystemExit(f'tabela com layout inesperado: {cab} -- nada foi feito')
        log(f'1. {ABA}!{TABELA}: {lo.ListRows.Count} linha(s), colunas {cab}')
        # 2. coluna D: so hash (ou vazia) -- senha em texto nao e convertida em silencio
        if lo.ListRows.Count:
            d = [str(r[0] or '').strip() for r in lo.ListColumns(4).DataBodyRange.Value]
            ruins = [i + 1 for i, x in enumerate(d) if x and not x.lower().startswith('pbkdf2_sha256$')]
            if ruins:
                raise SystemExit(f'coluna D com conteudo que nao e hash PBKDF2 nas linhas {ruins}: confira antes '
                                 '(senha em texto?) -- nada foi feito')
        if cab[3] != 'SENHA_HASH':
            lo.ListColumns(4).Name = 'SENHA_HASH'
            MUDANCAS.append('D = SENHA_HASH')
        if len(cab) < 7:
            lo.ListColumns.Add()
            lo.ListColumns(7).Name = 'SITUACAO'
            MUDANCAS.append('G = SITUACAO')
        if lo.ListRows.Count:
            g = lo.ListColumns(7).DataBodyRange
            v = [[x[0] if x[0] else 'ATIVO'] for x in g.Value]
            if any(not x[0] for x in g.Value):
                g.Value = tuple(tuple(x) for x in v)
                MUDANCAS.append('SITUACAO = ATIVO nas linhas existentes')
            lo.ListColumns(4).DataBodyRange.NumberFormat = '@'
        log('2. cabecalhos SENHA_HASH / SITUACAO')
        # 3. VBA
        vbp = wb.VBProject
        nomes = [c.Name for c in vbp.VBComponents]
        for n in SAI:
            if n in nomes:
                vbp.VBComponents.Remove(vbp.VBComponents(n))
                MUDANCAS.append(f'removido {n}')
        for mod, pasta in (('mSenha', 'comum'), ('mCadastro', 'cadastro')):
            texto = ii.codigo(os.path.join(SRC, pasta, mod + '.bas'))
            ii.substituir_codigo(vbp, mod, texto, criar=True)
            cm = vbp.VBComponents(mod).CodeModule
            if codigo_atual.normalizar(cm.Lines(1, cm.CountOfLines)).lower() != codigo_atual.normalizar(texto).lower():
                raise SystemExit(f'{mod} difere da fonte -- nada foi salvo')
        n = fi.criar_form(wb, log, nome='frmCadastroUsuario', fonte=os.path.join(SRC, 'cadastro', 'frmCadastroUsuario.frm.txt'),
                          form=FORM, controles=CONTROLES)
        if n:
            MUDANCAS.append('frmCadastroUsuario')
        if ii.compilar_vba(ex, wb) is not True:
            raise SystemExit('VBA NAO COMPILOU -- nada foi salvo')
        log('3. VBA: mSenha, mCadastro, frmCadastroUsuario; compilado')
        # 4. LOG_CADASTRO e botao
        if 'LOG_CADASTRO' not in [w.Name for w in wb.Worksheets]:
            lg = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            lg.Name = 'LOG_CADASTRO'
            lg.Range('A1').Value = 'Trilha do cadastro central (ADR-084). Sem senha nem hash.'
            lg.Range('A2:F2').Value = ('DATA_HORA', 'USUARIO_WINDOWS', 'ADMINISTRADOR', 'ACAO', 'LOGIN_ALVO', 'DETALHE')
            lg.Range('A2:F2').Font.Bold = True
            MUDANCAS.append('aba LOG_CADASTRO')
        if 'btnCadastro' not in [s.Name for s in ws.Shapes]:
            alvo = ws.Range('I2')
            b = ws.Shapes.AddShape(5, alvo.Left, alvo.Top, 200, 30)      # retangulo arredondado
            b.Name = 'btnCadastro'
            b.TextFrame2.TextRange.Text = 'CADASTRAR / REDEFINIR SENHA'
            b.TextFrame2.TextRange.Font.Bold = True
            b.TextFrame2.TextRange.Font.Size = 10
            b.Fill.ForeColor.RGB = 0x794E1F
            b.OnAction = 'mCadastro.AbrirCadastro'
            MUDANCAS.append('botao CADASTRAR')
        log('4. LOG_CADASTRO e botao')
        pasta = os.path.join(os.path.dirname(caminho), 'assinaturas')
        if not os.path.isdir(pasta):
            os.makedirs(pasta)
            MUDANCAS.append('pasta assinaturas')
        log(f'5. pasta de imagens: {pasta}')
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar and MUDANCAS:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a:
        raise SystemExit(__doc__)
    main(a[0], salvar='--sem-salvar' not in a)
