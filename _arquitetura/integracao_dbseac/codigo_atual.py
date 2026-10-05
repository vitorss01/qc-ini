# -*- coding: utf-8 -*-
"""codigo_atual.py -- instala de UMA vez todo o VBA versionado que mudou desde o ADR-057.

Os modulos se referenciam entre as entregas (mSeguranca chama mLotes e mEstatistica; mApp e mUI
chamam a protecao rapida do mSeguranca; mIntegracao chama mLotes...). Instalar um modulo por vez,
cada um no seu instalador, deixava o projeto sem compilar conforme a ORDEM. Aqui todos entram
juntos, conferidos contra a fonte e compilados; os instaladores (instalar_*.py) chamam esta
funcao e cuidam so da parte de planilha que e deles -- a ordem entre eles deixa de importar.
"""
import os

import instalar_integracao as ii

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, 'src')

# modulos padrao (fonte em src/<bio|hema> quando existe, senao src/comum)
MODULOS = ['mLotes', 'mApp', 'mIntegracao', 'mSeguranca', 'mEstatistica', 'mUI',
           'mWestgardKnowledge', 'mAuditoria',
           'mCEQ', 'mEQA', 'mPlanoQC',          # ADR-064: CEQ da incerteza; texto acentuado corrompido
           'mDados']                            # ADR-064 (auditoria): ultima data so do que participa
# modulos que podem ainda nao existir no arquivo (criados na primeira instalacao)
NOVOS = ['mIncerteza']                         # ADR-064
# modulos de aba: (nome da aba, arquivo-fonte)
ABAS = [('Painel', 'Painel.cls'), ('Configuração', 'Configuracao.cls')]
PASTA = 'EstaPastaDeTrabalho.cls'


def normalizar(t):
    linhas = str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n')
    return '\n'.join(l.rstrip() for l in linhas)


def _fonte(d, nome):
    for pasta in (os.path.join(SRC, d), os.path.join(SRC, 'comum')):
        f = os.path.join(pasta, nome)
        if os.path.exists(f):
            return f
    raise SystemExit(f'fonte nao encontrada: {nome}')


def _trocar(wb, comp, fonte, criar):
    texto = ii.codigo(fonte)
    ii.substituir_codigo(wb.VBProject, comp, texto, criar=criar)
    cm = wb.VBProject.VBComponents(comp).CodeModule
    # VBA e insensivel a maiusculas e o VBE reescreve a caixa de identificador ja conhecido
    if normalizar(cm.Lines(1, cm.CountOfLines)).lower() != normalizar(texto).lower():
        raise SystemExit(f'{comp} no arquivo difere da fonte {fonte}')


def instalar(ex, wb, produto, log=print):
    """Troca o VBA pelo das fontes e EXIGE compilacao. Devolve a lista do que foi instalado."""
    d = 'bio' if produto.startswith('Bio') else 'hema'
    feitos = []
    for m in MODULOS:
        _trocar(wb, m, _fonte(d, m + '.bas'), criar=False)
        feitos.append(m)
    for m in NOVOS:
        _trocar(wb, m, _fonte(d, m + '.bas'), criar=True)
        feitos.append(m)
    for aba, arq in ABAS:
        cn = [ws.CodeName for ws in wb.Worksheets if ws.Name == aba][0]
        _trocar(wb, cn, _fonte(d, arq), criar=False)
        feitos.append(f'{cn} ({aba})')
    _trocar(wb, wb.CodeName, _fonte(d, PASTA), criar=False)
    feitos.append(wb.CodeName)
    log('   VBA identico a fonte: ' + ', '.join(feitos))
    # Compilar e OBRIGATORIO: com Interactive=False um erro de compilacao nao abre modal na
    # instalacao -- o comando so continua habilitado; o modal apareceria no 1o uso.
    if ii.compilar_vba(ex, wb) is not True:
        raise SystemExit('VBA NAO COMPILOU (Depurar > Compilar continua habilitado) -- nada foi salvo')
    return feitos
