# -*- coding: utf-8 -*-
"""remover_motor_especificacoes.py - ADR-028: sai a camada antiga inteira

CLASSIFICACAO, DEPOIS DE RASTREAR AS DEPENDENCIAS

  Cfg_Especificacoes   C - OBSOLETA. Guardava o catalogo de fontes e modelos.
                       As fontes hoje sao a lista de validacao "CLIA,VB,FAB" em
                       Analitos!S, e o modelo de cada uma virou coluna do bloco.
  DB_Especificacoes    C - OBSOLETA. Era o historico por ano. O historico agora
                       vive na propria Analitos (linhas 46..89), que e o que o
                       ADR-028 construiu. Alem disso ela estava VAZIA: LimEspec
                       devolvia branco para os 40 analitos.
  Eng_Especificacoes   C - OBSOLETA. Camada de saida do motor. Nenhuma formula a
                       le desde o ADR-027; o ultimo consumidor era o mBI, agora
                       apontado para Analitos S/T/U.

O ADR-027 classificou DB_Especificacoes como INDISPENSAVEL porque a Analitos nao
tinha dimensao de ano. Isso deixou de ser verdade: o bloco 46..89 tem o ano, e a
justificativa caiu junto. Registrar a mudanca de veredito importa tanto quanto o
veredito.

O QUE E REMOVIDO

  shape btnEspec              botao "Especificacoes" na Analitos
  frmEspecificacoes           o formulario de cadastro
  mEspecificacoes.bas         o motor: ResolverEspec, GravarEspec, EspecETp...
  LimEspec (mEstatPeriodo)    unico consumidor do motor; nenhuma formula o chama
  AtualizarEngEspec           chamada em mOperacao que populava a Eng
  nomes eng*                  engETp, engCVtp, engBIAStp, engEspAnalito, engEspAno
  as tres abas

O QUE FICA, E POR QUE

  StatusCV e StatusETP        continuam em mEstatPeriodo e sao usados pelas
                              colunas J e Q da Estatistica. Nao dependem do motor.

CONSEQUENCIA QUE O GESTOR PRECISA SABER

Sai junto a tela de cadastro de especificacoes. O cadastro passa a ser feito
digitando no proprio bloco historico da Analitos (linhas 46..89), que e onde o
dado agora mora. Nao ha perda de informacao -- o DB estava vazio.

Uso: python remover_motor_especificacoes.py <arquivo.xlsm>
"""
import io
import os
import re
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
ABAS = ('Cfg_Especificacoes', 'DB_Especificacoes', 'Eng_Especificacoes')
NOMES = ('engETp', 'engCVtp', 'engBIAStp', 'engEspAnalito', 'engEspAno')


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (2, 4):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(8)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def tenta(fn, vezes=8):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            s = str(e).lower()
            if 'rejeitada' not in s and 'rejected' not in s:
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        for s in wb.Worksheets:
            try:
                s.Unprotect(SENHA)
            except Exception:
                pass

        # ---- 0. TRAVA: nenhuma formula pode citar as abas nem LimEspec ----
        pend = []
        for s in wb.Worksheets:
            if s.Name in ABAS:
                continue
            try:
                ur = s.UsedRange
                achou = ur.SpecialCells(-4123)     # xlCellTypeFormulas
            except Exception:
                continue
            for cel in achou:
                f = str(cel.Formula)
                for a in ABAS + ('LimEspec',):
                    if a in f:
                        pend.append('%s!%s -> %s' % (s.Name, cel.Address, a))
        if pend:
            for p in pend[:10]:
                print('   PENDENTE %s' % p)
            raise SystemExit('%d formula(s) ainda dependem da camada antiga -- nada removido'
                             % len(pend))
        print('nenhuma formula depende das abas antigas nem de LimEspec')

        # ---- 1. o botao ---------------------------------------------------
        an = wb.Worksheets('Analitos')
        rem = []
        for sh in list(an.Shapes):
            try:
                acao = str(sh.OnAction)
            except Exception:
                acao = ''
            if sh.Name == 'btnEspec' or 'Especificacoes' in acao:
                rem.append('%s (OnAction=%s)' % (sh.Name, acao))
                sh.Delete()
        print('shapes removidos da Analitos: %s' % (rem if rem else 'nenhum'))

        # ---- 2. VBA -------------------------------------------------------
        vbp = wb.VBProject
        # LimEspec sai de mEstatPeriodo
        for c in vbp.VBComponents:
            if c.Name == 'mEstatPeriodo':
                # Remocao pelo TEXTO EXATO da funcao, nao por ProcStartLine.
                #
                # ProcStartLine/ProcCountLines incluem os comentarios anteriores e
                # a fronteira nao e obvia: a primeira tentativa apagou 46 linhas e
                # deixou mEstatPeriodo sem compilar. Como TODAS as UDFs vivem
                # nesse modulo, o projeto inteiro parou e 642 celulas viraram
                # #NOME? -- EstatPeriodo, StatusCV e StatusETP juntas.
                #
                # Recortar da assinatura ate o "End Function" correspondente e
                # deterministico e nao depende de onde o VBE acha que a funcao
                # comeca.
                cm = c.CodeModule
                txt = cm.Lines(1, cm.CountOfLines)
                padrao = (r'(?is)\r?\n[^\r\n]*Public Function LimEspec\b'
                          r'.*?\r?\nEnd Function')
                novo_txt = re.sub(padrao, '', txt, count=1)
                if novo_txt == txt:
                    print('   LimEspec nao encontrada por texto -- nada removido')
                else:
                    cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(novo_txt)
                    print('LimEspec removida de mEstatPeriodo (%d -> %d linhas)'
                          % (txt.count(chr(10)), novo_txt.count(chr(10))))
            if c.Name == 'mOperacao':
                cm = c.CodeModule
                txt = cm.Lines(1, cm.CountOfLines)
                linhas = txt.split('\r\n')
                for i in range(len(linhas) - 1, -1, -1):
                    if 'AtualizarEngEspec' in linhas[i]:
                        cm.DeleteLines(i + 1, 1)
                        print('chamada AtualizarEngEspec removida de mOperacao')

        for alvo in ('frmEspecificacoes', 'mEspecificacoes'):
            for c in list(vbp.VBComponents):
                if c.Name == alvo:
                    vbp.VBComponents.Remove(c)
                    print('componente removido: %s' % alvo)

        # ---- 3. nomes -----------------------------------------------------
        tirados = []
        for n in NOMES:
            try:
                wb.Names(n).Delete()
                tirados.append(n)
            except Exception:
                pass
        print('nomes removidos: %s' % (tirados if tirados else 'nenhum'))

        # ---- 4. abas ------------------------------------------------------
        for nome in ABAS:
            for s in list(wb.Worksheets):
                if s.Name == nome:
                    s.Visible = -1
                    s.Delete()
                    print('aba removida: %s' % nome)

        # ---- 5. conferencia final ----------------------------------------
        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
        # SpecialCells(formulas, 16) devolve SO as celulas com ERRO.
        #
        # Percorrer todas as formulas de todas as abas e dezenas de milhares de
        # travessias COM, e o Excel corta com RPC_E_CALL_REJECTED no meio da
        # conferencia -- que e justamente quando nao se pode desistir. Deixar o
        # proprio Excel filtrar e instantaneo.
        ruins = 0
        for s in wb.Worksheets:
            try:
                achou = tenta(lambda ws=s: ws.UsedRange.SpecialCells(-4123, 16))
            except Exception:
                continue                       # nenhuma celula com erro nesta aba
            for cel in achou:
                t = str(tenta(lambda c=cel: c.Text))
                # #N/D e DELIBERADO no Calc: e a lacuna de serie que faz o
                # grafico interromper a linha em vez de liga-la em zero. A
                # propria suite o exclui na verificacao 4.6. Contar #N/D como
                # defeito faria a remocao abortar por causa de um recurso.
                if '#REF!' not in t and '#NOME?' not in t and '#NAME?' not in t:
                    continue
                ruins += 1
                if ruins <= 5:
                    print('   ERRO %s!%s = %s' % (s.Name, cel.Address, t))
        print('celulas com #REF!/#NOME? apos remocao: %d' % ruins)
        if ruins:
            raise SystemExit('remocao deixou referencia quebrada -- nada salvo')

        # ---- 5b. o projeto ainda COMPILA? ---------------------------------
        #
        # Uma UDF quebrada nao levanta erro: ela devolve #NOME? em silencio, em
        # centenas de celulas. Chamar uma delas e a unica prova de que o modulo
        # sobreviveu a remocao.
        try:
            r = xl.Run('StatusCV', 1.0, 2.0, '', '')
            print('prova de compilacao: StatusCV devolveu %r' % r)
        except Exception as e:
            raise SystemExit('VBA nao compila apos a remocao: %s' % str(e)[:120])

        nomes_restantes = [n.Name for n in wb.Names
                           if any(a in str(n.RefersTo) for a in ABAS)]
        if nomes_restantes:
            raise SystemExit('nomes orfaos apontando para abas removidas: %s' % nomes_restantes)
        print('nenhum nome definido aponta para as abas removidas')

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
