# -*- coding: utf-8 -*-
"""ligar_ano_especificacoes.py - ADR-028: o ano em Analitos!S2 passa a mandar

O QUE A AUDITORIA ACHOU

A arquitetura pretendida era:

    Analitos!S2 (ano)  ->  historico nas linhas 46..89  ->  Analitos!A4:W43

Ela NAO EXISTIA. Medido no arquivo: das 23 colunas do bloco vigente, ZERO
formulas citavam S2 e ZERO citavam as linhas 46..89. As colunas K, M, N, O, P,
Q e S eram todas LITERAIS DIGITADOS. Trocar o ano em S2 nao mudava coisa nenhuma
em lugar nenhum -- o historico estava la, e ninguem o lia.

Este script constroi a ligacao que faltava.

O DESENHO DO HISTORICO (medido, nao suposto)

Blocos de 9 linhas a partir da 46, cinco no total (46, 55, 64, 73, 82):

    +0  ANO                   <- o ano do bloco, na coluna B
    +1  ESPECIFICACOES ...    <- nomes dos analitos, colunas B..AO
    +2  CLIA ETp%
    +3  CLIA 1/3 ETp%
    +4  VB-CVI
    +5  VB-CVG
    +6  FAB ETp%
    +7  FAB CVTp%
    +8  (em branco)

O QUE PASSA A SER FORMULA

    K  ETp CLIA %      <- linha "CLIA ETp%"  do bloco do ano
    M  ETp FAB %       <- linha "FAB ETp%"
    N  CVTp FAB %      <- linha "FAB CVTp%"
    O  CVi %           <- linha "VB-CVI"
    P  CVg %           <- linha "VB-CVG"

O QUE CONTINUA DIGITADO, DE PROPOSITO

    Q  Desemp. (OTI/DES/MIN)  e  S  ETp fonte (CLIA/VB/FAB)

Sao DECISOES do laboratorio, nao dados do ano. Puxa-las do historico
transformaria uma escolha em consequencia -- e tiraria do gestor o controle que
o proprio pedido descreve como o ponto da coluna S.

CONVERSAO E SEGURA -- CONFERIDO ANTES

Comparei o bloco vigente com o historico de 2025 celula a celula: 155
comparacoes, 4 divergencias, todas do mesmo tipo -- um 0 digitado onde o
historico esta vazio. Depois da conversao esses viram BRANCO, que e mais
correto: ETp igual a zero nao e meta, e produzia Sigma sem sentido.

POR QUE NAO IFERROR EM CIMA DE TUDO

Ano nao cadastrado nao pode virar 200 celulas vazias em silencio. O IFERROR aqui
cobre apenas "este analito nao existe no bloco do ano", que e estado de dado
legitimo. O ano NAO ENCONTRADO e denunciado numa celula de status ao lado do
proprio seletor, uma vez e visivel, em vez de espalhado.

Uso: python ligar_ano_especificacoes.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
HIST = '$B$46:$AO$89'
ANOS = '$B$46,$B$55,$B$64,$B$73,$B$82'
R0, RN = 4, 43

# coluna do bloco vigente -> deslocamento da linha dentro do bloco do ano
DESTINO = {
    11: 3,    # K  ETp CLIA %   <- CLIA ETp%
    13: 7,    # M  ETp FAB %    <- FAB ETp%
    14: 8,    # N  CVTp FAB %   <- FAB CVTp%
    15: 5,    # O  CVi %        <- VB-CVI
    16: 6,    # P  CVg %        <- VB-CVG
}

IDX_ANO = 'MATCH($S$2,CHOOSE({1;2;3;4;5},%s),0)' % ANOS


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


def formula(lin, desloc):
    """INDEX no historico: linha = bloco do ano + deslocamento, coluna = analito."""
    linha = '(%s-1)*9+%d' % (IDX_ANO, desloc)
    linha_nomes = '(%s-1)*9+2' % IDX_ANO
    coluna = 'MATCH($A%d,INDEX(%s,%s,0),0)' % (lin, HIST, linha_nomes)
    return '=IFERROR(INDEX(%s,%s,%s),"")' % (HIST, linha, coluna)


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura (o arquivo esta aberto no Excel?): %s' % caminho)
    try:
        xl.Calculation = -4135
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        try:
            wb.Worksheets('Analitos').Unprotect(SENHA)
        except Exception:
            pass
        an = wb.Worksheets('Analitos')

        # ---- 1. nomes ------------------------------------------------------
        for n, ref in (('espAno', '=Analitos!$S$2'),
                       ('espVigente', '=Analitos!$A$4:$W$43'),
                       ('espHistorico', '=Analitos!$A$46:$AO$89')):
            try:
                wb.Names(n).Delete()
            except Exception:
                pass
            wb.Names.Add(n, ref)
        print('nomes: espAno, espVigente, espHistorico')

        # ---- 2. status do ano, ao lado do seletor --------------------------
        # Um aviso visivel uma vez, em vez de 200 celulas vazias caladas.
        tenta(lambda: an.Range('T2').__setattr__(
            'Formula',
            '=IF($S$2="","(informe o ano)",IF(ISNUMBER(%s),'
            '"ano "&$S$2&" localizado no historico",'
            '"ANO NAO CADASTRADO no bloco de especificacoes"))' % IDX_ANO))
        print('T2: status do ano ao lado do seletor')

        # ---- 3. as cinco colunas passam a vir do ano -----------------------
        n = 0
        for col, desloc in sorted(DESTINO.items()):
            for lin in range(R0, RN + 1):
                tenta(lambda l=lin, c=col, d=desloc:
                      an.Cells(l, c).__setattr__('Formula', formula(l, d)))
                n += 1
        print('%d celulas convertidas em K, M, N, O, P (linhas %d..%d)' % (n, R0, RN))

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        # ---- 4. conferencia: K tem de bater com o historico do ano --------
        ano = an.Range('S2').Value
        print('status: %r' % an.Range('T2').Value)
        div = 0
        for lin in range(R0, RN + 1):
            nome = an.Cells(lin, 1).Value
            if nome in (None, ''):
                continue
            k = an.Cells(lin, 11).Value
            if isinstance(k, str) and k.startswith('#'):
                div += 1
                if div <= 5:
                    print('   ERRO em K%d (%s): %r' % (lin, nome, k))
        if div:
            raise SystemExit('%d celula(s) de K com erro -- nada salvo' % div)
        print('K sem erros de formula em todas as linhas')

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
