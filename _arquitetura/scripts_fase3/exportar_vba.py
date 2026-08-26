# -*- coding: utf-8 -*-
"""exportar_vba.py - despeja TODOS os modulos VBA de um artefato em disco

POR QUE

A auditoria estatica precisa da verdade de campo, e ela nao esta em nenhuma
pasta de fonte: o artefato e a soma do workbook de producao (modulos que nunca
sao regerados) com os modulos que o build importa. Duplicata de nome publico
entre um modulo gerado e um modulo legado da producao so aparece olhando o
projeto INTEIRO -- e e justamente esse tipo de conflito que derruba a
compilacao com mensagens que apontam para o lugar errado.

Modulos VBA sao cp1252. Exporta com a codificacao certa para o analisador nao
tropecar em acento.

Uso: python exportar_vba.py <arquivo.xlsm> <pasta_destino>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

TIPO = {1: 'bas', 2: 'cls', 3: 'frm', 11: 'cls', 100: 'doc'}


def main():
    caminho = os.path.abspath(sys.argv[1])
    destino = os.path.abspath(sys.argv[2])
    if not os.path.isdir(destino):
        os.makedirs(destino)

    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, True)
    try:
        vbp = wb.VBProject
        n = 0
        for c in vbp.VBComponents:
            ext = TIPO.get(c.Type, 'txt')
            cm = c.CodeModule
            linhas = cm.CountOfLines
            texto = cm.Lines(1, linhas) if linhas else ''
            alvo = os.path.join(destino, '%s.%s' % (c.Name, ext))
            # newline='' e OBRIGATORIO: CodeModule.Lines ja devolve CRLF. Com
            # newline='\r\n' o Python traduz o \n que ja esta la e grava
            # "\r\r\n". Na releitura em modo texto o \r solto conta como quebra,
            # cada modulo aparece com o DOBRO de linhas e as fronteiras de
            # procedure saem deslocadas -- a auditoria estatica roda inteira
            # sobre lixo e devolve "nada encontrado" com ar de aprovacao.
            io.open(alvo, 'w', encoding='cp1252',
                    errors='replace', newline='').write(texto)
            n += 1
            print('  %-28s %-4s %5d linhas' % (c.Name, ext, linhas))
        print('%d componentes exportados para %s' % (n, destino))
    finally:
        wb.Close(False)
        xl.Quit()
    return 0


if __name__ == '__main__':
    sys.exit(main())
