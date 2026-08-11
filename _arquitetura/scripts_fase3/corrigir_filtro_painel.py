# -*- coding: utf-8 -*-
"""corrigir_filtro_painel.py - dois defeitos do filtro Ano/Mes do Painel

DEFEITO 1 - O CAMPO DO ANO CAIU DENTRO DE UMA CELULA MESCLADA

H3:I3 ja era mesclada no layout original. O rotulo "Ano" foi para H3 (a ancora)
e o campo I3 ficou DENTRO da mescla: escrever nele nao guarda nada. A linha 4
(Mes) nao e mesclada e por isso funcionava -- o que tornava o defeito ainda
mais confuso, porque metade do filtro respondia.

DEFEITO 2 - ANO AUSENTE VIRAVA ANO 2000, EM SILENCIO

Com I3 vazio, CLng("") falha, o On Error leva ao rotulo de saida com a = 0, e
DateSerial(0, 1, 1) devolve 01/01/2000 -- a regra de ano de dois digitos do VBA
trata 0 como 2000. O Painel passava a filtrar um periodo sem nenhum dado e
ficava vazio SEM DIZER POR QUE.

Este e o modo de falha que mais preocupa num sistema de controle de qualidade:
nao o erro visivel, e o numero plausivel e errado. Entrada ausente agora NAO
escreve periodo nenhum.

Uso: python corrigir_filtro_painel.py <caminho.xlsm>
"""
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'

ROTINA = [
    "' Ano e Mes sao ATALHO para preencher De/Ate, nao um filtro paralelo.",
    "'",
    "' O Calc filtra por filtroDe/filtroAte. Um segundo caminho de filtragem",
    "' seria uma segunda verdade sobre o mesmo periodo.",
    "'",
    "' ANO AUSENTE NAO ESCREVE PERIODO NENHUM.",
    "'",
    "' Antes, I3 vazio levava a = 0 e DateSerial(0, 1, 1) devolvia 01/01/2000 --",
    "' a regra de ano de dois digitos do VBA le 0 como 2000. O Painel filtrava um",
    "' periodo sem dado e ficava vazio sem dizer por que. Numero plausivel e",
    "' errado e pior que erro visivel: ninguem vai conferir o que parece certo.",
    "Public Sub AplicarFiltroAnoMes()",
    "    Dim a As Long, m As Long, s As String, ev As Boolean, v As Variant",
    "    ev = Application.EnableEvents",
    "    On Error GoTo fim",
    "    v = Me.Range(\"I3\").Value",
    "    If Not IsNumeric(v) Then Exit Sub",
    "    a = CLng(v)",
    "    If a < 1900 Or a > 2200 Then Exit Sub",
    "    Application.EnableEvents = False",
    "    s = Trim$(CStr(Me.Range(\"I4\").Value))",
    "    m = MesDoRotulo(s)",
    "    If m = 0 Then",
    "        Me.Range(\"G3\").Value = DateSerial(a, 1, 1)",
    "        Me.Range(\"G4\").Value = DateSerial(a, 12, 31)",
    "    Else",
    "        Me.Range(\"G3\").Value = DateSerial(a, m, 1)",
    "        Me.Range(\"G4\").Value = DateSerial(a, m + 1, 0)",
    "    End If",
    "    AtualizarEixos",
    "fim:",
    "    Application.EnableEvents = ev",
    "End Sub",
    "",
    "Private Function MesDoRotulo(ByVal s As String) As Long",
    "    Dim v As Variant, i As Long",
    "    v = Array(\"JAN\", \"FEV\", \"MAR\", \"ABR\", \"MAI\", \"JUN\", _",
    "              \"JUL\", \"AGO\", \"SET\", \"OUT\", \"NOV\", \"DEZ\")",
    "    For i = 0 To 11",
    "        If UCase$(Left$(s, 3)) = v(i) Then MesDoRotulo = i + 1: Exit Function",
    "    Next i",
    "End Function",
]


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura')
    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        pa = wb.Worksheets('Painel')
        for t in (lambda: pa.Unprotect(SENHA), lambda: pa.Unprotect()):
            try:
                t()
                break
            except Exception:
                pass

        larg_antes = pa.Columns(1).ColumnWidth

        # ---- 1. desfaz a mescla que engolia o campo do ano ----------------
        antes = [str(m.Address) for m in pa.UsedRange.MergeArea] if False else None
        cel = pa.Range('H3:I3')
        if cel.MergeCells:
            cel.UnMerge()
            print('H3:I3 desmesclada')
        else:
            print('H3:I3 ja nao estava mesclada')
        pa.Range('H3').Value = 'Ano'
        pa.Range('H3').Font.Bold = True
        pa.Range('H3').HorizontalAlignment = -4152
        campo = pa.Range('I3')
        campo.Interior.Color = 15921906
        campo.Borders.LineStyle = 1
        campo.HorizontalAlignment = -4108
        try:
            campo.Locked = False
        except Exception:
            pass

        # o dropdown precisa ser recriado: a mescla o invalidava
        db = wb.Worksheets('DB_Resultados')
        ult = db.Cells(db.Rows.Count, 1).End(-4162).Row
        anos = set()
        if ult >= 4:
            for linha in db.Range(db.Cells(4, 2), db.Cells(ult, 2)).Value:
                if hasattr(linha[0], 'year'):
                    anos.add(linha[0].year)
        if not anos:
            import datetime
            anos = {datetime.date.today().year}
        anos = sorted(anos)
        campo.Validation.Delete()
        campo.Validation.Add(3, 1, 1, ','.join(str(a) for a in anos))
        campo.Validation.InCellDropdown = True
        campo.Value = anos[-1]
        print('I3: dropdown de ano recriado (%s) e valor = %d'
              % (','.join(str(a) for a in anos), anos[-1]))

        # ---- 2. rotina com guarda de ano plausivel -----------------------
        comp = None
        for c in wb.VBProject.VBComponents:
            try:
                if c.Type == 100 and c.Properties('Name').Value == 'Painel':
                    comp = c
            except Exception:
                pass
        if comp is None:
            raise SystemExit('modulo do Painel nao encontrado')
        cm = comp.CodeModule
        for nome in ('AplicarFiltroAnoMes', 'MesDoRotulo'):
            try:
                ini = cm.ProcStartLine(nome, 0)
                n = cm.ProcCountLines(nome, 0)
                if ini > 0:
                    cm.DeleteLines(ini, n)
            except Exception:
                pass
        cm.InsertLines(cm.CountOfLines + 1, '\r\n'.join([''] + ROTINA))
        txt = cm.Lines(1, cm.CountOfLines)
        if txt.count('Private Sub Worksheet_Change') != 1:
            raise SystemExit('Worksheet_Change duplicado apos o patch')
        if 'a < 1900' not in txt:
            raise SystemExit('guarda de ano nao entrou')
        print('AplicarFiltroAnoMes reescrita com guarda de ano plausivel')

        larg_depois = pa.Columns(1).ColumnWidth
        if abs(larg_antes - larg_depois) > 0.001:
            raise SystemExit('largura da coluna A mudou')
        print('largura da coluna A intacta: %s' % larg_depois)

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
