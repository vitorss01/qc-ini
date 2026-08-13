Attribute VB_Name = "mBI"
Option Explicit
' ===== CAMADA DE DADOS PARA BI (ADR-026) =====
'
' O QUE ESTA ABA E, E O QUE ELA NAO E
'
' BI_Data NAO e uma copia da planilha. E uma TABELA FATO: uma linha por
' resultado de controle, granularidade (Lote, RUN, Nivel, Analito), com chaves
' estaveis que nao dependem de posicao de celula. O Power BI le daqui e de
' mais lugar nenhum.
'
' O Excel continua sendo a fonte operacional. Esta camada nao calcula nada que
' o QC_INI ja nao calcule -- ela DESNORMALIZA o que o motor produz para que o
' BI nao precise reimplementar a regra. Onde ha conta, a conta e a mesma, e o
' teste de reconciliacao (mBI.ReconciliarComCalc) prova isso a cada build.
'
' DE ONDE VEM CADA COISA
'
'   resultado, data, status, RUN, nivel, lote .. DB_Resultados (A:G)
'   media e DP alvo POR LOTE ................... LotesStore, bloco do lote
'   CVtp / BIAStp / ETp ........................ Eng_Especificacoes
'   area e unidade ............................. Analitos (B, C)
'   Westgard ................................... recalculado aqui, identico a Calc
'
' ALVO POR LOTE, E NAO O ALVO DA TELA
'
' A aba Analitos mostra as metas do LOTE ATIVO -- e so dele. O historico dos
' demais lotes vive no LotesStore, em blocos de 40 linhas (mLotes.SalvarViewNoBloco).
' Ler o alvo da tela para um resultado de 2026 com o lote de 2027 carregado daria
' um Z errado, plausivel e silencioso. Aqui o alvo vem sempre do bloco do lote
' A QUE O RESULTADO PERTENCE.
'
' WESTGARD: AS REGRAS QUE EXISTEM DE VERDADE
'
' O QC_INI implementa TRES regras, e a copia aqui e literal (Calc!K3, L3, M3):
'
'   1_3s   ABS(Z) > 3                                   -- por resultado
'   2_2s   Z(N1) > 2 E Z(N2) > 2, ou ambos < -2         -- entre NIVEIS da mesma corrida
'   R_4s   Z(N1) > 2 E Z(N2) < -2, ou o inverso         -- entre NIVEIS da mesma corrida
'
' Calc!N3 e Calc!O3 sao "IF(OR(FALSE,FALSE),1,0)": 4_1s e 10x estao previstas na
' estrutura mas NAO implementadas. Esta camada nao inventa o que o motor nao
' calcula -- as colunas existem, sempre 0, documentadas como nao implementadas.
' Publicar um 4_1s inventado no painel do gestor seria pior do que nao ter.
'
' O QUE NAO EXISTE E POR ISSO NAO ESTA AQUI
'
' Equipamento e Setor. A producao nao tem as abas Corridas/Cfg_Status, e nenhuma
' outra guarda esses campos. Coluna sem origem e coluna que alguem vai filtrar e
' obter conclusao errada.

Public Const BI_ABA As String = "BI_Data"
Public Const BI_TABELA As String = "tblBI_Fato"
Public Const BI_CAB As Long = 1
Public Const BI_R0 As Long = 2
Public Const BI_NCOL As Long = 34

Private Const LS_CAP As Long = 40          ' linhas por bloco no LotesStore
Private Const LS_C0 As Long = 3            ' coluna C = Analitos!E (Media N1)

Private Function Cab() As Variant
    Cab = Array( _
        "ID_Result", "ID_Corrida", "Data", "Ano", "Mes", "Trimestre", "Competencia", _
        "ID_Analito", "Analito", "Area", "Unidade", _
        "ID_Lote", "Lote", "Nivel", "RUN", _
        "Resultado", "Status", "Ativo", _
        "Media_Alvo", "DP_Alvo", "Z", _
        "Lim_m3s", "Lim_m2s", "Lim_m1s", "Lim_p1s", "Lim_p2s", "Lim_p3s", _
        "CVtp_pct", "BIAStp_pct", "ETp_pct", _
        "W_1_3s", "W_2_2s", "W_R_4s", "Veredito")
End Function

' Indice do lote no registro de lotes (mesma conta de mLotes.BlocoDoLote, que e
' Private). Devolve 0 quando o lote nao esta registrado.
Private Function BlocoDoLoteBI(ByVal loteCore As String) As Long
    Dim rng As Range, c As Range, k As Long
    If Len(Trim$(loteCore)) = 0 Then Exit Function
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    For Each c In rng
        k = k + 1
        If Trim$(CStr(c.Value)) = Trim$(loteCore) Then BlocoDoLoteBI = k: Exit Function
    Next c
End Function

' Media e DP do analito NAQUELE lote, no nivel pedido.
' Devolve False quando nao ha alvo utilizavel -- e o chamador deixa Z vazio, em
' vez de dividir por zero e publicar um numero.
Private Function AlvoDoLote(ByVal iBloco As Long, ByVal iAnalito As Long, _
                            ByVal nivel As Long, ByRef media As Double, _
                            ByRef dp As Double) As Boolean
    Dim ws As Worksheet, r As Long, cM As Long
    Dim vM As Variant, vD As Variant
    If iBloco < 1 Or iAnalito < 1 Or nivel < 1 Or nivel > 3 Then Exit Function
    Set ws = ThisWorkbook.Sheets("LotesStore")
    r = 2 + (iBloco - 1) * LS_CAP + (iAnalito - 1)
    cM = LS_C0 + (nivel - 1) * 2          ' C/D = N1, E/F = N2, G/H = N3
    vM = ws.Cells(r, cM).Value
    vD = ws.Cells(r, cM + 1).Value
    If Not IsNumeric(vM) Or Not IsNumeric(vD) Then Exit Function
    If Len(Trim$(CStr(vM))) = 0 Or Len(Trim$(CStr(vD))) = 0 Then Exit Function
    media = CDbl(vM)
    dp = CDbl(vD)
    AlvoDoLote = (dp > 0)
End Function

' Reconstroi BI_Data inteira. O(n) sobre o banco, uma escrita em bloco.
Public Sub AtualizarBIData()
    Dim ws As Worksheet, wsA As Worksheet, wsE As Worksheet
    Dim dados As Variant, i As Long, n As Long, ult As Long
    Dim saida() As Variant, k As Long
    Dim idxAnalito As Object, idxBloco As Object, especCV As Object
    Dim especBIAS As Object, especET As Object, area As Object, unid As Object
    Dim zPorChave As Object
    Dim prot As Boolean, protEstava As Boolean

    Set wsA = ThisWorkbook.Sheets("Analitos")
    Set ws = GarantirAba()

    ' --- catalogo de analitos: indice na Analitos, area, unidade -----------
    Set idxAnalito = CreateObject("Scripting.Dictionary")
    Set area = CreateObject("Scripting.Dictionary")
    Set unid = CreateObject("Scripting.Dictionary")
    idxAnalito.CompareMode = 1: area.CompareMode = 1: unid.CompareMode = 1
    For i = 4 To 43
        Dim nm As String
        nm = Trim$(CStr(wsA.Cells(i, 1).Value))
        If Len(nm) > 0 Then
            If Not idxAnalito.Exists(nm) Then
                idxAnalito.Add nm, i - 3                  ' 1..40, alinhado ao LotesStore
                area.Add nm, CStr(wsA.Cells(i, 2).Value)
                unid.Add nm, CStr(wsA.Cells(i, 3).Value)
            End If
        End If
    Next i

    ' --- especificacoes vigentes (saida do motor do ADR-022) --------------
    Set especCV = CreateObject("Scripting.Dictionary")
    Set especBIAS = CreateObject("Scripting.Dictionary")
    Set especET = CreateObject("Scripting.Dictionary")
    especCV.CompareMode = 1: especBIAS.CompareMode = 1: especET.CompareMode = 1
    On Error Resume Next
    Set wsE = ThisWorkbook.Sheets("Eng_Especificacoes")
    On Error GoTo 0
    If Not wsE Is Nothing Then
        For i = 4 To 43
            Dim ne As String
            ne = Trim$(CStr(wsE.Cells(i, 1).Value))
            If Len(ne) > 0 Then
                If Not especCV.Exists(ne) Then
                    especCV.Add ne, wsE.Cells(i, 4).Value
                    especBIAS.Add ne, wsE.Cells(i, 5).Value
                    especET.Add ne, wsE.Cells(i, 6).Value
                End If
            End If
        Next i
    End If

    ' --- cache de blocos por lote ----------------------------------------
    Set idxBloco = CreateObject("Scripting.Dictionary")
    idxBloco.CompareMode = 1

    ' --- banco ------------------------------------------------------------
    ult = UltimaLinhaBanco()
    If ult < BANCO_R0 Then
        LimparCorpo ws
        Exit Sub
    End If
    dados = ThisWorkbook.Sheets(BANCO).Range( _
        ThisWorkbook.Sheets(BANCO).Cells(BANCO_R0, COL_RUN), _
        ThisWorkbook.Sheets(BANCO).Cells(ult, COL_STATUS)).Value
    n = UBound(dados, 1)
    ReDim saida(1 To n, 1 To BI_NCOL)

    ' PASSO 1 -- Z de cada resultado, guardado por (lote|run|analito|nivel).
    ' Westgard 2_2s e R_4s comparam NIVEIS DA MESMA CORRIDA, entao o Z do outro
    ' nivel precisa existir antes de julgar qualquer linha. Por isso duas
    ' passadas, e nao uma.
    Set zPorChave = CreateObject("Scripting.Dictionary")
    zPorChave.CompareMode = 1

    For i = 1 To n
        Dim an As String, lote As String, nucleo As String
        Dim nv As Long, run As Long, st As String
        an = Trim$(CStr(dados(i, COL_ANALITO)))
        If Len(an) = 0 Then GoTo proxima1
        st = Trim$(CStr(dados(i, COL_STATUS)))
        If st <> ST_ATIVO Then GoTo proxima1
        If Not IsNumeric(dados(i, COL_RESULT)) Then GoTo proxima1
        lote = Trim$(CStr(dados(i, COL_LOTE)))
        If Len(lote) < 6 Then GoTo proxima1
        nucleo = NucleoLote(lote)
        nv = CLng(Val(CStr(dados(i, COL_NIVEL))))
        run = CLng(Val(CStr(dados(i, COL_RUN))))

        Dim iB As Long
        If idxBloco.Exists(nucleo) Then
            iB = idxBloco(nucleo)
        Else
            iB = BlocoDoLoteBI(nucleo)
            idxBloco.Add nucleo, iB
        End If

        Dim md As Double, sd As Double
        If Not idxAnalito.Exists(an) Then GoTo proxima1
        If AlvoDoLote(iB, idxAnalito(an), nv, md, sd) Then
            zPorChave(nucleo & "|" & run & "|" & UCase$(an) & "|" & nv) = _
                (CDbl(dados(i, COL_RESULT)) - md) / sd
        End If
proxima1:
    Next i

    ' PASSO 2 -- monta a linha da tabela fato
    For i = 1 To n
        Dim an2 As String, lt As String, nuc As String, st2 As String
        Dim nv2 As Long, run2 As Long
        an2 = Trim$(CStr(dados(i, COL_ANALITO)))
        If Len(an2) = 0 Then GoTo proxima2
        lt = Trim$(CStr(dados(i, COL_LOTE)))
        nuc = ""
        If Len(lt) >= 6 Then nuc = NucleoLote(lt)
        st2 = Trim$(CStr(dados(i, COL_STATUS)))
        nv2 = CLng(Val(CStr(dados(i, COL_NIVEL))))
        run2 = CLng(Val(CStr(dados(i, COL_RUN))))

        k = k + 1
        saida(k, 1) = nuc & "|" & run2 & "|" & nv2 & "|" & UCase$(an2)   ' ID_Result
        saida(k, 2) = nuc & "|" & run2                                   ' ID_Corrida

        Dim dt As Date, temData As Boolean
        temData = IsDate(dados(i, COL_DATA))
        If temData Then
            dt = CDate(dados(i, COL_DATA))
            saida(k, 3) = Int(CDbl(dt))
            saida(k, 4) = Year(dt)
            saida(k, 5) = Month(dt)
            saida(k, 6) = Int((Month(dt) - 1) / 3) + 1
            saida(k, 7) = Format$(dt, "yyyy-mm")
        End If

        saida(k, 8) = UCase$(an2)
        saida(k, 9) = an2
        If area.Exists(an2) Then saida(k, 10) = area(an2)
        If unid.Exists(an2) Then saida(k, 11) = unid(an2)
        saida(k, 12) = nuc
        saida(k, 13) = lt
        saida(k, 14) = nv2
        saida(k, 15) = run2
        If IsNumeric(dados(i, COL_RESULT)) Then saida(k, 16) = CDbl(dados(i, COL_RESULT))
        saida(k, 17) = st2
        saida(k, 18) = IIf(st2 = ST_ATIVO, 1, 0)

        Dim md2 As Double, sd2 As Double, temAlvo As Boolean
        Dim iB2 As Long
        iB2 = 0
        If idxBloco.Exists(nuc) Then iB2 = idxBloco(nuc)
        temAlvo = False
        If idxAnalito.Exists(an2) Then
            temAlvo = AlvoDoLote(iB2, idxAnalito(an2), nv2, md2, sd2)
        End If
        If temAlvo Then
            saida(k, 19) = md2
            saida(k, 20) = sd2
            saida(k, 22) = md2 - 3 * sd2
            saida(k, 23) = md2 - 2 * sd2
            saida(k, 24) = md2 - sd2
            saida(k, 25) = md2 + sd2
            saida(k, 26) = md2 + 2 * sd2
            saida(k, 27) = md2 + 3 * sd2
        End If

        Dim chave As String, z As Variant, temZ As Boolean
        chave = nuc & "|" & run2 & "|" & UCase$(an2) & "|" & nv2
        temZ = zPorChave.Exists(chave)
        If temZ Then
            z = zPorChave(chave)
            saida(k, 21) = z
        End If

        If especCV.Exists(an2) Then saida(k, 28) = especCV(an2)
        If especBIAS.Exists(an2) Then saida(k, 29) = especBIAS(an2)
        If especET.Exists(an2) Then saida(k, 30) = especET(an2)

        ' --- Westgard, literalmente como Calc!K3, L3 e M3 -----------------
        '
        ' ZERAR A CADA LINHA, explicitamente.
        '
        ' Em VBA o Dim dentro de um laco NAO cria variavel nova a cada volta: o
        ' escopo e o procedimento inteiro e o valor sobrevive a iteracao. Sem
        ' estas tres atribuicoes, a primeira linha que viola uma regra contamina
        ' TODAS as seguintes -- a reconciliacao acusou 45 divergencias em 50
        ' pontos, todas no sentido "o motor diz OK e o BI diz REJEITADO".
        '
        ' O Z estava certo nas 50; so o veredito e que grudava. E o tipo de
        ' defeito que passaria despercebido numa conferencia por amostragem,
        ' porque o numero continua plausivel -- so que reprovando corrida boa.
        Dim w13 As Long, w22 As Long, wr4 As Long
        w13 = 0: w22 = 0: wr4 = 0
        If temZ Then
            If Abs(CDbl(z)) > 3 Then w13 = 1

            ' o OUTRO nivel da MESMA corrida: 2_2s e R_4s sao entre niveis
            Dim outro As Long, chaveO As String, zo As Variant, temZO As Boolean
            outro = IIf(nv2 = 1, 2, 1)
            chaveO = nuc & "|" & run2 & "|" & UCase$(an2) & "|" & outro
            temZO = zPorChave.Exists(chaveO)
            If temZO Then
                zo = zPorChave(chaveO)
                If (CDbl(z) > 2 And CDbl(zo) > 2) Or (CDbl(z) < -2 And CDbl(zo) < -2) Then w22 = 1
                If (CDbl(z) > 2 And CDbl(zo) < -2) Or (CDbl(z) < -2 And CDbl(zo) > 2) Then wr4 = 1
            End If
        End If
        saida(k, 31) = w13
        saida(k, 32) = w22
        saida(k, 33) = wr4
        If Len(Trim$(CStr(saida(k, 16)))) = 0 Then
            saida(k, 34) = ""
        ElseIf w13 + w22 + wr4 > 0 Then
            saida(k, 34) = "REJEITADO"
        Else
            saida(k, 34) = "OK"
        End If
proxima2:
    Next i

    ' --- escrita, com a protecao tratada (mesma licao do ADR-025) ---------
    prot = ws.ProtectContents
    On Error GoTo restaura
    If prot Then ws.Unprotect Password:="qcini2025"

    LimparCorpo ws
    If k > 0 Then
        Dim bloco() As Variant, r2 As Long, c2 As Long
        ReDim bloco(1 To k, 1 To BI_NCOL)
        For r2 = 1 To k
            For c2 = 1 To BI_NCOL
                bloco(r2, c2) = saida(r2, c2)
            Next c2
        Next r2
        ws.Range(ws.Cells(BI_R0, 1), ws.Cells(BI_R0 + k - 1, BI_NCOL)).Value = bloco
        ws.Range(ws.Cells(BI_R0, 3), ws.Cells(BI_R0 + k - 1, 3)).NumberFormat = "yyyy-mm-dd"
        ws.Range(ws.Cells(BI_R0, 13), ws.Cells(BI_R0 + k - 1, 13)).NumberFormat = "@"
    End If
    AjustarTabela ws, k

restaura:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    On Error Resume Next
    If prot Then
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If nErr <> 0 Then Err.Raise nErr, "mBI.AtualizarBIData", sErr
End Sub

Private Function GarantirAba() As Worksheet
    Dim ws As Worksheet, i As Long, cb As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BI_ABA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BI_ABA
    End If
    cb = Cab()
    For i = 0 To UBound(cb)
        ws.Cells(BI_CAB, i + 1).Value = cb(i)
    Next i
    ws.Rows(BI_CAB).Font.Bold = True
    Set GarantirAba = ws
End Function

Private Sub LimparCorpo(ByVal ws As Worksheet)
    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult >= BI_R0 Then
        ws.Range(ws.Cells(BI_R0, 1), ws.Cells(ult, BI_NCOL)).ClearContents
    End If
End Sub

' Tabela ESTRUTURADA, nao faixa de celulas.
'
' O Power Query referencia a tabela pelo NOME. Uma faixa A1:AH9999 quebraria a
' cada linha a mais ou a menos; o ListObject acompanha sozinho e o M nao muda.
Private Sub AjustarTabela(ByVal ws As Worksheet, ByVal nLinhas As Long)
    Dim lo As ListObject, alvo As Range
    Dim ultima As Long
    ultima = BI_R0 + IIf(nLinhas > 0, nLinhas - 1, 0)
    Set alvo = ws.Range(ws.Cells(BI_CAB, 1), ws.Cells(ultima, BI_NCOL))
    On Error Resume Next
    Set lo = ws.ListObjects(BI_TABELA)
    On Error GoTo 0
    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, alvo, , xlYes)
        lo.Name = BI_TABELA
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.Resize alvo
    End If
End Sub

' ---------------------------------------------------------------------------
' RECONCILIACAO: a camada BI tem de concordar com o motor, nao aproximar-se dele
'
' Compara, para o analito e o lote que a aba Calc esta exibindo, o Z e o veredito
' de Westgard linha a linha. Devolve "comparados|divergencias|primeira".
'
' Qualquer divergencia e defeito: as duas contas saem do mesmo dado e da mesma
' regra. Se divergirem, uma das duas esta errada -- e num sistema de CQ nao da
' para saber qual sem investigar.
Public Function ReconciliarComCalc() As String
    Dim wsC As Worksheet, wsB As Worksheet
    Dim selAn As String, lote As String
    Dim i As Long, ultB As Long, comp As Long, div As Long, prim As String
    Dim idx As Object, chave As String

    Set wsC = ThisWorkbook.Sheets("Calc")
    Set wsB = ThisWorkbook.Sheets(BI_ABA)
    selAn = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    lote = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
    If Len(selAn) = 0 Then ReconciliarComCalc = "0|0|sem analito selecionado": Exit Function

    ' indexa a BI_Data por RUN|NIVEL para o analito e lote em tela
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = 1
    ultB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    For i = BI_R0 To ultB
        If UCase$(Trim$(CStr(wsB.Cells(i, 9).Value))) = UCase$(selAn) Then
            If Trim$(CStr(wsB.Cells(i, 12).Value)) = lote Then
                chave = CStr(wsB.Cells(i, 15).Value) & "|" & CStr(wsB.Cells(i, 14).Value)
                If Not idx.Exists(chave) Then idx.Add chave, i
            End If
        End If
    Next i

    ' Calc: linha 3..182, coluna B = RUN, G = Z do N1, P = veredito N1,
    '                              AC = Z do N2, AL = veredito N2
    Dim nv As Long, colZ As Long, colV As Long
    For i = 3 To 182
        If Len(Trim$(CStr(wsC.Cells(i, 2).Value))) = 0 Then GoTo prox
        For nv = 1 To 2
            colZ = IIf(nv = 1, 7, 29)
            colV = IIf(nv = 1, 16, 38)
            If Len(Trim$(CStr(wsC.Cells(i, colZ).Value))) = 0 Then GoTo proxNivel
            chave = CStr(wsC.Cells(i, 2).Value) & "|" & CStr(nv)
            If Not idx.Exists(chave) Then GoTo proxNivel
            comp = comp + 1
            Dim zC As Double, zB As Double, vC As String, vB As String
            zC = CDbl(wsC.Cells(i, colZ).Value)
            zB = CDbl(wsB.Cells(idx(chave), 21).Value)
            vC = Trim$(CStr(wsC.Cells(i, colV).Value))
            vB = Trim$(CStr(wsB.Cells(idx(chave), 34).Value))
            If Abs(zC - zB) > 0.000001 Or vC <> vB Then
                div = div + 1
                If Len(prim) = 0 Then
                    prim = "RUN " & wsC.Cells(i, 2).Value & " N" & nv & _
                           " Z calc=" & Format$(zC, "0.000000") & " bi=" & Format$(zB, "0.000000") & _
                           " veredito calc=" & vC & " bi=" & vB
                End If
            End If
proxNivel:
        Next nv
prox:
    Next i
    ReconciliarComCalc = CStr(comp) & "|" & CStr(div) & "|" & prim
End Function
