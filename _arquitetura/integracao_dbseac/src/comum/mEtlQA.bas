Attribute VB_Name = "mEtlQA"
' ============================================================================
'  mEtlQA -- conferencia da integracao (tblQA_Integracao) em VBA (ADR-079)
'
'  Porte de pq/QA_INTEGRACAO.m (24 codigos: E01-E11, A01-A10, I01-I03), com as mesmas regras e os mesmos textos.
'  Le a tblCQ_Final JA CALCULADA em memoria (o PQ gastava 17-52 s so copiando 1,33 milhao de celulas da planilha
'  para o motor de consulta) e a MESMA leitura do DB_SEAC feita pelo recebimento (o PQ relia o arquivo inteiro).
'  Diferencas deliberadas, documentadas no ADR-079:
'    - dentro de um codigo, linhas por DATA_HORA (nulo primeiro) e, no empate, a ordem de entrada -- o PQ usava um
'      introsort instavel; o CONJUNTO de achados e o mesmo (quem le a tabela so conta por severidade);
'    - DATA_HORA nula no texto do A03/E06 vira "?" (no PQ derrubava a atualizacao inteira).
' ============================================================================
Option Explicit
Option Compare Binary

Private Const SIM As String = "SIM"
Private Const NAO As String = "NÃO"

' colunas da tblCQ_Final (ordem de mEtlCQ.NomesColunas)
Private Const Q_ID As Long = 1, Q_OR As Long = 2, Q_EQ As Long = 4, Q_MZ As Long = 5, Q_DT As Long = 7, Q_DH As Long = 9
Private Const Q_COR As Long = 10, Q_RUN As Long = 11, Q_NV As Long = 12, Q_LT As Long = 14, Q_AN As Long = 15
Private Const Q_CAD As Long = 17, Q_ST As Long = 21, Q_PART As Long = 22, Q_INATR As Long = 25, Q_JUST As Long = 27
Private Const Q_MOTEX As Long = 32, Q_REL As Long = 33

Private mOut As Collection      ' cada item: Array(sev, cod, teste, id, analito, nivel, data_hora, detalhe)
Private mPrefixo As String
Private mTempos As String
Private mT0 As Double

Public Function TemposUltimoQA() As String
    TemposUltimoQA = mTempos
End Function

Private Sub MarcaQ(ByVal fase As String)
    mTempos = mTempos & fase & "=" & Format$(Timer - mT0, "0.00") & ";"
    mT0 = Timer
End Sub

Public Function NomesColunasQA() As Variant
    NomesColunasQA = Array("SEVERIDADE", "CODIGO", "TESTE", "ID_REGISTRO", "ANALITO", "NIVEL", "DATA_HORA", "DETALHE")
End Function

' ---------------------------------------------------------------- utilitarios (mesma semantica do mEtlCQ)
Private Function Vazio(ByVal v As Variant) As Boolean
    Vazio = IsEmpty(v) Or IsNull(v)
End Function

Private Function TrimU(ByVal s As String) As String
    Dim i As Long, j As Long, c As Long
    i = 1: j = Len(s)
    Do While i <= j
        c = AscW(Mid$(s, i, 1)): If c < 0 Then c = c + 65536
        If Not (c = 32 Or (c >= 9 And c <= 13) Or c = 133 Or c = 160 Or c = 5760 Or (c >= 8192 And c <= 8202) Or _
                c = 8232 Or c = 8233 Or c = 8239 Or c = 8287 Or c = 12288) Then Exit Do
        i = i + 1
    Loop
    Do While j >= i
        c = AscW(Mid$(s, j, 1)): If c < 0 Then c = c + 65536
        If Not (c = 32 Or (c >= 9 And c <= 13) Or c = 133 Or c = 160 Or c = 5760 Or (c >= 8192 And c <= 8202) Or _
                c = 8232 Or c = 8233 Or c = 8239 Or c = 8287 Or c = 12288) Then Exit Do
        j = j - 1
    Loop
    If j >= i Then TrimU = Mid$(s, i, j - i + 1)
End Function

Private Function TF(ByVal x As Variant) As Variant          ' Text.From (pt-BR)
    If Vazio(x) Then TF = Empty: Exit Function
    Select Case VarType(x)
        Case vbString: TF = x
        Case vbBoolean: TF = IIf(x, "true", "false")
        Case vbDate: TF = Format$(x, "dd/mm/yyyy hh:nn:ss")
        Case Else
            If IsNumeric(x) Then
                If CDbl(x) = Int(CDbl(x)) And Abs(CDbl(x)) < 1E+15 Then TF = Format$(CDbl(x), "0") Else TF = CStr(CDbl(x))
            Else
                TF = CStr(x)
            End If
    End Select
End Function

Private Function Txt(ByVal x As Variant) As Variant
    Dim t As String
    If Vazio(x) Then Txt = Empty: Exit Function
    t = TrimU(CStr(TF(x)))
    If Len(t) = 0 Then Txt = Empty Else Txt = t
End Function

Private Function TxtOu(ByVal x As Variant, ByVal padrao As String) As String
    If Vazio(Txt(x)) Then TxtOu = padrao Else TxtOu = Txt(x)
End Function

Private Function SoDig(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then Exit Function
    Next i
    SoDig = True
End Function

Private Function SemZ(ByVal d As String) As String
    Dim i As Long
    i = 1
    Do While i <= Len(d)
        If Mid$(d, i, 1) <> "0" Then Exit Do
        i = i + 1
    Loop
    If i > Len(d) Then SemZ = "0" Else SemZ = Mid$(d, i)
End Function

Private Function NormIdQ(ByVal x As Variant) As Variant
    Dim t As Variant, u As String, pre As String, dg As String
    t = Txt(x)
    If IsEmpty(t) Then NormIdQ = Empty: Exit Function
    u = UCase$(Replace(Replace(CStr(t), " ", ""), ChrW$(160), ""))
    pre = mPrefixo & "-"
    If SoDig(u) Then
        NormIdQ = pre & SemZ(u)
    ElseIf Left$(u, Len(pre)) = pre And SoDig(Mid$(u, Len(pre) + 1)) Then
        NormIdQ = pre & SemZ(Mid$(u, Len(pre) + 1))
    ElseIf Left$(u, 4) = "MAN_" And SoDig(Mid$(u, 5)) Then
        dg = Mid$(u, 5)
        If Len(dg) < 4 Then dg = String$(4 - Len(dg), "0") & dg
        NormIdQ = "MAN_" & dg
    Else
        NormIdQ = u
    End If
End Function

Private Function AnKeyQ(ByVal x As Variant) As Variant
    If IsEmpty(Txt(x)) Then AnKeyQ = Empty Else AnKeyQ = UCase$(Txt(x))
End Function

Private Function DH(ByVal v As Variant) As String              ' 'dd/MM/yyyy HH:mm' (nulo -> "?")
    If Vazio(v) Then DH = "?" Else DH = Format$(CDate(CDbl(v)), "dd/mm/yyyy hh:nn")
End Function

Private Function Dia(ByVal v As Variant) As String
    If Vazio(v) Then Dia = "?" Else Dia = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
End Function

Private Function Num(ByVal v As Variant) As Variant             ' serial/numero ou Empty
    If Vazio(v) Then
        Num = Empty
    ElseIf VarType(v) = vbDate Then
        Num = CDbl(v)
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        Num = CDbl(v)
    Else
        Num = Empty
    End If
End Function

Private Sub Achado(ByVal sev As String, ByVal cod As String, ByVal teste As String, ByVal id As Variant, ByVal an As Variant, _
                   ByVal nv As Variant, ByVal dataHora As Variant, ByVal det As Variant)
    If Not Vazio(nv) Then If IsNumeric(nv) Then nv = CDbl(Round(CDbl(nv), 0))
    mOut.Add Array(sev, cod, teste, id, an, nv, Num(dataHora), det)
End Sub

' ---------------------------------------------------------------- a conferencia
' F = tblCQ_Final calculada (mEtlCQ.CalcularCQFinal); rec/cabRec = recebimento (layout da tblDB_Recebimento);
' orig/nomesOrig/nOrig = a tabela do DB_SEAC lida nesta atualizacao (nOrig = 0 quando nao lida); historico = MODO.
Public Function CalcularQA(ByRef F As Variant, ByRef rec As Variant, ByVal nRec As Long, ByVal cabRec As Object, _
                           ByRef orig As Variant, ByVal nomesOrig As Variant, ByVal nOrig As Long, _
                           ByVal historico As Boolean) As Variant
    Dim nF As Long, i As Long, j As Long, k As Variant, C As Object, cab As Object, dI As Variant, nI As Long, dK As Variant, nK As Long
    Dim lo As ListObject, h As Variant, s As String
    Set mOut = New Collection
    mTempos = "": mT0 = Timer
    If IsEmpty(F) Then nF = 0 Else nF = UBound(F, 1)

    ' CFG
    Set C = CreateObject("Scripting.Dictionary")
    Set lo = TabelaQ("tblConfigIntegracao")
    If Not lo Is Nothing Then
        If lo.ListRows.Count > 0 Then
            h = lo.DataBodyRange.Value
            Set cab = Cabecalho(lo)
            For i = 1 To UBound(h, 1)
                If Not IsEmpty(Txt(h(i, cab("CHAVE")))) Then
                    If VarType(h(i, cab("VALOR"))) = vbString Then C(UCase$(Txt(h(i, cab("CHAVE"))))) = TrimU(h(i, cab("VALOR"))) _
                    Else C(UCase$(Txt(h(i, cab("CHAVE"))))) = h(i, cab("VALOR"))
                End If
            Next i
        End If
    End If
    mPrefixo = UCase$(TrimU(CStr(TF(C("PREFIXO_ID")))))

    MarcaQ "cfg"
    ' ---- E01
    For i = 1 To nF
        If F(i, Q_ST) = "INATIVADO" And F(i, Q_JUST) <> SIM Then
            If Vazio(F(i, Q_DH)) Then s = Empty Else s = "ERRO DE GOVERNANÇA: " & F(i, Q_ID) & " | " & TxtOu(F(i, Q_AN), "?") & _
                " | nivel " & TxtOu(F(i, Q_NV), "?") & " | " & DH(F(i, Q_DH)) & " -- inativado sem justificativa; escrever o MOTIVO na linha da aba Inativar"
            Achado "ERRO", "E01", "Resultado inativado sem justificativa tecnica", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), IIf(Vazio(F(i, Q_DH)), Empty, s)
        End If
    Next i

    MarcaQ "E01"
    ' ---- inativacao: E02, E03, E10, E11, A08
    Dim cI As Object, legado As Boolean, gL As Object, ordemI As Collection, idn As Variant, idsF As Object, porIdF As Object
    Dim nIdx As Long
    Set lo = TabelaQ("tblInativacao_NaoConformes")
    Set cI = Cabecalho(lo)
    legado = Not cI.Exists("MOTIVO")
    nI = 0
    If Not lo Is Nothing Then If lo.ListRows.Count > 0 Then dI = lo.DataBodyRange.Value: nI = lo.ListRows.Count
    Set gL = CreateObject("Scripting.Dictionary")
    Set ordemI = New Collection
    For i = 1 To nI
        idn = NormIdQ(dI(i, cI("ID_REGISTRO")))
        If Not IsEmpty(idn) Then
            If Not gL.Exists(idn) Then gL.Add idn, CStr(i): ordemI.Add Array(idn, i) Else gL(idn) = gL(idn) & ", " & i
        End If
    Next i
    For Each k In gL.Keys
        If InStr(gL(k), ",") > 0 Then
            Achado "ERRO", "E02", "ID_REGISTRO duplicado na tabela de inativacao", k, Empty, Empty, Empty, _
                "aparece " & (UBound(Split(gL(k), ",")) + 1) & " vezes (linhas " & gL(k) & " da tabela); vale so a primeira, a contagem nao duplica"
        End If
    Next k
    ' linhas da Principal de cada ID INATIVADO (so os IDs da aba Inativar entram no indice)
    Set porIdF = CreateObject("Scripting.Dictionary")
    If ordemI.Count > 0 Then nIdx = nF Else nIdx = 0           ' sem inativacao, o indice nao e usado
    For i = 1 To nIdx
        If Not Vazio(F(i, Q_ID)) Then
            If gL.Exists(F(i, Q_ID)) Then
                If Not porIdF.Exists(F(i, Q_ID)) Then porIdF.Add F(i, Q_ID), CStr(i) Else porIdF(F(i, Q_ID)) = porIdF(F(i, Q_ID)) & "," & i
            End If
        End If
    Next i
    Set idsF = porIdF
    Dim par As Variant, f_ As Variant, anDig As Variant
    For Each par In ordemI
        If Not idsF.Exists(par(0)) Then
            Achado "ERRO", "E03", "ID inativado que nao existe na base", par(0), Empty, Empty, Empty, _
                "ID digitado na linha " & par(1) & " da inativacao nao corresponde a nenhum resultado recebido ou manual"
        End If
    Next par
    If Not legado Then
        For Each par In ordemI
            If porIdF.Exists(par(0)) Then
                anDig = dI(par(1), cI("ANALITO"))
                For Each f_ In Split(porIdF(par(0)), ",")
                    If Not IsEmpty(AnKeyQ(anDig)) Then
                        If AnKeyQ(anDig) <> AnKeyQ(F(f_, Q_AN)) Or IsEmpty(AnKeyQ(F(f_, Q_AN))) Then
                            Achado "ERRO", "E10", "ID nao pertence ao analito informado na inativacao", F(f_, Q_ID), F(f_, Q_AN), F(f_, Q_NV), F(f_, Q_DH), _
                                "linha " & par(1) & " da aba Inativar: o ID e de " & TxtOu(F(f_, Q_AN), "?") & " (nivel " & TxtOu(F(f_, Q_NV), "?") & _
                                ", RUN " & TxtOu(F(f_, Q_RUN), "?") & "), mas a linha diz '" & TF(anDig) & _
                                "'. A inativacao NAO foi aplicada: confira o ID no grafico e corrija o ID ou o analito."
                        End If
                    Else
                        Achado "ERRO", "E11", "Inativacao sem analito informado", F(f_, Q_ID), F(f_, Q_AN), F(f_, Q_NV), F(f_, Q_DH), _
                            "linha " & par(1) & " da aba Inativar sem ANALITO: o ID e de " & TxtOu(F(f_, Q_AN), "?") & " (nivel " & _
                            TxtOu(F(f_, Q_NV), "?") & ", RUN " & TxtOu(F(f_, Q_RUN), "?") & "). A inativacao NAO foi aplicada: escolha o analito na linha."
                    End If
                Next f_
            End If
        Next par
    End If

    MarcaQ "inativ"
    ' ---- E04, E05, A01, A02, A04, E09, A09 (linhas da propria tabela)
    Dim idsInat As Object
    Set idsInat = CreateObject("Scripting.Dictionary")
    For i = 1 To nF
        If F(i, Q_ST) = "INATIVADO" Then idsInat(F(i, Q_ID)) = 1
    Next i
    For i = 1 To nF
        If F(i, Q_ST) = "MANUAL_INCOMPLETO" Then Achado "ERRO", "E04", "Resultado manual com campo obrigatorio ausente", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), F(i, Q_MOTEX)
    Next i
    For i = 1 To nF
        If Not Vazio(F(i, Q_COR)) Then If F(i, Q_COR) > 99 Then Achado "ERRO", "E05", "Mais de 99 corridas no mesmo dia", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), "RUN nao atribuido; revisar o dado do dia"
    Next i
    For i = 1 To nF
        If F(i, Q_ST) = "CONFLITO_MANUAL" Then
            If Vazio(F(i, Q_MOTEX)) Then s = Empty Else s = F(i, Q_MOTEX) & IIf(Vazio(F(i, Q_REL)), "", " -> " & F(i, Q_REL)) & _
                ". O manual ficou FORA da estatistica; confirme e inative um deles ou corrija o manual."
            Achado "ALERTA", "A01", "Possivel duplicidade de resultado manual", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), IIf(Vazio(F(i, Q_MOTEX)), Empty, s)
        End If
    Next i
    For i = 1 To nF
        If F(i, Q_ST) = "DUPLICIDADE_ORIGEM" Then Achado "ALERTA", "A02", "Retransmissao do interfaceamento (duplicidade de origem)", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), _
            "copia de " & TxtOu(F(i, Q_REL), "?") & " (mesma amostra, item, instante e valor) -- nao entra em calculo nem no grafico"
    Next i

    MarcaQ "linhas"
    ' ---- A03, E06, E07, E08 (origem x recebimento)
    Dim cO As Object, cR As Object, recPorId As Object, ido As Variant, idr As Variant, d As String, vr As Variant, va As Variant
    Set cO = CreateObject("Scripting.Dictionary")
    If nOrig > 0 Then
        For j = 1 To UBound(nomesOrig): cO(CStr(nomesOrig(j))) = j: Next j
    End If
    Set cR = cabRec
    Set recPorId = CreateObject("Scripting.Dictionary")
    For i = 1 To nRec
        idr = rec(i, cR("ID_REGISTRO"))
        If Not Vazio(idr) Then
            If Not recPorId.Exists(idr) Then recPorId.Add idr, CStr(i) Else recPorId(idr) = recPorId(idr) & "," & i
        End If
    Next i
    Dim r_ As Variant, dl As Collection, item As Variant, ixR() As Long, ixO() As Long, cIdo As Long, cc As Variant
    If Not historico Then
        ' colunas resolvidas uma vez (o caminho rapido roda para cada linha da origem)
        ReDim ixR(1 To 6): ReDim ixO(1 To 6)
        j = 0
        For Each cc In Array("VALOR", "FLAG", "ANALITO", "LOTE", "NIVEL", "DATA_HORA")
            j = j + 1
            If cR.Exists(cc) Then ixR(j) = cR(cc)
            If cO.Exists(cc) Then ixO(j) = cO(cc)
        Next cc
        If nOrig > 0 Then cIdo = cO("ID_ORIGEM")
        For i = 1 To nOrig
            ido = orig(i, cIdo)
            If Not Vazio(ido) And IsNumeric(ido) Then
                idr = mPrefixo & "-" & Format$(CDbl(Round(CDbl(ido), 0)), "0")
                If recPorId.Exists(idr) Then
                    For Each r_ In Split(recPorId(idr), ",")
                        If BrutoIgual(rec, CLng(r_), ixR, orig, i, ixO) Then GoTo proxR
                        Set dl = New Collection
                        vr = Num(rec(r_, cR("VALOR"))): va = Num(orig(i, cO("VALOR")))
                        If Not IgualN(vr, va) Then dl.Add "VALOR " & IIf(Vazio(rec(r_, cR("VALOR"))), "vazio", TF(rec(r_, cR("VALOR")))) & " -> " & IIf(Vazio(va), "vazio", TF(va))
                        If Not IgualT(Txt(ValCab(rec, r_, cR, "FLAG")), Txt(ValCab(orig, i, cO, "FLAG"))) Then dl.Add "FLAG"
                        If Not IgualT(Txt(rec(r_, cR("ANALITO"))), Txt(orig(i, cO("ANALITO")))) Then dl.Add "ANALITO " & TxtOu(rec(r_, cR("ANALITO")), "vazio") & " -> " & TxtOu(orig(i, cO("ANALITO")), "vazio")
                        If Not IgualT(Txt(rec(r_, cR("LOTE"))), Txt(orig(i, cO("LOTE")))) Then dl.Add "LOTE " & TxtOu(rec(r_, cR("LOTE")), "vazio") & " -> " & TxtOu(orig(i, cO("LOTE")), "vazio")
                        If Not IgualN(Num(rec(r_, cR("NIVEL"))), Num(orig(i, cO("NIVEL")))) Then dl.Add "NIVEL"
                        If Not IgualN(Segundos(rec(r_, cR("DATA_HORA"))), Segundos(orig(i, cO("DATA_HORA")))) Then dl.Add "DATA_HORA " & DH(Num(rec(r_, cR("DATA_HORA")))) & " -> " & DH(Num(orig(i, cO("DATA_HORA"))))
                        If dl.Count > 0 Then
                            d = ""
                            For Each item In dl: d = d & IIf(Len(d) > 0, "; ", "") & item: Next item
                            Achado "ALERTA", "A03", "Valor na origem diferente do valor recebido", idr, orig(i, cO("ANALITO")), orig(i, cO("NIVEL")), orig(i, cO("DATA_HORA")), _
                                "mudou na origem: " & d & ". Vale o RECEBIDO (registro imutavel); investigar a correcao na origem."
                        End If
proxR:
                    Next r_
                End If
            End If
        Next i
        For i = 1 To nOrig
            ido = orig(i, cIdo)
            If Vazio(ido) Or Not IsNumeric(ido) Then
                Achado "ERRO", "E06", "Linha da origem sem ID (rejeitada)", Empty, orig(i, cO("ANALITO")), orig(i, cO("NIVEL")), orig(i, cO("DATA_HORA")), _
                    "linha do DB_SEAC sem ID valido -- NAO entrou no recebimento: ID_AMOSTRA " & TxtOu(ValCab(orig, i, cO, "ID_AMOSTRA"), "?") & " | " & _
                    TxtOu(orig(i, cO("ANALITO")), "?") & " | nivel " & TxtOu(orig(i, cO("NIVEL")), "?") & " | " & DH(Num(orig(i, cO("DATA_HORA")))) & _
                    " | ID_ORIGEM informado: " & TxtOu(ido, "vazio") & ". Corrigir o ID na origem (SIPEC/DB_SEAC)."
            End If
        Next i
    End If
    MarcaQ "A03E06"
    ' E07: ID repetido na origem, depois no recebimento (o 1o de cada ID). Conta primeiro (barato); detalhe so dos
    ' repetidos, na ordem da 1a aparicao.
    Dim vistos7 As Object, grp As Object, ordem7 As Collection, onde As Variant, cnt As Object, idsLista() As String, nL As Long
    Set vistos7 = CreateObject("Scripting.Dictionary")
    For Each onde In Array("na origem (DB_SEAC)", "no recebimento")
        Set grp = CreateObject("Scripting.Dictionary")
        Set ordem7 = New Collection
        Set cnt = CreateObject("Scripting.Dictionary")
        If onde = "no recebimento" Then nL = nRec Else If historico Then nL = 0 Else nL = nOrig
        ReDim idsLista(1 To IIf(nL > 0, nL, 1))
        For i = 1 To nL
            If onde = "no recebimento" Then
                idr = rec(i, cR("ID_REGISTRO"))
                If Vazio(idr) Then idsLista(i) = "" Else idsLista(i) = CStr(idr)
            Else
                ido = orig(i, cO("ID_ORIGEM"))
                If Vazio(ido) Or Not IsNumeric(ido) Then idsLista(i) = "" Else idsLista(i) = mPrefixo & "-" & Format$(CDbl(Round(CDbl(ido), 0)), "0")
            End If
            If Len(idsLista(i)) > 0 Then cnt(idsLista(i)) = cnt(idsLista(i)) + 1
        Next i
        If cnt.Count < nL Then
            For i = 1 To nL
                If Len(idsLista(i)) > 0 Then
                    If cnt(idsLista(i)) > 1 Then
                        If onde = "no recebimento" Then
                            GrupoE07 grp, ordem7, idsLista(i), rec(i, cR("ANALITO")), rec(i, cR("NIVEL")), Num(rec(i, cR("DATA_HORA")))
                        Else
                            GrupoE07 grp, ordem7, idsLista(i), orig(i, cO("ANALITO")), orig(i, cO("NIVEL")), Num(orig(i, cO("DATA_HORA")))
                        End If
                    End If
                End If
            Next i
        End If
        For Each k In ordem7
            item = grp(k)
            If item(0) > 1 And Not vistos7.Exists(k) Then
                vistos7.Add k, 1
                Achado "ERRO", "E07", "ID repetido na origem", k, item(1), item(2), item(3), _
                    "o ID aparece " & item(0) & " vezes " & onde & "; entrou UMA vez (a de menor ITEM_ID). Conferir a origem: dois resultados nunca podem ter o mesmo ID."
            End If
        Next k
    Next onde
    MarcaQ "E07"
    ' E08: prefixo do ID recebido diferente do CFG
    Dim pg As Object, pre As String
    Set pg = CreateObject("Scripting.Dictionary")
    For i = 1 To nRec
        idr = rec(i, cR("ID_REGISTRO"))
        If Not Vazio(idr) Then
            If Left$(UCase$(CStr(TF(idr))), Len(mPrefixo) + 1) <> mPrefixo & "-" Then
                pre = Split(CStr(TF(idr)), "-")(0)
                If Not pg.Exists(pre) Then pg.Add pre, Array(1, CStr(TF(idr))) Else pg(pre) = Array(pg(pre)(0) + 1, IIf(CStr(TF(idr)) < pg(pre)(1), CStr(TF(idr)), pg(pre)(1)))
            End If
        End If
    Next i
    For Each k In pg.Keys
        Achado "ERRO", "E08", "Prefixo do ID recebido diferente do CFG", pg(k)(1), Empty, Empty, Empty, _
            pg(k)(0) & " ID(s) recebido(s) com prefixo '" & k & "-' e o CFG PREFIXO_ID = '" & mPrefixo & _
            "'. Conferir o CFG: com o prefixo trocado, todo o DB_SEAC entraria de novo como resultado novo."
    Next k

    MarcaQ "E08"
    ' ---- A04
    For i = 1 To nF
        If F(i, Q_INATR) = SIM And F(i, Q_ST) <> "INATIVADO" Then Achado "ALERTA", "A04", "Inativacao redundante", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), _
            "o resultado ja estava fora da estatistica por " & F(i, Q_ST) & "; a inativacao nao muda nada"
    Next i

    ' ---- A05: comentario de resultado nao inativado (1o comentario de cada ID)
    Dim cK As Object, inatReg As Object, vistosK As Object
    Set lo = TabelaQ("tblComentariosTecnicos")
    Set cK = Cabecalho(lo)
    nK = 0
    If Not lo Is Nothing Then If lo.ListRows.Count > 0 Then dK = lo.DataBodyRange.Value: nK = lo.ListRows.Count
    Set inatReg = CreateObject("Scripting.Dictionary")
    For i = 1 To nF
        If F(i, Q_INATR) = SIM Then inatReg(F(i, Q_ID)) = 1
    Next i
    Set vistosK = CreateObject("Scripting.Dictionary")
    For i = 1 To nK
        idn = NormIdQ(dK(i, cK("ID_REGISTRO")))
        If Not IsEmpty(idn) And Not IsEmpty(Txt(dK(i, cK("COMENTARIO_TECNICO")))) Then
            If Not vistosK.Exists(idn) Then
                vistosK.Add idn, 1
                If Not inatReg.Exists(idn) Then Achado "INFO", "A05", "Comentario tecnico de resultado nao inativado", idn, Empty, Empty, Empty, _
                    "comentario mantido como historico (resultado reativado ou ID nunca inativado)"
            End If
        End If
    Next i

    MarcaQ "A04A05"
    ' ---- I01, I02 (por analito, na ordem da 1a aparicao)
    GrupoAnalito F, nF, Q_CAD, NAO, "I01", "Analito recebido sem cadastro no QC_INI", _
        " resultado(s) guardados na base; nao aparecem no Painel ate o analito ser cadastrado/mapeado"
    GrupoAnalito F, nF, Q_ST, "SEM_VALOR", "I02", "Resultado sem valor numerico (flag do equipamento)", _
        " resultado(s) com flag e sem numero; guardados, fora de calculo e grafico"

    MarcaQ "I01I02"
    ' ---- A06: manual ATIVO e interfaceamento (ATIVO/INATIVADO) com a mesma chave no dia
    Dim intDia As Object, k6 As String, lista As String, partes() As String, nPar As Long
    Set intDia = CreateObject("Scripting.Dictionary")
    Dim temManual As Boolean, nA6 As Long
    For i = 1 To nF
        If F(i, Q_OR) = "MANUAL" Then If F(i, Q_ST) = "ATIVO" Then temManual = True: Exit For
    Next i
    If temManual Then nA6 = nF Else nA6 = 0
    For i = 1 To nA6
        If F(i, Q_OR) = "INTERFACEAMENTO" And (F(i, Q_ST) = "ATIVO" Or F(i, Q_ST) = "INATIVADO") Then
            k6 = Chave6(F, i)
            If intDia.Exists(k6) Then intDia(k6) = intDia(k6) & Chr$(0) & F(i, Q_ID) Else intDia.Add k6, CStr(F(i, Q_ID))
        End If
    Next i
    For i = 1 To nA6
        If F(i, Q_OR) = "MANUAL" And F(i, Q_ST) = "ATIVO" Then
            k6 = Chave6(F, i)
            If intDia.Exists(k6) Then
                partes = Split(intDia(k6), Chr$(0))
                lista = ""
                For nPar = 0 To IIf(UBound(partes) > 4, 4, UBound(partes))
                    lista = lista & IIf(nPar > 0, ", ", "") & partes(nPar)
                Next nPar
                Achado "ALERTA", "A06", "Resultado manual e do interfaceamento no mesmo dia (fora da tolerancia)", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), _
                    "mesmo equipamento/lote/nivel/analito no dia: " & lista & ". Os dois participam (corridas diferentes). Se forem o MESMO resultado, inative o manual."
            End If
        End If
    Next i

    MarcaQ "A06"
    ' ---- I03: lote recebido sem cadastro
    Dim lotesCad As Object, regs As Variant, t As Variant, pl As Object, ordemL As Collection, kl As String, info As Variant
    Set lotesCad = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    regs = ThisWorkbook.Names("regLoteCol").RefersToRange.Value
    On Error GoTo 0
    If IsArray(regs) Then
        For i = 1 To UBound(regs, 1)
            t = Txt(regs(i, 1))
            If Not IsEmpty(t) Then
                If VarType(regs(i, 1)) = vbDouble Then lotesCad(CStr(CDbl(Round(regs(i, 1), 0)))) = 1 Else lotesCad(t) = 1
            End If
        Next i
    End If
    Set pl = CreateObject("Scripting.Dictionary")
    Set ordemL = New Collection
    For i = 1 To nF
        If F(i, Q_CAD) = SIM And F(i, Q_OR) = "INTERFACEAMENTO" Then
            kl = KT(F(i, Q_EQ)) & Chr$(0) & KT(F(i, Q_LT))
            If Not pl.Exists(kl) Then
                pl.Add kl, Array(F(i, Q_EQ), F(i, Q_LT), 0, Empty, Empty)
                ordemL.Add kl
            End If
            info = pl(kl)
            info(2) = info(2) + 1
            If Not Vazio(F(i, Q_DH)) Then
                If IsEmpty(info(3)) Then info(3) = F(i, Q_DH) Else If F(i, Q_DH) < info(3) Then info(3) = F(i, Q_DH)
                If IsEmpty(info(4)) Then info(4) = F(i, Q_DH) Else If F(i, Q_DH) > info(4) Then info(4) = F(i, Q_DH)
            End If
            pl(kl) = info
        End If
    Next i
    For Each k In ordemL
        info = pl(k)
        If Not lotesCad.Exists(CStr(TF(info(1)))) Then
            Achado "INFO", "I03", "Lote recebido sem cadastro no QC_INI", Empty, Empty, Empty, info(4), _
                "lote " & TF(info(1)) & " (" & TF(info(0)) & "): " & info(2) & " resultado(s) de " & Dia(info(3)) & " a " & Dia(info(4)) & _
                ". Cadastre o lote (Configuracao) e a media/DP para ve-lo no Painel."
        End If
    Next k

    MarcaQ "I03"
    ' ---- A07: modo historico
    Dim ult As Variant
    If historico Then
        For i = 1 To nRec
            If Not Vazio(Num(rec(i, cR("DATA_HORA")))) Then
                If IsEmpty(ult) Then ult = Num(rec(i, cR("DATA_HORA"))) Else If Num(rec(i, cR("DATA_HORA"))) > ult Then ult = Num(rec(i, cR("DATA_HORA")))
            End If
        Next i
        Achado "ALERTA", "A07", "MODO HISTORICO: DB_SEAC nao lido nesta atualizacao", Empty, Empty, Empty, ult, _
            "nenhum resultado novo foi recebido; vale o historico ate " & DH(ult) & ". Na rede do laboratorio, voltar Cfg_Integracao MODO_FONTE = SEAC."
    End If

    ' ---- E09
    For i = 1 To nF
        If F(i, Q_OR) = "INTERFACEAMENTO" And Vazio(F(i, Q_DH)) Then Achado "ERRO", "E09", "Resultado do interfaceamento sem DATA_HORA", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), _
            "sem DATA_HORA na origem: fora da corrida (RUN vazio), do grafico e da estatistica. Corrigir na origem."
    Next i

    ' ---- A08: REGISTRAR - LJ com valor nao reconhecido
    Dim lj As Variant
    For i = 1 To nI
        idn = NormIdQ(dI(i, cI("ID_REGISTRO")))
        If Not IsEmpty(idn) And cI.Exists("REGISTRAR - LJ") Then
            lj = dI(i, cI("REGISTRAR - LJ"))
            If Not (Vazio(lj) Or VarType(lj) = vbBoolean Or (IsNumeric(lj) And VarType(lj) <> vbString)) Then
                Select Case UCase$(TrimU(CStr(TF(lj))))
                    Case "", "SIM", "S", "TRUE", "VERDADEIRO", "V", "X", "1", NAO, "NAO", "N", "FALSE", "FALSO", "F", "0"
                    Case Else
                        Achado "ALERTA", "A08", "REGISTRAR - LJ com valor nao reconhecido", idn, IIf(cI.Exists("ANALITO"), dI(i, cI("ANALITO")), Empty), Empty, Empty, _
                            "valor '" & TF(lj) & "' na inativacao: tratado como NAO (o X nao aparece). Marque ou desmarque a caixa."
                End Select
            End If
        End If
    Next i

    ' ---- A09: manual em conflito com um interfaceamento INATIVADO
    For i = 1 To nF
        If F(i, Q_OR) = "MANUAL" And F(i, Q_ST) = "CONFLITO_MANUAL" And Not Vazio(F(i, Q_REL)) Then
            If idsInat.Exists(F(i, Q_REL)) Then Achado "ALERTA", "A09", "Manual coincide com resultado do interfaceamento inativado", F(i, Q_ID), F(i, Q_AN), F(i, Q_NV), F(i, Q_DH), _
                "o automatico " & F(i, Q_REL) & " esta INATIVADO e este manual tem a mesma chave e instante: o manual NAO participa: " & _
                "resultado manual nunca substitui um automatico (decisao do laboratorio, 06/10/2026)."
        End If
    Next i

    ' ---- A10: Hematologia, resultado de outro equipamento
    Dim eqPad As String, oe As Object, ordemE As Collection, ke As String
    eqPad = UCase$(TrimU(CStr(TF(C("EQUIPAMENTO_PADRAO")))))
    Set oe = CreateObject("Scripting.Dictionary"): Set ordemE = New Collection
    If UCase$(TrimU(CStr(TF(C("SETOR"))))) = "HEMATOLOGIA" And eqPad <> "" Then
        For i = 1 To nF
            If F(i, Q_PART) = SIM And UCase$(TrimU(CStr(TF(F(i, Q_EQ))))) <> eqPad Then
                ke = KT(F(i, Q_EQ)) & Chr$(0) & KT(F(i, Q_AN))
                If Not oe.Exists(ke) Then oe.Add ke, Array(F(i, Q_EQ), F(i, Q_AN), 0, Empty): ordemE.Add ke
                info = oe(ke)
                info(2) = info(2) + 1
                If Not Vazio(F(i, Q_DH)) Then If IsEmpty(info(3)) Or F(i, Q_DH) > info(3) Then info(3) = F(i, Q_DH)
                oe(ke) = info
            End If
        Next i
        For Each k In ordemE
            info = oe(k)
            Achado "ALERTA", "A10", "Resultado de outro equipamento fora do Painel (Hematologia)", Empty, info(1), Empty, info(3), _
                info(2) & " resultado(s) do equipamento '" & TF(info(0)) & "': o Painel, a Estatistica e o Westgard da Hematologia sao " & _
                "do equipamento padrao (CFG EQUIPAMENTO_PADRAO = " & eqPad & "). Se for o mesmo aparelho com outro nome, corrija o lancamento."
        Next k
    End If

    MarcaQ "resto"
    CalcularQA = Ordenar()
    MarcaQ "ordenar"
End Function

' ---------------------------------------------------------------- apoio
Private Function TabelaQ(ByVal nome As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set TabelaQ = ws.ListObjects(nome)
        On Error GoTo 0
        If Not TabelaQ Is Nothing Then Exit Function
    Next ws
End Function

Private Function Cabecalho(ByVal lo As ListObject) As Object
    Dim h As Variant, j As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If Not lo Is Nothing Then
        h = lo.HeaderRowRange.Value
        For j = 1 To UBound(h, 2): d(CStr(h(1, j))) = j: Next j
    End If
    Set Cabecalho = d
End Function

Private Function ValCab(ByRef m As Variant, ByVal r As Long, ByVal cab As Object, ByVal nome As String) As Variant
    If cab.Exists(nome) Then ValCab = m(r, cab(nome)) Else ValCab = Empty
End Function

Private Function KT(ByVal v As Variant) As String
    If Vazio(v) Then KT = Chr$(1) Else KT = "s" & CStr(TF(v))
End Function

Private Function Chave6(ByRef F As Variant, ByVal i As Long) As String
    Chave6 = KT(F(i, Q_EQ)) & Chr$(0) & KT(F(i, Q_MZ)) & Chr$(0) & KT(F(i, Q_LT)) & Chr$(0) & KT(F(i, Q_NV)) & Chr$(0) & _
             KT(F(i, Q_AN)) & Chr$(0) & KT(F(i, Q_DT))
End Function

' Valores brutos iguais nos campos que o A03 confere (VALOR, FLAG, ANALITO, LOTE, NIVEL, DATA_HORA ao segundo)?
' Igual aqui implica igual na comparacao normalizada; diferente cai na comparacao completa.
' ixR/ixO = colunas de VALOR, FLAG, ANALITO, LOTE, NIVEL, DATA_HORA (0 = ausente).
Private Function BrutoIgual(ByRef rec As Variant, ByVal r As Long, ByRef ixR() As Long, ByRef orig As Variant, ByVal i As Long, _
                            ByRef ixO() As Long) As Boolean
    Dim k As Long, a As Variant, b As Variant
    For k = 1 To 5
        If ixR(k) > 0 Then a = rec(r, ixR(k)) Else a = Empty
        If ixO(k) > 0 Then b = orig(i, ixO(k)) Else b = Empty
        If Vazio(a) Or Vazio(b) Then
            If Not (Vazio(a) And Vazio(b)) Then Exit Function
        ElseIf VarType(a) <> VarType(b) And (VarType(a) = vbString Or VarType(b) = vbString) Then
            Exit Function
        ElseIf a <> b Then
            Exit Function
        End If
    Next k
    If ixR(6) = 0 Or ixO(6) = 0 Then Exit Function
    a = Segundos(rec(r, ixR(6))): b = Segundos(orig(i, ixO(6)))
    BrutoIgual = IgualN(a, b)
End Function

Private Function IgualN(ByVal a As Variant, ByVal b As Variant) As Boolean
    If Vazio(a) Or Vazio(b) Then IgualN = (Vazio(a) And Vazio(b)) Else IgualN = (CDbl(a) = CDbl(b))
End Function

Private Function IgualT(ByVal a As Variant, ByVal b As Variant) As Boolean
    If Vazio(a) Or Vazio(b) Then IgualT = (Vazio(a) And Vazio(b)) Else IgualT = (CStr(a) = CStr(b))
End Function

Private Function Segundos(ByVal v As Variant) As Variant
    If Vazio(Num(v)) Then Segundos = Empty Else Segundos = Round(Num(v) * 86400#, 0)
End Function

Private Sub GrupoE07(ByVal grp As Object, ByVal ordem As Collection, ByVal id As Variant, ByVal an As Variant, ByVal nv As Variant, _
                     ByVal dataHora As Variant)
    Dim info As Variant
    If Not grp.Exists(id) Then
        grp.Add id, Array(1, an, nv, dataHora): ordem.Add id
    Else
        info = grp(id)
        info(0) = info(0) + 1
        If Not Vazio(dataHora) Then If IsEmpty(info(3)) Or dataHora < info(3) Then info(3) = dataHora
        grp(id) = info
    End If
End Sub

Private Sub GrupoAnalito(ByRef F As Variant, ByVal nF As Long, ByVal col As Long, ByVal valor As String, ByVal cod As String, _
                         ByVal teste As String, ByVal sufixo As String)
    Dim g As Object, ordem As Collection, i As Long, k As Variant, ka As String
    Set g = CreateObject("Scripting.Dictionary"): Set ordem = New Collection
    For i = 1 To nF
        If F(i, col) = valor Then
            ka = KT(F(i, Q_AN))
            If Not g.Exists(ka) Then g.Add ka, Array(F(i, Q_AN), 0): ordem.Add ka
            g(ka) = Array(g(ka)(0), g(ka)(1) + 1)
        End If
    Next i
    For Each k In ordem
        Achado "INFO", cod, teste, Empty, g(k)(0), Empty, Empty, g(k)(1) & sufixo
    Next k
End Sub

' partes por (severidade, codigo); dentro de cada uma, DATA_HORA (nulo primeiro) e a ordem de entrada no empate
Private Function Ordenar() As Variant
    Dim n As Long, i As Long, chave() As String, a As Variant, peso As String, res() As Variant, j As Long, ordem() As Long
    Dim s As String
    n = mOut.Count
    If n = 0 Then Ordenar = Empty: Exit Function
    ReDim chave(1 To n)
    For i = 1 To n
        a = mOut(i)
        Select Case a(0)
            Case "ERRO": peso = "1"
            Case "ALERTA": peso = "2"
            Case Else: peso = "3"
        End Select
        If IsEmpty(a(6)) Then
            s = Chr$(1)
        Else
            s = CStr(Int(CDbl(a(6)) * 86400000# + 0.5))
            s = Chr$(2) & String$(15 - Len(s), "0") & s
        End If
        chave(i) = peso & a(1) & Chr$(0) & s & Chr$(0) & Format$(i, "0000000000")
    Next i
    ordem = OrdenarQ(chave, n)
    ReDim res(1 To n, 1 To 8)
    For i = 1 To n
        a = mOut(ordem(i))
        For j = 0 To 7: res(i, j + 1) = a(j): Next j
    Next i
    Ordenar = res
End Function

Private Function OrdenarQ(ByRef chave() As String, ByVal n As Long) As Long()
    Dim a() As Long, b() As Long, w As Long, i As Long, lo As Long, md As Long, hi As Long, l As Long, r As Long, k As Long
    Dim t() As Long
    ReDim a(1 To n): ReDim b(1 To n)
    For i = 1 To n: a(i) = i: Next i
    w = 1
    Do While w < n
        lo = 1
        Do While lo <= n
            md = lo + w - 1: If md > n Then md = n
            hi = lo + 2 * w - 1: If hi > n Then hi = n
            l = lo: r = md + 1: k = lo
            Do While l <= md And r <= hi
                If chave(a(r)) < chave(a(l)) Then b(k) = a(r): r = r + 1 Else b(k) = a(l): l = l + 1
                k = k + 1
            Loop
            Do While l <= md: b(k) = a(l): l = l + 1: k = k + 1: Loop
            Do While r <= hi: b(k) = a(r): r = r + 1: k = k + 1: Loop
            lo = lo + 2 * w
        Loop
        t = a: a = b: b = t
        w = w * 2
    Loop
    OrdenarQ = a
End Function
