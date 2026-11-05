Attribute VB_Name = "mCEQForm"
Option Explicit
' ============================================================================
'  mCEQForm -- registro e consulta do Controle Externo da Qualidade (ADR-087)
'
'  A janela frmCEQ organiza o registro na sequencia natural de quem tem o
'  relatorio do provedor na mao: 1) a rodada (provedor, survey/rodada, ano);
'  2) o analito e a amostra; 3) o resultado do laboratorio e o do provedor
'  (alvo, SD, SDI, limites, n de laboratorios); 4) a avaliacao e a fonte.
'  Grava na MESMA tabela de digitacao do provedor de sempre (tblEQA_CAP_Dados /
'  tblEQA_Controllab_Dados, colunas A..N, Q, R -- as de formula nao sao tocadas)
'  e consolida a EQA_Base pelo mesmo mEQA.AtualizarEQABase. Nenhum campo mudou de
'  significado e nenhuma regra do CEQ mudou: a janela so organiza, valida e evita
'  duplicidade (a chave da EQA_Base: provedor|ano|rodada|analito|amostra).
'  Digitar direto na tabela continua possivel, como antes.
'  ADR-089: corrigir e excluir pela janela (a aba e bloqueada para o ANALISTA), aviso do analito fora do mapa ou do
'  par %/# com confirmacao, consolidacao na hora e trilha (CEQ_ALTERADO, CEQ_EXCLUIDO, EQA_BASE_CONSOLIDADA).
' ============================================================================
Private Const TB_CAP As String = "tblEQA_CAP_Dados"
Private Const TB_CTL As String = "tblEQA_Controllab_Dados"
' colunas digitadas da tabela do provedor (1-based)
Public Const CQ_PROV As Long = 1
Public Const CQ_RODADA As Long = 2
Public Const CQ_ANO As Long = 3
Public Const CQ_ANALITO As Long = 4
Public Const CQ_AMOSTRA As Long = 5
Public Const CQ_RES As Long = 6
Public Const CQ_ALVO As Long = 7
Public Const CQ_SD As Long = 8
Public Const CQ_SDI As Long = 9
Public Const CQ_LINF As Long = 10
Public Const CQ_LSUP As Long = 11
Public Const CQ_AVAL As Long = 12
Public Const CQ_NLABS As Long = 13
Public Const CQ_UNID As Long = 14
Public Const CQ_PAG As Long = 17
Public Const CQ_ARQ As Long = 18
' ADR-089 (OC-0032): o que estava selecionado numa aba de digitacao antes da edicao direta (o "antes" da trilha)
Private mSelAba As String
Private mSelEnd As String
Private mSelVal As Variant


' ADR-090 (OC-0033): registrar CEQ exige a sessao de login (mSeguranca.SessaoAtiva), como inativar. Sem papel minimo
' novo: qualquer usuario logado registra (o papel minimo depende de autorizacao).
Public Function TextoSemSessaoCEQ() As String
    TextoSemSessaoCEQ = "Faça o login para registrar resultados do CEQ."
End Function

Public Sub AbrirCEQ()
    On Error Resume Next
    Dim f As Object
    If Not mSeguranca.SessaoAtiva() Then MsgBox TextoSemSessaoCEQ(), vbExclamation, "Controle externo (CEQ)": Exit Sub
    Set f = VBA.UserForms.Add("frmCEQ")
    If f Is Nothing Then MsgBox "A janela de registro do CEQ nao esta instalada.", vbExclamation: Exit Sub
    f.Show
    Unload f
End Sub

Public Function TabelaCEQ(ByVal provedor As String) As ListObject
    Dim nome As String, ws As Worksheet
    On Error Resume Next
    If UCase$(Trim$(provedor)) = "CAP" Then nome = TB_CAP Else nome = TB_CTL
    For Each ws In ThisWorkbook.Worksheets
        Set TabelaCEQ = ws.ListObjects(nome)
        If Not TabelaCEQ Is Nothing Then Exit Function
    Next ws
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

' Numero digitado (virgula ou ponto decimal). ok = False quando nao e numero; vazio = Empty e ok = True.
Public Function NumeroCEQ(ByVal s As String, ByRef ok As Boolean) As Variant
    Dim t As String
    ok = True
    t = Trim$(s)
    If Len(t) = 0 Then NumeroCEQ = Empty: Exit Function
    t = Replace(t, " ", "")
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then t = Replace(t, ".", "")   ' 1.234,5
    t = Replace(t, ",", Application.International(xlDecimalSeparator))
    t = Replace(t, ".", Application.International(xlDecimalSeparator))
    If IsNumeric(t) Then NumeroCEQ = CDbl(t) Else ok = False
End Function

' Valores distintos (ordem alfabetica; rodadas e anos em ordem decrescente) de uma coluna da tabela do provedor.
' filtroCol/filtroVal: so as linhas com aquele valor (ex.: os analitos de uma rodada). Para as listas da janela.
Public Function DistintosCEQ(ByVal provedor As String, ByVal coluna As Long, Optional ByVal decrescente As Boolean = False, _
                             Optional ByVal filtroCol As Long = 0, Optional ByVal filtroVal As String = "") As Variant
    Dim lo As ListObject, v As Variant, d As Object, i As Long, s As String, a() As String, j As Long, t As String, n As Long
    Dim k As Variant
    DistintosCEQ = Array()
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        s = Txt(v(i, coluna))
        ' ADR-090 (OC-0034): as linhas com Uso = NAO (simulacao: mEQA.UsoAnalitico do Arquivo fonte) nao entram nas listas
        If Len(s) > 0 And Len(Txt(v(i, CQ_ANALITO))) > 0 And mEQA.UsoAnalitico(v(i, CQ_ARQ)) <> mEQA.USO_NAO Then
            If filtroCol = 0 Or StrComp(Txt(v(i, IIf(filtroCol > 0, filtroCol, 1))), filtroVal, vbTextCompare) = 0 Then
                If Not d.Exists(s) Then d.Add s, True
            End If
        End If
    Next i
    n = d.Count
    If n = 0 Then Exit Function
    ReDim a(0 To n - 1)
    i = 0
    For Each k In d.Keys
        a(i) = CStr(k): i = i + 1
    Next k
    For i = 1 To n - 1
        t = a(i): j = i - 1
        Do While j >= 0
            If decrescente Then
                If StrComp(a(j), t, vbTextCompare) >= 0 Then Exit Do
            Else
                If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            End If
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    DistintosCEQ = a
sai:
End Function

' Primeiro valor da coluna "coluna" na linha mais recente em que "chaveCol" = chaveVal (ex.: o ano de uma rodada que
' ja existe, a unidade de um analito). "" se nao houver. Reaproveita o que ja esta registrado, sem inventar.
Public Function ValorRegistrado(ByVal provedor As String, ByVal chaveCol As Long, ByVal chaveVal As String, _
                                ByVal coluna As Long) As String
    Dim lo As ListObject, v As Variant, i As Long
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    For i = UBound(v, 1) To 1 Step -1
        If StrComp(Txt(v(i, chaveCol)), chaveVal, vbTextCompare) = 0 And Len(Txt(v(i, coluna))) > 0 Then
            ValorRegistrado = Txt(v(i, coluna))
            Exit Function
        End If
    Next i
sai:
End Function

' Registros da rodada para a lista da janela: matriz (n x 5) Analito | Amostra | Resultado | Avaliacao | linha da aba
' (ADR-089: a linha, oculta na lista, e o que o Corrigir/Excluir selecionado usa); Empty = nenhum.
Public Function RegistrosDaRodada(ByVal provedor As String, ByVal rodada As String, Optional ByVal ano As String = "") As Variant
    Dim lo As ListObject, v As Variant, i As Long, n As Long, out() As Variant, idx() As Long
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Or Len(rodada) = 0 Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    ReDim idx(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If StrComp(Txt(v(i, CQ_RODADA)), rodada, vbTextCompare) = 0 And Len(Txt(v(i, CQ_ANALITO))) > 0 Then
            If Len(ano) = 0 Or Txt(v(i, CQ_ANO)) = ano Then n = n + 1: idx(n) = i
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim out(1 To n, 1 To 5)
    For i = 1 To n
        out(i, 1) = Txt(v(idx(i), CQ_ANALITO))
        out(i, 2) = Txt(v(idx(i), CQ_AMOSTRA))
        out(i, 3) = Txt(v(idx(i), CQ_RES))
        out(i, 4) = Txt(v(idx(i), CQ_AVAL))
        out(i, 5) = lo.DataBodyRange.Row + idx(i) - 1
    Next i
    RegistrosDaRodada = out
sai:
End Function

' ADR-089: a linha da aba e um registro da tabela do provedor (dentro do corpo, com analito)?
Private Function LinhaDaTabela(ByVal lo As ListObject, ByVal linha As Long) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    If linha < lo.DataBodyRange.Row Or linha > lo.DataBodyRange.Row + lo.ListRows.Count - 1 Then Exit Function
    LinhaDaTabela = (Len(Txt(lo.Parent.Cells(linha, lo.Range.Column + CQ_ANALITO - 1).Value)) > 0)
End Function

' Texto de uma celula como o usuario digitaria (numero com a virgula da maquina; o NumeroCEQ aceita os dois).
Private Function TxtCel(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    TxtCel = Trim$(CStr(v))
End Function

' ADR-089: os 16 campos de um registro gravado (a ordem da janela: Provedor, Rodada, Ano, Analito, Amostra, Resultado,
' Alvo, SD, SDI, LimInf, LimSup, Avaliacao, NLabs, Unidade, Pagina, Arquivo). Empty se a linha nao e registro.
Public Function LinhaCEQ(ByVal provedor As String, ByVal linha As Long) As Variant
    Dim lo As ListObject, ws As Worksheet, c0 As Long, out(0 To 15) As String, k As Long, cols As Variant
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If Not LinhaDaTabela(lo, linha) Then Exit Function
    Set ws = lo.Parent
    c0 = lo.Range.Column - 1
    cols = Array(CQ_PROV, CQ_RODADA, CQ_ANO, CQ_ANALITO, CQ_AMOSTRA, CQ_RES, CQ_ALVO, CQ_SD, CQ_SDI, CQ_LINF, CQ_LSUP, _
                 CQ_AVAL, CQ_NLABS, CQ_UNID, CQ_PAG, CQ_ARQ)
    For k = 0 To 15
        out(k) = TxtCel(ws.Cells(linha, c0 + cols(k)).Value)
    Next k
    LinhaCEQ = out
sai:
End Function

Private Function NomesCampos() As Variant
    NomesCampos = Array("", "Provedor", "Rodada / Survey", "Ano", "Analito", "Amostra", "Seu resultado", "Valor-alvo", "SD", _
                        "SDI", "Limite inferior", "Limite superior", "Avaliação", "Nº de laboratórios", "Unidade", "Página", _
                        "Arquivo fonte")
End Function

' "CAP C-A 2025 · Glucose, serum · CHM-01 (resultado 110)" -- para a confirmacao da exclusao e a trilha.
Public Function DescricaoCEQ(ByVal provedor As String, ByVal linha As Long) As String
    Dim v As Variant
    v = LinhaCEQ(provedor, linha)
    If Not IsArray(v) Then Exit Function
    DescricaoCEQ = v(0) & " " & RodadaAno(v(1), v(2)) & " · " & v(3) & " · " & v(4) & " (resultado " & v(5) & ")"
End Function

' ADR-090 (OC-0055): "C-A 2025" + ano 2025 = "C-A 2025" (o ano so entra quando a rodada nao o traz)
Public Function RodadaAno(ByVal rodada As String, ByVal ano As String) As String
    rodada = Trim$(rodada): ano = Trim$(ano)
    If Len(ano) = 0 Then
        RodadaAno = rodada
    ElseIf InStr(1, rodada, ano, vbTextCompare) > 0 Then
        RodadaAno = rodada
    Else
        RodadaAno = Trim$(rodada & " " & ano)
    End If
End Function

Private Function ContextoCEQ(ByVal v As Variant) As String
    ContextoCEQ = v(0) & " | " & v(1) & " | " & v(2) & " | " & v(3) & " | " & v(4)
End Function

' Campos que mudaram: "Seu resultado: 110 -> 105; Unidade: mg/dL -> g/dL" ("" = nada mudou).
Private Function DiferencasCEQ(ByVal a As Variant, ByVal b As Variant) As String
    Dim k As Long, nomes As Variant, s As String
    nomes = NomesCampos()
    For k = 0 To 15
        If StrComp(CStr(a(k)), CStr(b(k)), vbBinaryCompare) <> 0 Then
            s = s & IIf(Len(s) > 0, "; ", "") & nomes(k + 1) & ": " & IIf(Len(a(k)) = 0, "(vazio)", a(k)) & " " & _
                ChrW$(8594) & " " & IIf(Len(b(k)) = 0, "(vazio)", b(k))
        End If
    Next k
    DiferencasCEQ = s
End Function

' ---------------------------------------------------------------------------
' ADR-089 (OC-0029/OC-0030): o analito digitado no mapa do provedor (EQA_Base W:Y). So aviso e confirmacao: nenhum
' mapa e criado e a gravacao continua possivel.
' ---------------------------------------------------------------------------
Private Function DicMapa(ByVal provedor As String) As Object
    Dim m As Variant, i As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    m = mEQA.MapaDoProvedor(provedor)
    If IsArray(m) Then
        For i = 1 To UBound(m, 1)
            If Not d.Exists(CStr(m(i, 1))) Then d.Add CStr(m(i, 1)), CStr(m(i, 2))
        Next i
    End If
    Set DicMapa = d
End Function

' Outro nome do mapa que comeca pelo digitado e vai para OUTRO canonico ("Lymphs" -> "Lymphs Absolute" = LYMPH#).
' Devolve o nome e o canonico do primeiro par, ou "" quando nao ha par.
' ADR-090 (OC-0052): primeiro o prefixo de palavra ("Lymphs" + espaco); se nao houver, o prefixo literal ("MCH" de "MCHC").
Private Function ParDoMapa(ByVal mapa As Object, ByVal analito As String, ByVal canon As String, ByRef canonPar As String) As String
    Dim k As Variant, u As String, passada As Long
    For passada = 1 To 2
        u = UCase$(Trim$(analito)) & IIf(passada = 1, " ", "")
        For Each k In mapa.Keys
            If Len(CStr(k)) > Len(u) And Left$(UCase$(CStr(k)), Len(u)) = u And _
               StrComp(CStr(mapa(k)), canon, vbTextCompare) <> 0 Then
                ParDoMapa = CStr(k): canonPar = CStr(mapa(k))
                Exit Function
            End If
        Next k
    Next passada
End Function

' Aviso da janela para o analito digitado: "AZUL|No QC_INI: ..." (entra no bias), "AMBAR|..." (fora do mapa) ou "".
Public Function AvisoAnalitoCEQ(ByVal provedor As String, ByVal analito As String) As String
    Dim mapa As Object, canon As String, par As String, cp As String, s As String
    On Error GoTo sai
    analito = Trim$(analito)
    If Len(analito) = 0 Then Exit Function
    Set mapa = DicMapa(provedor)
    If mapa.Exists(analito) Then canon = mapa(analito)
    If Len(canon) > 0 Then
        s = "No QC_INI: " & canon & ". Este resultado entra no bias deste analito."
        par = ParDoMapa(mapa, analito, canon, cp)
        If Len(par) > 0 Then s = s & " Atenção: existe também " & ChrW$(8220) & par & ChrW$(8221) & " (" & cp & ")."
        AvisoAnalitoCEQ = "AZUL|" & s
    Else
        AvisoAnalitoCEQ = "AMBAR|" & ChrW$(9888) & " " & ChrW$(8220) & analito & ChrW$(8221) & " não está no mapa do " & _
            IIf(UCase$(Trim$(provedor)) = "CAP", "CAP", "Controllab") & ": este resultado NÃO entra no bias, no ET nem no " & _
            "Sigma. Use o nome da lista (como no relatório) ou peça ao administrador para mapear."
    End If
sai:
End Function

' A unidade mais usada nos registros ja gravados do mesmo canonico (fora a linha em correcao e o que e simulacao).
Private Function UnidadeDoCanonico(ByVal provedor As String, ByVal mapa As Object, ByVal canon As String, _
                                   ByVal linhaIgnorar As Long) As String
    Dim lo As ListObject, v As Variant, i As Long, d As Object, an As String, u As String, k As Variant, melhor As Long
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        If lo.DataBodyRange.Row + i - 1 <> linhaIgnorar Then
            an = Txt(v(i, CQ_ANALITO)): u = Txt(v(i, CQ_UNID))
            If Len(an) > 0 And Len(u) > 0 And mEQA.UsoAnalitico(v(i, CQ_ARQ)) = mEQA.USO_SIM Then
                If mapa.Exists(an) Then
                    If StrComp(CStr(mapa(an)), canon, vbTextCompare) = 0 Then d(u) = d(u) + 1
                End If
            End If
        End If
    Next i
    For Each k In d.Keys
        If d(k) > melhor Then melhor = d(k): UnidadeDoCanonico = CStr(k)
    Next k
sai:
End Function

Private Function SemEspacos(ByVal s As String) As String
    SemEspacos = UCase$(Replace(Trim$(s), " ", ""))
End Function

' O que precisa da confirmacao do usuario ANTES de gravar (a janela pergunta; Nao = nada gravado). "" = nada.
' campos: os 16 da janela; linhaAlvo: a linha em correcao (0 = registro novo).
Public Function ConfirmacoesCEQ(ByVal campos As Variant, Optional ByVal linhaAlvo As Long = 0) As String
    Dim prov As String, an As String, un As String, mapa As Object, canon As String, uReg As String, par As String
    Dim cp As String, s As String
    On Error GoTo sai
    prov = Trim$(CStr(campos(0))): an = Trim$(CStr(campos(3))): un = Trim$(CStr(campos(13)))
    If Len(an) = 0 Then Exit Function
    Set mapa = DicMapa(prov)
    If mapa.Exists(an) Then canon = mapa(an)
    If Len(canon) = 0 Then
        s = "O analito " & ChrW$(8220) & an & ChrW$(8221) & " não está no mapa do " & prov & ". O registro será gravado, " & _
            "mas NÃO entra no bias, no ET nem no Sigma até o administrador incluir o nome no mapa (EQA_Base, colunas W:Y)."
    ElseIf Len(un) > 0 Then
        uReg = UnidadeDoCanonico(prov, mapa, canon, linhaAlvo)
        If Len(uReg) > 0 And SemEspacos(uReg) <> SemEspacos(un) Then
            s = "A unidade " & ChrW$(8220) & un & ChrW$(8221) & " é diferente da já registrada para " & canon & " (" & _
                ChrW$(8220) & uReg & ChrW$(8221) & ")."
            par = ParDoMapa(mapa, an, canon, cp)
            If Len(par) > 0 Then s = s & " Se o resultado é do " & cp & ", use o nome " & ChrW$(8220) & par & ChrW$(8221) & "."
        End If
    End If
    ' ADR-090 (OC-0031): sem o valor-alvo a amostra e gravada, mas nao entra no bias (so na verificacao dos limites)
    If Len(Trim$(CStr(campos(6)))) = 0 Then
        s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & TextoSemAlvoCEQ()
    End If
    ConfirmacoesCEQ = s
sai:
End Function

Public Function TextoSemAlvoCEQ() As String
    TextoSemAlvoCEQ = "Sem o valor-alvo, esta amostra NÃO entra no bias (nem no ET e no Sigma). Ela conta só na " & _
                      "verificação dos limites."
End Function

' Aviso devolvido no OK (terceiro campo): fora do bias ou unidade diferente.
Private Function AvisoGravado(ByVal prov As String, ByVal an As String, ByVal un As String, ByVal linha As Long) As String
    Dim mapa As Object, canon As String, uReg As String
    On Error GoTo sai
    Set mapa = DicMapa(prov)
    If mapa.Exists(an) Then canon = mapa(an)
    If Len(canon) = 0 Then
        AvisoGravado = "Gravado fora do bias (analito não mapeado)."
    ElseIf Len(un) > 0 Then
        uReg = UnidadeDoCanonico(prov, mapa, canon, linha)
        If Len(uReg) > 0 And SemEspacos(uReg) <> SemEspacos(un) Then
            AvisoGravado = "No bias do " & canon & ", com unidade diferente da já registrada (" & ChrW$(8220) & uReg & _
                           ChrW$(8221) & ")."
        End If
    End If
sai:
End Function

' Grava um resultado de CEQ (nucleo sem dialogo -- a janela e o QA chamam este). campos = Array de 16 textos na ordem
' Provedor, Rodada, Ano, Analito, Amostra, Resultado, Alvo, SD, SDI, LimInf, LimSup, Avaliacao, NLabs, Unidade,
' Pagina, Arquivo. Valida (obrigatorios, numeros, limites, duplicidade pela chave da EQA_Base) ANTES de gravar; grava
' na 1a linha vazia da tabela do provedor (ou acrescenta uma); registra CEQ_REGISTRADO no Audit_Log; consolida a
' EQA_Base. Devolve "OK|<linha da aba>|<aviso>" ou "ERRO|<campo>|<mensagem>" (nada gravado no ERRO).
' ADR-089: o aviso diz quando o analito fica fora do bias (nao mapeado) ou a unidade diverge do canonico.
Public Function SalvarCEQ(ByVal campos As Variant) As String
    If Not mSeguranca.SessaoAtiva() Then SalvarCEQ = "ERRO|0|Sessão encerrada: faça o login de novo. Nada foi gravado.": Exit Function   ' ADR-090 (OC-0033)
    SalvarCEQ = GravarCEQ(campos, 0)
End Function

' ADR-089 (OC-0027): corrige o registro da linha <linha> da aba do provedor pela janela -- mesma validacao do
' registro novo (a duplicidade ignora a propria linha), grava A..N, Q, R na MESMA linha, registra CEQ_ALTERADO com o
' antes e o depois de cada campo e consolida a EQA_Base. Mesmo retorno do SalvarCEQ.
Public Function CorrigirCEQ(ByVal linha As Long, ByVal campos As Variant) As String
    If linha <= 0 Then CorrigirCEQ = "ERRO|0|Escolha o registro a corrigir.": Exit Function
    CorrigirCEQ = GravarCEQ(campos, linha)
End Function

Private Function GravarCEQ(ByVal campos As Variant, ByVal linhaAlvo As Long) As String
    Dim prov As String, rod As String, ano As String, an As String, am As String, i As Long
    Dim lo As ListObject, ws As Worksheet, v As Variant, r As Long, prot As Boolean, ev As Boolean, sE As String
    Dim nums(1 To 16) As Variant, ok As Boolean, aviso As String, livre As Long, nomes As Variant
    Dim antes As Variant, depois As Variant, dif As String
    On Error GoTo falha
    nomes = NomesCampos()
    prov = Trim$(CStr(campos(0))): rod = Trim$(CStr(campos(1))): ano = Trim$(CStr(campos(2)))
    an = Trim$(CStr(campos(3))): am = Trim$(CStr(campos(4)))
    If UCase$(prov) <> "CAP" And UCase$(prov) <> "CONTROLLAB" Then GravarCEQ = "ERRO|1|Escolha o provedor (CAP ou Controllab).": Exit Function
    If Len(rod) = 0 Then GravarCEQ = "ERRO|2|Informe a rodada / survey.": Exit Function
    ' so os 4 algarismos (IsNumeric aceitava "2,02", "20.6", "-202" e "9e99", que estourava o CLng no meio da gravacao)
    If Not ano Like "####" Or Val(ano) < 1990 Or Val(ano) > 2100 Then GravarCEQ = "ERRO|3|Ano com 4 dígitos (ex.: 2026).": Exit Function
    If Len(an) = 0 Then GravarCEQ = "ERRO|4|Informe o analito (nome como no relatório do provedor).": Exit Function
    If Len(am) = 0 Then GravarCEQ = "ERRO|5|Informe a amostra (ex.: CHM-01).": Exit Function
    Dim cn As Variant
    For Each cn In Array(6, 7, 8, 9, 10, 11, 13, 15)
        nums(cn) = NumeroCEQ(CStr(campos(cn - 1)), ok)
        If Not ok Then GravarCEQ = "ERRO|" & cn & "|" & nomes(cn) & ": digite um número.": Exit Function
    Next cn
    If IsEmpty(nums(6)) Then GravarCEQ = "ERRO|6|Informe o resultado do laboratório.": Exit Function
    If Len(Trim$(CStr(campos(11)))) = 0 Then GravarCEQ = "ERRO|12|Informe a avaliação do provedor.": Exit Function
    If Not IsEmpty(nums(10)) And Not IsEmpty(nums(11)) Then
        If nums(10) > nums(11) Then GravarCEQ = "ERRO|10|O limite inferior está maior que o superior.": Exit Function
    End If
    If Not IsEmpty(nums(8)) Then
        If nums(8) < 0 Then GravarCEQ = "ERRO|8|O SD não pode ser negativo.": Exit Function
    End If
    Set lo = TabelaCEQ(prov)
    If lo Is Nothing Then GravarCEQ = "ERRO|1|A tabela de digitação do " & prov & " não foi encontrada.": Exit Function
    Set ws = lo.Parent
    If linhaAlvo > 0 Then
        antes = LinhaCEQ(prov, linhaAlvo)
        If Not IsArray(antes) Then
            GravarCEQ = "ERRO|0|A linha " & linhaAlvo & " não é um registro da aba " & ws.Name & ". Nada foi alterado."
            Exit Function
        End If
    End If
    ' duplicidade: a chave da EQA_Base (provedor|ano|rodada|analito|amostra); na correcao, a propria linha nao conta
    If lo.ListRows.Count > 0 Then
        v = lo.DataBodyRange.Value
        For r = 1 To UBound(v, 1)
            If Len(Txt(v(r, CQ_ANALITO))) > 0 Then
                If lo.DataBodyRange.Row + r - 1 <> linhaAlvo Then
                    If StrComp(Txt(v(r, CQ_RODADA)), rod, vbTextCompare) = 0 And Txt(v(r, CQ_ANO)) = ano And _
                       StrComp(Txt(v(r, CQ_ANALITO)), an, vbTextCompare) = 0 And _
                       StrComp(Txt(v(r, CQ_AMOSTRA)), am, vbTextCompare) = 0 Then
                        ' ADR-089 (OC-0027): a aba e bloqueada para o ANALISTA -- o caminho e o botao Corrigir da janela
                        GravarCEQ = "ERRO|5|Já existe este registro (" & prov & " " & RodadaAno(rod, ano) & " · " & an & " · " & _
                                    am & ") na linha " & (lo.DataBodyRange.Row + r - 1) & ". Para corrigir, selecione-o em " & _
                                    ChrW$(8220) & "Registros desta rodada" & ChrW$(8221) & " e clique em Corrigir selecionado."
                        Exit Function
                    End If
                End If
            ElseIf livre = 0 Then
                If Len(Txt(v(r, CQ_RODADA)) & Txt(v(r, CQ_AMOSTRA)) & Txt(v(r, CQ_RES))) = 0 Then livre = r
            End If
        Next r
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    If linhaAlvo > 0 Then
        r = linhaAlvo
    Else
        If livre = 0 Then
            lo.ListRows.Add
            livre = lo.ListRows.Count
        End If
        r = lo.DataBodyRange.Row + livre - 1
    End If
    ws.Cells(r, lo.Range.Column + CQ_PROV - 1).Value = IIf(UCase$(prov) = "CAP", "CAP", "Controllab")
    ws.Cells(r, lo.Range.Column + CQ_RODADA - 1).Value = "'" & rod
    ws.Cells(r, lo.Range.Column + CQ_ANO - 1).Value = CLng(ano)
    ws.Cells(r, lo.Range.Column + CQ_ANALITO - 1).Value = "'" & an
    ws.Cells(r, lo.Range.Column + CQ_AMOSTRA - 1).Value = "'" & am
    For Each cn In Array(6, 7, 8, 9, 10, 11, 13)
        If IsEmpty(nums(cn)) Then
            ws.Cells(r, lo.Range.Column + cn - 1).ClearContents
        Else
            ws.Cells(r, lo.Range.Column + cn - 1).Value = nums(cn)
        End If
    Next cn
    ws.Cells(r, lo.Range.Column + CQ_AVAL - 1).Value = "'" & Trim$(CStr(campos(11)))
    ws.Cells(r, lo.Range.Column + CQ_UNID - 1).Value = "'" & Trim$(CStr(campos(13)))
    If IsEmpty(nums(15)) Then
        ws.Cells(r, lo.Range.Column + CQ_PAG - 1).ClearContents
    Else
        ws.Cells(r, lo.Range.Column + CQ_PAG - 1).Value = nums(15)
    End If
    ws.Cells(r, lo.Range.Column + CQ_ARQ - 1).Value = "'" & Trim$(CStr(campos(15)))
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    On Error Resume Next
    If linhaAlvo > 0 Then
        depois = LinhaCEQ(prov, r)
        dif = DiferencasCEQ(antes, depois)
        If Len(dif) = 0 Then
            GravarCEQ = "OK|" & r & "|Nenhum campo mudou: nada a corrigir."
            Exit Function
        End If
        mAuditoria.RegistrarLog "CEQ_ALTERADO", ContextoCEQ(antes) & " | " & dif & " | pela janela | aba " & ws.Name & _
                                " linha " & r
    Else
        mAuditoria.RegistrarLog "CEQ_REGISTRADO", prov & " | " & rod & " | " & ano & " | " & an & " | " & am & " | resultado " & _
                                CStr(nums(6)) & " | avaliação " & Trim$(CStr(campos(11))) & " | aba " & ws.Name & " linha " & r
    End If
    Err.Clear
    aviso = AvisoGravado(prov, an, Trim$(CStr(campos(13))), r)
    If IsEmpty(nums(7)) Then aviso = Trim$(aviso & " Gravado fora do bias (sem valor-alvo).")      ' ADR-090 (OC-0031)
    Err.Clear
    mEQA.ConsolidarEQABaseDe IIf(linhaAlvo > 0, "janela do CEQ (correção)", "janela do CEQ (registro)")
    If Err.Number <> 0 Then aviso = Trim$(aviso & " Gravado, mas a consolidação da EQA_Base falhou (" & Err.Description & "). " & _
                                    "Clique em ATUALIZAR DADOS.")
    On Error GoTo 0
    GravarCEQ = "OK|" & r & "|" & aviso
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    GravarCEQ = "ERRO|0|Não foi possível gravar (detalhe: " & sE & "). Nada foi alterado."
End Function

' ADR-089 (OC-0027): exclui o registro da linha <linha> (a janela confirma antes). Limpa A..N, Q, R -- a linha da
' tabela fica, vazia, e e reaproveitada pelo proximo registro --, registra CEQ_EXCLUIDO com os valores apagados e
' consolida a EQA_Base (o registro sai do bias, do ET e do Sigma). "OK|<linha>|<aviso>" ou "ERRO|0|<mensagem>".
Public Function ExcluirCEQ(ByVal provedor As String, ByVal linha As Long) As String
    Dim lo As ListObject, ws As Worksheet, v As Variant, prot As Boolean, ev As Boolean, cn As Variant, sE As String
    Dim aviso As String, nomes As Variant, k As Long, s As String
    On Error GoTo falha
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then ExcluirCEQ = "ERRO|0|A tabela de digitação do " & provedor & " não foi encontrada.": Exit Function
    Set ws = lo.Parent
    v = LinhaCEQ(provedor, linha)
    If Not IsArray(v) Then
        ExcluirCEQ = "ERRO|0|A linha " & linha & " não é um registro da aba " & ws.Name & ". Nada foi excluído."
        Exit Function
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    For Each cn In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, CQ_PAG, CQ_ARQ)
        ws.Cells(linha, lo.Range.Column + cn - 1).ClearContents
    Next cn
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    On Error Resume Next
    nomes = NomesCampos()
    For k = 5 To 15
        If Len(v(k)) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & nomes(k + 1) & " " & v(k)
    Next k
    mAuditoria.RegistrarLog "CEQ_EXCLUIDO", ContextoCEQ(v) & " | valores apagados: " & s & " | pela janela | aba " & _
                            ws.Name & " linha " & linha
    Err.Clear
    mEQA.ConsolidarEQABaseDe "janela do CEQ (exclusão)"
    If Err.Number <> 0 Then aviso = "Excluído, mas a consolidação da EQA_Base falhou (" & Err.Description & "). " & _
                                    "Clique em ATUALIZAR DADOS."
    On Error GoTo 0
    ExcluirCEQ = "OK|" & linha & "|" & aviso
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    ExcluirCEQ = "ERRO|0|Não foi possível excluir (detalhe: " & sE & "). Nada foi alterado."
End Function

' ---------------------------------------------------------------------------
' ADR-089 (OC-0028/OC-0032): edicao DIRETA nas abas EQA.CAP_Dados / EQA.Controllab_Dados (quem tem a aba liberada).
' O modulo da aba chama GuardarSelecaoCEQ (SelectionChange: o "antes") e EdicaoDiretaCEQ (Change): marca a edicao
' pendente (Z4 -- o Z3 avisa e o ATUALIZAR consolida) e registra CEQ_ALTERADO com o antes e o depois de cada celula.
' ---------------------------------------------------------------------------
Public Sub GuardarSelecaoCEQ(ByVal alvo As Range)
    On Error GoTo sai
    mSelAba = "": mSelEnd = "": mSelVal = Empty
    If alvo.Areas.Count > 1 Then Exit Sub
    If alvo.Cells.CountLarge > 2000 Then Exit Sub
    mSelVal = alvo.Value
    mSelAba = alvo.Worksheet.Name
    mSelEnd = alvo.Address
sai:
End Sub

Private Function ValorAntes(ByVal c As Range) As String
    Dim rg As Range, i As Long, j As Long
    ValorAntes = "?"
    On Error GoTo sai
    If mSelAba <> c.Worksheet.Name Or Len(mSelEnd) = 0 Then Exit Function
    Set rg = c.Worksheet.Range(mSelEnd)
    If Intersect(rg, c) Is Nothing Then Exit Function
    i = c.Row - rg.Row + 1: j = c.Column - rg.Column + 1
    If IsArray(mSelVal) Then ValorAntes = TxtCel(mSelVal(i, j)) Else ValorAntes = TxtCel(mSelVal)
sai:
End Function

' ADR-090 (OC-0054): o valor do campo na EQA_Base para o registro da linha (chave provedor|ano|rodada|analito|amostra,
' a mesma da consolidacao). Sem o registro na base (registro novo ou campo da chave editado) ou sem coluna na base
' (No de laboratorios): devolve "padrao" (o "?" de antes).
Private Function ValorConsolidado(ByVal prov As String, ByVal lo As ListObject, ByVal c As Range, ByVal col As Long, _
                                  ByVal padrao As String) As String
    Dim b As Worksheet, ch As String, m As Variant, bc As Long, ws As Worksheet
    ValorConsolidado = padrao
    On Error GoTo sai
    Select Case col
        Case CQ_PROV: bc = mEQA.B_PROVEDOR
        Case CQ_RODADA: bc = mEQA.B_RODADA
        Case CQ_ANO: bc = mEQA.B_ANO
        Case CQ_ANALITO: bc = mEQA.B_ANALITO
        Case CQ_AMOSTRA: bc = mEQA.B_AMOSTRA
        Case CQ_RES: bc = mEQA.B_RESULTADO
        Case CQ_ALVO: bc = mEQA.B_ALVO
        Case CQ_SD: bc = mEQA.B_SD
        Case CQ_SDI: bc = mEQA.B_SDI
        Case CQ_LINF: bc = mEQA.B_LIMINF
        Case CQ_LSUP: bc = mEQA.B_LIMSUP
        Case CQ_AVAL: bc = mEQA.B_AVAL_ORIG
        Case CQ_UNID: bc = mEQA.B_UNIDADE
        Case CQ_PAG: bc = mEQA.B_PAGINA
        Case CQ_ARQ: bc = mEQA.B_ARQUIVO
        Case Else: Exit Function
    End Select
    Set ws = c.Worksheet
    ch = prov & "|" & Txt(ws.Cells(c.Row, lo.Range.Column + CQ_ANO - 1).Value) & "|" & _
         Txt(ws.Cells(c.Row, lo.Range.Column + CQ_RODADA - 1).Value) & "|" & _
         Txt(ws.Cells(c.Row, lo.Range.Column + CQ_ANALITO - 1).Value) & "|" & _
         Txt(ws.Cells(c.Row, lo.Range.Column + CQ_AMOSTRA - 1).Value)
    Set b = ThisWorkbook.Worksheets(mEQA.EQA_BASE)
    m = Application.Match(ch, b.Range(b.Cells(mEQA.BASE_R0, mEQA.B_CHAVE), b.Cells(mEQA.BASE_RN, mEQA.B_CHAVE)), 0)
    If IsError(m) Then Exit Function
    ValorConsolidado = TxtCel(b.Cells(mEQA.BASE_R0 + CLng(m) - 1, bc).Value)
sai:
End Function

Public Sub EdicaoDiretaCEQ(ByVal alvo As Range)
    Dim ws As Worksheet, lo As ListObject, zona As Range, c As Range, prov As String, n As Long, det As String
    Dim col As Long, a As String, d As String, cab As String, ctx As Variant, resto As Long, nF As Long, detF As String
    On Error GoTo sai
    Set ws = alvo.Worksheet
    If ws.Name = mEQA.ABA_CAP Then prov = "CAP" Else prov = "Controllab"
    Set lo = TabelaCEQ(prov)
    If lo Is Nothing Then GoTo sai
    If lo.Parent.Name <> ws.Name Or lo.ListRows.Count = 0 Then GoTo sai
    Set zona = Intersect(alvo, lo.DataBodyRange)
    If zona Is Nothing Then GoTo sai
    For Each c In zona.Cells
        col = c.Column - lo.Range.Column + 1
        ' ADR-090 (OC-0053): coluna de FORMULA (O Bias, P |Bias| e as de apoio S..X) sobrescrita -- a consolidacao copia o
        ' valor digitado para a EQA_Base e o bias do ET/Sigma o usaria. O Z3 avisa (FORMULA SOBRESCRITA) e fica a trilha.
        If col > CQ_UNID And col <> CQ_PAG And col <> CQ_ARQ And Not c.HasFormula Then
            nF = nF + 1
            If nF <= 30 Then detF = detF & IIf(Len(detF) > 0, ", ", "") & c.Address(False, False) & " (" & _
                TxtCel(lo.HeaderRowRange.Cells(1, col).Value) & ") = " & IIf(Len(TxtCel(c.Value)) = 0, "(vazio)", TxtCel(c.Value))
        End If
        If (col >= CQ_PROV And col <= CQ_UNID) Or col = CQ_PAG Or col = CQ_ARQ Then
            a = ValorAntes(c): d = TxtCel(c.Value)
            ' ADR-090 (OC-0054): sem o "antes" da selecao, vale o que esta consolidado na EQA_Base para o registro
            If a = "?" Then a = ValorConsolidado(prov, lo, c, col, a)
            If a <> d Then
                n = n + 1
                If n <= 30 Then
                    cab = TxtCel(lo.HeaderRowRange.Cells(1, col).Value)
                    ctx = LinhaCEQ(prov, c.Row)
                    det = det & IIf(Len(det) > 0, " || ", "") & "linha " & c.Row & _
                          IIf(IsArray(ctx), " (" & ContextoCEQ(ctx) & ")", "") & " · " & cab & ": " & _
                          IIf(Len(a) = 0, "(vazio)", a) & " " & ChrW$(8594) & " " & IIf(Len(d) = 0, "(vazio)", d)
                Else
                    resto = resto + 1
                End If
            End If
        End If
    Next c
    If n > 0 Then
        mEQA.MarcarEdicaoCEQ ws
        If resto > 0 Then det = det & " || e mais " & resto & " célula(s)"
        mAuditoria.RegistrarLog "CEQ_ALTERADO", "edição direta na tabela | aba " & ws.Name & " | " & det & _
                                " | falta consolidar (Consolidar na EQA_Base ou ATUALIZAR DADOS)"
    End If
    If nF > 0 Then
        mAuditoria.RegistrarLog "CEQ_FORMULA_SOBRESCRITA", "edição direta numa coluna de fórmula | aba " & ws.Name & " | " & _
                                detF & IIf(nF > 30, " e mais " & (nF - 30) & " célula(s)", "") & _
                                " | o bias usaria o valor digitado: restaure a fórmula (copie de outra linha da coluna)"
    End If
sai:
    GuardarSelecaoCEQ alvo               ' o "antes" da proxima edicao e o valor de agora
End Sub

' ADR-089 (OC-0032): troca do provedor do CEQ que vale para TODOS os analitos (Hematologia, Estatistica!R4).
' Na Bioquimica o provedor e por analito (Painel M4 -> mEstatistica.GravarProvedorAnalito, que registra o mesmo evento).
Public Sub ProvedorEQAlterado(ByVal antes As String, ByVal depois As String, ByVal onde As String)
    On Error Resume Next
    antes = Trim$(antes): depois = Trim$(depois)
    If StrComp(antes, depois, vbTextCompare) = 0 And Len(antes) > 0 Then Exit Sub
    mAuditoria.RegistrarLog "PROVEDOR_ALTERADO", onde & " | provedor do CEQ: " & IIf(Len(antes) = 0, "(não registrado)", antes) & _
                            " " & ChrW$(8594) & " " & IIf(Len(depois) = 0, "(vazio)", depois)
End Sub
