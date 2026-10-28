Attribute VB_Name = "mUsuariosCentral"
Option Explicit
' ============================================================================
'  mUsuariosCentral -- copia local do cadastro central de profissionais (ADR-084)
'
'  Fonte unica: QC_INI_cadastrar usuarios.xlsm, tabela tb_registro_usuarios
'  (A FUNCIONARIOS, B LOGIN, C CARGO, D SENHA_HASH, E ASSINATURA, F CONSELHO,
'  G SITUACAO). O QC le essa tabela por ACE numa COPIA temporaria -- o mesmo
'  caminho do DB_SEAC no motor de dados VBA (mEtlReceb, ADR-079); o Power Query
'  nao roda com MOTOR_DADOS = VBA -- e grava na aba tecnica Usuarios_Central
'  (oculta para todos, protegida). Login e assinatura consultam primeiro esta
'  copia; o login que nao esta no central continua na aba Usuarios local (os
'  usuarios que ja existem nao perdem o acesso).
'
'  Caminho: Cfg_Integracao CAMINHO_USUARIOS (o administrador informa, como o
'  CAMINHO_DB_SEAC). Vazio = a mesma pasta do QC, quando ela e uma pasta do
'  Windows (no OneDrive com endereco https a chave tem de ser preenchida).
'
'  Falha de leitura (arquivo ausente, rede fora, layout errado): a copia
'  anterior FICA como estava -- ninguem perde o acesso nem e apagado em
'  silencio; a falha vai para o Audit_Log (uma vez por sessao). Login repetido
'  no central (sem diferenciar maiusculas) nao entra: e ambiguo, e registrado.
'  O hash ja vem pronto do central (mSenha, PBKDF2); a senha nunca passa por aqui.
' ============================================================================
Public Const ABA As String = "Usuarios_Central"
Public Const ARQUIVO_PADRAO As String = "QC_INI_cadastrar usuários.xlsm"
Private Const ABA_CENTRAL As String = "REGISTRO_DE_USUÁRIOS"
Private Const R0 As Long = 4
Private Const NCOL As Long = 9
' colunas da aba Usuarios_Central
Public Const UC_LOGIN As Long = 1
Public Const UC_NOME As Long = 2
Public Const UC_CARGO As Long = 3
Public Const UC_PAPEL As Long = 4
Public Const UC_HASH As Long = 5
Public Const UC_ASSIN As Long = 6
Public Const UC_CONSELHO As Long = 7
Public Const UC_SITUACAO As Long = 8
Public Const UC_SINC As Long = 9
Private Const MIN_ENTRE_SINC As Double = 5# / 1440#      ' 5 min: login seguido nao rele o arquivo
Private mUltimaSinc As Double
Private mAvisouFalha As Boolean

' Papel do QC (ADM / ANALISTA / TECNICO) pelo CARGO do cadastro central; "" = cargo sem acesso ao QC.
Public Function PapelDoCargo(ByVal cargo As String) As String
    Dim c As String
    c = UCase$(Trim$(cargo))
    If c = "ADM" Or c = "ADMIN" Or c = "ADMINISTRADOR" Or Left$(c, 6) = "ADMIN " Then
        PapelDoCargo = "ADM"
    ElseIf Left$(c, 8) = "ANALISTA" Then
        PapelDoCargo = "ANALISTA"
    ElseIf Left$(c, 7) = "TÉCNICO" Or Left$(c, 7) = "TECNICO" Then
        PapelDoCargo = "TÉCNICO"
    End If
End Function

' Caminho do cadastro central ("" = nao ha como saber: chave vazia e o QC aberto de um endereco https).
Public Function Caminho() As String
    Dim c As String
    On Error Resume Next
    c = Trim$(mIntegracao.Cfg("CAMINHO_USUARIOS"))
    On Error GoTo 0
    If Len(c) = 0 Then
        If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
        c = ThisWorkbook.Path & "\" & ARQUIVO_PADRAO
    End If
    Caminho = c
End Function

' Le o central e substitui a copia. forcar = False: no maximo uma leitura a cada 5 min por sessao.
' Devolve "OK|<n usuarios>|<ignorados>", "PULADO|recente" ou "ERRO|<motivo>" (a copia anterior fica).
Public Function Sincronizar(Optional ByVal forcar As Boolean = False) As String
    Dim cam As String, msg As String, v As Variant, n As Long, i As Long, out() As Variant, k As Long
    Dim d As Object, lg As String, rep As Object, ign As String, ws As Worksheet, prot As Boolean, ev As Boolean
    On Error GoTo falha
    If Not forcar And mUltimaSinc > 0 And Now - mUltimaSinc < MIN_ENTRE_SINC Then Sincronizar = "PULADO|recente": Exit Function
    cam = Caminho()
    If Len(cam) = 0 Then
        Sincronizar = Falhou("Cfg_Integracao CAMINHO_USUARIOS vazio e o QC aberto de um endereco da web: informe o " & _
                             "caminho do " & ARQUIVO_PADRAO)
        Exit Function
    End If
    msg = LerCentral(cam, v, n)
    If Len(msg) > 0 Then Sincronizar = Falhou(msg): Exit Function
    ' logins repetidos no central: nenhum dos dois entra (ambiguo)
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Set rep = CreateObject("Scripting.Dictionary"): rep.CompareMode = 1
    For i = 1 To n
        lg = Trim$(CStr(v(i, 2)))
        If Len(lg) > 0 Then
            If d.Exists(lg) Then rep(lg) = 1 Else d.Add lg, i
        End If
    Next i
    If n > 0 Then ReDim out(1 To n, 1 To NCOL)
    For i = 1 To n
        lg = Trim$(CStr(v(i, 2)))
        If Len(lg) = 0 Then
        ElseIf rep.Exists(lg) Then
            If InStr(1, ign, "|" & lg & "|", vbTextCompare) = 0 Then ign = ign & "|" & lg & "|"
        Else
            k = k + 1
            out(k, UC_LOGIN) = lg
            out(k, UC_NOME) = Trim$(CStr(v(i, 1)))
            out(k, UC_CARGO) = Trim$(CStr(v(i, 3)))
            out(k, UC_PAPEL) = PapelDoCargo(CStr(v(i, 3)))
            out(k, UC_HASH) = Trim$(CStr(v(i, 4)))
            out(k, UC_ASSIN) = Trim$(CStr(v(i, 5)))
            out(k, UC_CONSELHO) = Trim$(CStr(v(i, 6)))
            out(k, UC_SITUACAO) = IIf(Len(Trim$(CStr(v(i, 7)))) = 0, "ATIVO", UCase$(Trim$(CStr(v(i, 7)))))
            out(k, UC_SINC) = Now
        End If
    Next i
    Set ws = ThisWorkbook.Worksheets(ABA)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    ws.Range(ws.Cells(R0, 1), ws.Cells(ws.Rows.Count, NCOL)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(R0, 1), ws.Cells(R0 + n - 1, NCOL)).Value = out    ' linhas k+1..n ficam vazias
        ws.Range(ws.Cells(R0, UC_HASH), ws.Cells(R0 + k - 1, UC_HASH)).NumberFormat = "@"
    End If
    ws.Range("B1").Value = Now
    ws.Range("D1").Value = cam
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mUltimaSinc = Now
    mAvisouFalha = False
    If Len(ign) > 0 Then
        Auditar "USUARIOS_CENTRAL_LOGIN_REPETIDO", "login(s) repetido(s) no cadastro central, sem acesso ate corrigir: " & _
                Replace(Mid$(ign, 2, Len(ign) - 2), "||", ", ")
    End If
    Sincronizar = "OK|" & k & "|" & Replace(Replace(ign, "||", ","), "|", "")
    Exit Function
falha:
    msg = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = True
    Sincronizar = Falhou("erro ao gravar a copia: " & msg)
End Function

' Dados de um login na copia do central. False = nao esta no central (vale a aba Usuarios local).
Public Function Buscar(ByRef login As String, ByRef nome As String, ByRef cargo As String, ByRef papel As String, _
                       ByRef hash As String, ByRef assinatura As String, ByRef conselho As String, _
                       ByRef situacao As String) As Boolean
    Dim ws As Worksheet, ult As Long, v As Variant, i As Long
    On Error GoTo sai
    If Len(Trim$(login)) = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets(ABA)
    ult = ws.Cells(ws.Rows.Count, UC_LOGIN).End(xlUp).Row
    If ult < R0 Then Exit Function
    v = ws.Range(ws.Cells(R0, 1), ws.Cells(ult + 1, NCOL)).Value
    For i = 1 To UBound(v, 1) - 1
        If UCase$(Trim$(CStr(v(i, UC_LOGIN)))) = UCase$(Trim$(login)) Then
            login = CStr(v(i, UC_LOGIN))
            nome = CStr(v(i, UC_NOME)): cargo = CStr(v(i, UC_CARGO)): papel = CStr(v(i, UC_PAPEL))
            hash = CStr(v(i, UC_HASH)): assinatura = CStr(v(i, UC_ASSIN)): conselho = CStr(v(i, UC_CONSELHO))
            situacao = CStr(v(i, UC_SITUACAO))
            Buscar = True
            Exit Function
        End If
    Next i
sai:
End Function

' Arquivo de imagem da assinatura ("" = sem imagem): o valor da coluna E pode ser o caminho completo ou so o nome
' do arquivo -- nesse caso, a pasta "assinaturas" ao lado do cadastro central.
Public Function ArquivoAssinatura(ByVal ref As String) As String
    Dim c As String
    ref = Trim$(ref)
    If Len(ref) = 0 Then Exit Function
    If InStr(ref, ":") > 0 Or Left$(ref, 2) = "\\" Then
        c = ref
    Else
        c = Caminho()
        If Len(c) = 0 Then Exit Function
        c = Left$(c, InStrRev(c, "\")) & "assinaturas\" & ref
    End If
    On Error Resume Next
    If Len(Dir$(c)) > 0 Then ArquivoAssinatura = c
End Function

' ---------------------------------------------------------------- leitura
Private Function LerCentral(ByVal cam As String, ByRef v As Variant, ByRef n As Long) As String
    Dim fso As Object, copia As String, cn As Object, rs As Object, x As Variant, i As Long, j As Long, cab As String
    On Error GoTo falha
    n = 0
    If Len(Dir$(cam)) = 0 Then LerCentral = "cadastro central nao encontrado: " & cam: Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    copia = Environ$("TEMP") & "\qcini_usuarios_leitura.xlsm"
    fso.CopyFile cam, copia, True                      ' le uma copia: o administrador pode estar com o arquivo aberto
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & copia & _
            ";Extended Properties=""Excel 12.0 Macro;HDR=NO;IMEX=1"";Mode=Read;"
    Set rs = cn.Execute("SELECT * FROM [" & ABA_CENTRAL & "$A2:G2000]")
    If rs.EOF Then LerCentral = "aba " & ABA_CENTRAL & " vazia no cadastro central": GoTo fecha
    x = rs.GetRows                                     ' (coluna, linha) 0-based; linha 0 = cabecalho
    For j = 0 To 1
        cab = cab & "|" & UCase$(Trim$(CStr(Nz(x(j, 0)))))
    Next j
    If InStr(cab, "|FUNCION") = 0 Or InStr(cab, "|LOGIN") = 0 Then
        LerCentral = "cadastro central com layout inesperado (A2:B2 = " & Mid$(cab, 2) & ")"
        GoTo fecha
    End If
    n = UBound(x, 2)
    If n > 0 Then
        ReDim v(1 To n, 1 To 7)
        For i = 1 To n
            For j = 0 To 6
                If j <= UBound(x, 1) Then v(i, j + 1) = Nz(x(j, i))
            Next j
        Next i
    Else
        ReDim v(1 To 1, 1 To 7)
    End If
fecha:
    On Error Resume Next
    rs.Close
    cn.Close
    Kill copia
    Exit Function
falha:
    LerCentral = "nao foi possivel ler o cadastro central (" & Err.Number & "): " & Err.Description
    Resume fecha
End Function

Private Function Nz(ByVal x As Variant) As Variant
    If IsNull(x) Or IsError(x) Then Nz = "" Else Nz = x
End Function

Private Function Falhou(ByVal msg As String) As String
    Falhou = "ERRO|" & msg
    If Not mAvisouFalha Then
        Auditar "USUARIOS_CENTRAL_INDISPONIVEL", msg & " -- vale a ultima copia (" & UltimaCopia() & ")"
        mAvisouFalha = True
    End If
End Function

Public Function UltimaCopia() As String
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Worksheets(ABA).Range("B1").Value
    If IsDate(v) Then UltimaCopia = Format$(v, "dd/mm/yyyy hh:mm") Else UltimaCopia = "nenhuma"
End Function

Private Sub Auditar(ByVal acao As String, ByVal detalhe As String)
    On Error Resume Next
    mAuditoria.Auditar mAuditoria.CAT_SEG, acao, "mUsuariosCentral", 0, Empty, "", "", 0, "", Empty, Empty, _
                       "", "", "", detalhe
End Sub
