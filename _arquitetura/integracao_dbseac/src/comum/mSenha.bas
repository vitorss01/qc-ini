Attribute VB_Name = "mSenha"
Option Explicit
' ============================================================================
'  mSenha -- hash de senha PBKDF2-HMAC-SHA256 com salt aleatorio (ADR-084)
'
'  Mesmo modulo no QC de Bioquimica, no de Hematologia e no cadastro central
'  (QC_INI_cadastrar usuarios.xlsm). Nada de senha em texto: so o derivado.
'
'  Formato gravado (uma celula, autodescritivo -- algoritmo, parametros, salt e
'  resultado; o mesmo do Django/passlib):
'      pbkdf2_sha256$<iteracoes>$<salt 16 bytes em hex>$<chave 32 bytes em hex>
'
'  Criptografia do Windows (CNG, bcrypt.dll), nao VBA: BCryptDeriveKeyPBKDF2 com
'  HMAC-SHA256 e BCryptGenRandom (gerador criptografico do sistema; nunca Rnd).
'  Senha em UTF-8 (WideCharToMultiByte): acentos geram o mesmo hash em qualquer
'  maquina e no Python (hashlib.pbkdf2_hmac('sha256', senha.encode('utf-8'), ...)).
'  Declaracoes PtrSafe/LongPtr (Office 64 bits, o do laboratorio); no Office 32
'  bits o contador de iteracoes (ULONGLONG) vai como Currency (#If Win64).
'
'  ITERACOES_PADRAO = 310.000 (OWASP 2021; ~0,8 s por login nesta maquina). 600.000 (OWASP 2023) mediu 1,6 s --
'  lento para assinar cada corrida. Alterar aqui quando as maquinas permitirem.
'  O numero fica gravado em cada hash: aumentar a constante vale para as senhas
'  novas; as antigas continuam conferindo com o numero delas.
' ============================================================================
Public Const ITERACOES_PADRAO As Long = 310000
Public Const PREFIXO As String = "pbkdf2_sha256"
Private Const BYTES_SALT As Long = 16
Private Const BYTES_CHAVE As Long = 32
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2
Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, _
    ByVal pbBuffer As LongPtr, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" (ByVal hPrf As LongPtr, _
    ByVal pbPassword As LongPtr, ByVal cbPassword As Long, ByVal pbSalt As LongPtr, ByVal cbSalt As Long, _
    ByVal cIterations As LongLong, ByVal pbDerivedKey As LongPtr, ByVal cbDerivedKey As Long, _
    ByVal dwFlags As Long) As Long
#Else
Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" (ByVal hPrf As LongPtr, _
    ByVal pbPassword As LongPtr, ByVal cbPassword As Long, ByVal pbSalt As LongPtr, ByVal cbSalt As Long, _
    ByVal cIterations As Currency, ByVal pbDerivedKey As LongPtr, ByVal cbDerivedKey As Long, _
    ByVal dwFlags As Long) As Long
#End If
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

' Hash novo para gravar (salt novo a cada chamada). Senha vazia = erro (quem chama recusa antes).
Public Function GerarHashSenha(ByVal senha As String, Optional ByVal iteracoes As Long = 0) As String
    Dim salt() As Byte
    If Len(senha) = 0 Then Err.Raise vbObjectError + 841, "mSenha", "senha vazia"
    If iteracoes <= 0 Then iteracoes = ITERACOES_PADRAO
    salt = BytesAleatorios(BYTES_SALT)
    GerarHashSenha = PREFIXO & "$" & iteracoes & "$" & Hex2(salt) & "$" & Hex2(Derivar(senha, salt, iteracoes))
End Function

' A senha confere com o hash gravado (so o formato pbkdf2_sha256$...; o SHA-256 antigo fica com quem chama).
' Comparacao em tempo constante. Hash mal formado = False (nunca erro).
Public Function VerificarSenha(ByVal senha As String, ByVal gravado As String) As Boolean
    Dim p As Variant, it As Long, salt() As Byte, esperado() As Byte, obtido() As Byte
    On Error GoTo nao
    If Len(senha) = 0 Then Exit Function
    p = Split(Trim$(gravado), "$")
    If UBound(p) <> 3 Then Exit Function
    If LCase$(p(0)) <> PREFIXO Then Exit Function
    it = CLng(p(1))
    If it < 1000 Then Exit Function                    ' parametro absurdo: nao aceita
    salt = DeHex(CStr(p(2)))
    esperado = DeHex(CStr(p(3)))
    obtido = Derivar(senha, salt, it, UBound(esperado) + 1)
    VerificarSenha = IguaisTempoConstante(obtido, esperado)
nao:
End Function

Public Function EhHashPbkdf2(ByVal gravado As String) As Boolean
    EhHashPbkdf2 = (LCase$(Left$(Trim$(gravado), Len(PREFIXO) + 1)) = PREFIXO & "$")
End Function

' Para o QA: o derivado com salt e iteracoes dados (comparado com o hashlib do Python).
Public Function DerivarHex(ByVal senha As String, ByVal saltHex As String, ByVal iteracoes As Long) As String
    DerivarHex = Hex2(Derivar(senha, DeHex(saltHex), iteracoes))
End Function

' ---------------------------------------------------------------- CNG
Private Function Derivar(ByVal senha As String, ByRef salt() As Byte, ByVal iteracoes As Long, _
                         Optional ByVal tamanho As Long = BYTES_CHAVE) As Byte()
    Dim hAlg As LongPtr, st As Long, pw() As Byte, npw As Long, dk() As Byte, alg As String
    alg = "SHA256"
    st = BCryptOpenAlgorithmProvider(hAlg, StrPtr(alg), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If st <> 0 Then Err.Raise vbObjectError + 842, "mSenha", "CNG indisponivel (BCryptOpenAlgorithmProvider " & Hex$(st) & ")"
    pw = Utf8(senha)
    npw = UBound(pw) + 1
    ReDim dk(0 To tamanho - 1)
#If Win64 Then
    st = BCryptDeriveKeyPBKDF2(hAlg, VarPtr(pw(0)), npw, VarPtr(salt(0)), UBound(salt) + 1, CLngLng(iteracoes), _
                               VarPtr(dk(0)), tamanho, 0)
#Else
    st = BCryptDeriveKeyPBKDF2(hAlg, VarPtr(pw(0)), npw, VarPtr(salt(0)), UBound(salt) + 1, CCur(iteracoes) / 10000@, _
                               VarPtr(dk(0)), tamanho, 0)
#End If
    Erase pw                                          ' a senha em bytes nao fica na memoria do modulo
    BCryptCloseAlgorithmProvider hAlg, 0
    If st <> 0 Then Err.Raise vbObjectError + 843, "mSenha", "falha no PBKDF2 (" & Hex$(st) & ")"
    Derivar = dk
End Function

Private Function BytesAleatorios(ByVal n As Long) As Byte()
    Dim b() As Byte, st As Long
    ReDim b(0 To n - 1)
    st = BCryptGenRandom(0, VarPtr(b(0)), n, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    If st <> 0 Then Err.Raise vbObjectError + 844, "mSenha", "gerador aleatorio do sistema indisponivel (" & Hex$(st) & ")"
    BytesAleatorios = b
End Function

Private Function Utf8(ByVal s As String) As Byte()
    Dim n As Long, b() As Byte
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If n <= 0 Then Err.Raise vbObjectError + 845, "mSenha", "senha nao convertida para UTF-8"
    ReDim b(0 To n - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    Utf8 = b
End Function

' ---------------------------------------------------------------- hex e comparacao
Private Function Hex2(ByRef b() As Byte) As String
    Dim i As Long, s As String
    For i = LBound(b) To UBound(b)
        s = s & Right$("0" & LCase$(Hex$(b(i))), 2)
    Next i
    Hex2 = s
End Function

Private Function DeHex(ByVal h As String) As Byte()
    Dim i As Long, b() As Byte
    h = Trim$(h)
    If Len(h) = 0 Or Len(h) Mod 2 <> 0 Then Err.Raise vbObjectError + 846, "mSenha", "hex invalido"
    ReDim b(0 To Len(h) \ 2 - 1)
    For i = 0 To UBound(b)
        b(i) = CByte("&H" & Mid$(h, 2 * i + 1, 2))
    Next i
    DeHex = b
End Function

Private Function IguaisTempoConstante(ByRef a() As Byte, ByRef b() As Byte) As Boolean
    Dim i As Long, dif As Long
    If UBound(a) <> UBound(b) Then Exit Function
    For i = 0 To UBound(a)
        dif = dif Or (a(i) Xor b(i))
    Next i
    IguaisTempoConstante = (dif = 0)
End Function
