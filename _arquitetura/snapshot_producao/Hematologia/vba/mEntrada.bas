Attribute VB_Name = "mEntrada"
Option Explicit
' ===== UTILITARIOS DE ENTRADA (Fase 2) =====

' Aceita "12,5" e "12.5"; devolve ok=False se nao for numero.
Public Function ParseNum(ByVal s As String, ByRef ok As Boolean) As Double
    Dim t As String
    ok = False
    t = Trim$(s)
    If t = "" Then Exit Function
    t = Replace(t, " ", "")
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        t = Replace(t, ".", ""): t = Replace(t, ",", ".")
    ElseIf InStr(t, ",") > 0 Then
        t = Replace(t, ",", ".")
    End If
    If Not IsNumeric(t) Then Exit Function
    ParseNum = Val(t)
    ok = True
End Function

Public Function ParseData(ByVal s As String, ByRef ok As Boolean) As Date
    Dim t As String
    ok = False
    t = Trim$(s)
    If t = "" Then Exit Function
    On Error Resume Next
    If IsDate(t) Then
        ParseData = CDate(t)
        ok = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

' Codigo completo do lote a partir do nucleo de 6 digitos + nivel.
Public Function CodigoLote(ByVal loteCore As String, ByVal nivel As Long) As String
    CodigoLote = "QC-" & loteCore & Format(nivel, "00")
End Function

Public Function NucleoLote(ByVal codigo As String) As String
    NucleoLote = Mid$(Trim$(codigo), 4, Len(Trim$(codigo)) - 5)
End Function

