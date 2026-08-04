# teste_estresse.ps1 - item 4.5 do Quality Gate
#
# O QUE ESTA SENDO TESTADO, e por que assim.
#
# O gate pede "5.000 RUNs". Para a Bioquimica isso seriam 200.000 linhas
# (20 analitos x 2 niveis por corrida), e as formulas auxiliares do banco
# cobrem 15.000. O teto real do sistema e a LINHA, nao a corrida -- entao a
# carga e medida em linhas, ate o teto declarado.
#
# O risco especifico: as colunas BB e BC do DB_Resultados sao 30.000 COUNTIFS
# de INTERVALO EXPANSIVO (R4C5:RC5). Custo n^2. A 1.000 linhas ninguem sente;
# a pergunta e onde deixa de ser usavel. Medir um ponto so nao responde --
# e preciso a CURVA.
#
# A carga e gerada por VBA, nao por COM: atribuir array 2D grande a um Range
# pelo PowerShell falha em silencio neste ambiente (ver criar_corridas.ps1).
# O VBA escreve bloco de forma confiavel e e ordens de grandeza mais rapido.
#
# NAO TOCA NA PRODUCAO: exige caminho de build, como rodar_motor.ps1.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\teste_estresse.ps1 -Workbook <build.xlsm> [-Marcos 1000,5000,10000,15000]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    # [string[]], nao [int[]]: com powershell.exe -File, parametro de ARRAY
    # chega como UMA string com virgulas, e a conversao para int[] falha antes
    # do corpo do script rodar. Mesma armadilha ja documentada em
    # rodar_motor.ps1 para -Rotinas. Normalizar aqui cobre as duas formas de
    # chamada sem obrigar quem chama a saber da diferenca.
    [string[]]$Marcos = @('1000', '5000', '10000', '15000'),
    [string]$OutCsv
)

$ErrorActionPreference = 'Stop'

$Marcos = @($Marcos | ForEach-Object { $_ -split ',' } |
    Where-Object { $_.Trim() -ne '' } | ForEach-Object { [int]$_.Trim() })

if ($Workbook -notmatch 'build_hardening') {
    throw "Recusado: so roda em copia de build. Recebido: $Workbook"
}

$VBA_CARGA = @'
Option Explicit

' Preenche DB_Resultados com carga sintetica ATE totalLinhas.
' Valores oscilam em torno do alvo de cada analito para que o Westgard tenha
' o que avaliar -- carga constante nao exercita as regras.
Public Function GerarCarga(ByVal totalLinhas As Long) As String
    ' Tratador OBRIGATORIO. Sem ele, erro aqui abre caixa modal -- e com o Excel
    ' em Visible=False a caixa fica invisivel, o processo trava e o COM reporta
    ' "falha na chamada de procedimento remoto", que parece Excel morto e nao e.
    On Error GoTo falha

    Dim ws As Worksheet, wa As Worksheet
    Dim analitos() As String, nAna As Long, i As Long, r As Long
    Dim buf() As Variant, n As Long, run As Long, nivel As Long
    Dim base As Double, dt As Date
    Dim estavaProtegida As Boolean
    Dim etapa As String
    etapa = "inicio"

    Set ws = ThisWorkbook.Sheets("DB_Resultados")
    Set wa = ThisWorkbook.Sheets("Analitos")

    ' O banco sai do build PROTEGIDO (blindagem do entregavel). Destravar e
    ' retravar aqui e o que permite a carga sem afrouxar o artefato.
    estavaProtegida = ws.ProtectContents
    etapa = "unprotect"
    If estavaProtegida Then ws.Unprotect Password:="qcini2025"

    etapa = "ler analitos"
    ReDim analitos(1 To 40)
    nAna = 0
    For i = 4 To 43
        If Trim$(CStr(wa.Cells(i, 1).Value)) <> "" Then
            nAna = nAna + 1
            analitos(nAna) = Trim$(CStr(wa.Cells(i, 1).Value))
        End If
    Next i
    If nAna = 0 Then GerarCarga = "ERRO|sem analitos": Exit Function

    etapa = "limpar banco"
    ws.Range(ws.Cells(4, 1), ws.Cells(15003, 7)).ClearContents

    etapa = "dimensionar buffer " & totalLinhas
    ReDim buf(1 To totalLinhas, 1 To 7)
    n = 0: run = 0: dt = DateSerial(2020, 1, 1)

    etapa = "gerar linhas"
    Do While n < totalLinhas
        run = run + 1
        dt = dt + 1
        For nivel = 1 To NLV
            For i = 1 To nAna
                If n >= totalLinhas Then Exit Do
                n = n + 1
                base = 100 * nivel
                buf(n, 1) = run
                buf(n, 2) = CDbl(dt)
                buf(n, 3) = nivel
                buf(n, 4) = "QC-522611" & Format$(nivel, "00")
                buf(n, 5) = analitos(i)
                ' oscilacao deterministica: +-3% em torno do alvo
                buf(n, 6) = base * (1 + 0.03 * Sin(n / 7#))
                buf(n, 7) = "Ativo"
            Next i
        Next nivel
    Loop

    etapa = "escrever bloco n=" & n
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 7)).Value = buf
    etapa = "formatar datas"
    ws.Range(ws.Cells(4, 2), ws.Cells(3 + n, 2)).NumberFormat = "dd/mm/yyyy"

    If estavaProtegida Then
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If

    GerarCarga = "OK|" & n & "|" & run
    Exit Function
falha:
    ' Capturar ANTES de qualquer outra coisa: On Error GoTo 0 -- e o proprio
    ' Resume Next -- ZERAM o objeto Err. Ler depois devolve "ERRO 0:", que
    ' esconde exatamente a informacao pela qual o tratador existe.
    Dim nErr As Long, dErr As String
    nErr = Err.Number
    dErr = Err.Description

    On Error Resume Next
    If estavaProtegida Then
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    GerarCarga = "ERRO " & nErr & " na etapa [" & etapa & "]: " & dErr
End Function
'@

function Novo-Excel {
    $ultimo = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep -Seconds 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

$tmp = [System.IO.Path]::GetTempFileName() + '.bas'
Set-Content $tmp $VBA_CARGA -Encoding Default

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Arquivo em somente leitura: $Workbook" }

# remove versao anterior do modulo de carga, se houver
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq 'mEstresse') { $wb.VBProject.VBComponents.Remove($c); break }
}
$wb.VBProject.VBComponents.Import($tmp) | Out-Null
Remove-Item $tmp -Force

$res = New-Object System.Collections.ArrayList

foreach ($alvo in $Marcos) {
    $xl.Calculation = -4135          # manual durante a carga
    $ret = [string]$xl.Run("$($wb.Name)!GerarCarga", $alvo)
    if ($ret -notlike 'OK|*') { throw "GerarCarga falhou: $ret" }
    $p = $ret -split '\|'
    $linhas = [int]$p[1]; $corridas = [int]$p[2]

    $xl.Calculation = -4105          # automatico
    $t1 = [Diagnostics.Stopwatch]::StartNew()
    $xl.CalculateFull()
    $t1.Stop()

    # O motor PODE recusar-se a terminar, e isso e resultado valido do teste:
    # a partir de certo volume o buffer de 5.000 eventos de Westgard estoura e
    # a guarda instalada em 04/08/2026 interrompe com erro em vez de descartar
    # violacao em silencio. Registrar como medicao, nao como falha do teste --
    # e justamente o limite que o estresse existe para encontrar.
    $motorOk = $true
    $motorMsg = ''
    $t2 = [Diagnostics.Stopwatch]::StartNew()
    try { $xl.Run("$($wb.Name)!AtualizarEstatistica") | Out-Null }
    catch {
        $motorOk = $false
        $motorMsg = ($_.Exception.Message -replace '\s+', ' ')
        if ($motorMsg.Length -gt 90) { $motorMsg = $motorMsg.Substring(0, 90) }
    }
    $t2.Stop()

    $eventos = [long]$wb.Worksheets('Eventos_Westgard').Range('J2').Value2

    [void]$res.Add([pscustomobject]@{
            Linhas       = $linhas
            Corridas     = $corridas
            RecalcSeg    = [math]::Round($t1.Elapsed.TotalSeconds, 2)
            MotorSeg     = [math]::Round($t2.Elapsed.TotalSeconds, 2)
            TotalSeg     = [math]::Round($t1.Elapsed.TotalSeconds + $t2.Elapsed.TotalSeconds, 2)
            EventosWestg = $eventos
            MotorOk      = $motorOk
            MotorMsg     = $motorMsg
        })
    "{0,6} linhas / {1,4} corridas  ->  recalc {2,6:N2}s  motor {3,6:N2}s  total {4,6:N2}s  eventos {5,6}  {6}" -f `
        $linhas, $corridas, $t1.Elapsed.TotalSeconds, $t2.Elapsed.TotalSeconds, `
    ($t1.Elapsed.TotalSeconds + $t2.Elapsed.TotalSeconds), $eventos, `
    $(if ($motorOk) { "motor OK" } else { "MOTOR INTERROMPIDO: $motorMsg" })
}

""
"=== curva de custo (o que interessa e a INCLINACAO, nao o ponto)"
$ant = $null
foreach ($r in $res) {
    if ($ant) {
        $fatorL = $r.Linhas / $ant.Linhas
        $fatorT = if ($ant.TotalSeg -gt 0) { $r.TotalSeg / $ant.TotalSeg } else { 0 }
        # linear -> fatorT ~ fatorL ; quadratico -> fatorT ~ fatorL^2
        $exp = if ($fatorL -gt 1 -and $fatorT -gt 0) { [math]::Round([math]::Log($fatorT) / [math]::Log($fatorL), 2) } else { 0 }
        "  {0,6} -> {1,6} linhas : tempo x{2,5:N2}  (linhas x{3:N2})  expoente ~{4}" -f `
            $ant.Linhas, $r.Linhas, $fatorT, $fatorL, $exp
    }
    $ant = $r
}

# o modulo de carga NAO fica no artefato
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq 'mEstresse') { $wb.VBProject.VBComponents.Remove($c); break }
}

# A pasta e FECHADA SEM SALVAR: a carga sintetica jamais pode virar o artefato.
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

if ($OutCsv) {
    $res | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    ""
    "medicoes em: $OutCsv"
}
""
"Artefato fechado SEM salvar: a carga sintetica nao ficou no arquivo."
