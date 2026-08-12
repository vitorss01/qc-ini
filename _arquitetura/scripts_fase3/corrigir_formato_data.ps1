# corrigir_formato_data.ps1 - a data do banco volta a mostrar o ano
#
# O DEFEITO
#
# DB_Resultados!B exibia "05/01/yyyy": dia e mes certos, e a palavra "yyyy" no
# lugar do ano. O VALOR estava correto (serial 46027 = 05/01/2026); quem estava
# errado era o FORMATO.
#
# A CAUSA, que e uma armadilha de localidade e nao um erro de digitacao
#
# O codigo do ano no formato numerico do Excel depende do IDIOMA DE FORMATOS da
# instalacao. Em ingles e "yyyy"; em portugues (1046, o desta maquina) e "aaaa".
# O formato gravado era "dd/mm/yyyy", entao o Excel pt-BR leu "dd/mm/" como
# formato e "yyyy" como TEXTO LITERAL -- exatamente como pediram.
#
# Verificado: NumberFormat e NumberFormatLocal devolviam a MESMA string
# ("dd/mm/yyyy"), o que ja denuncia que nao ha traducao acontecendo nesta
# instalacao; trocar por "dd/mm/aaaa" faz a celula exibir 05/01/2026.
#
# POR QUE NAO FIXAR "aaaa" E PRONTO
#
# Isso trocaria uma dependencia de localidade por outra: num Excel em ingles,
# "aaaa" viraria literal do mesmo jeito. O script TENTA e CONFERE -- aplica um
# candidato, le o .Text de volta e so aceita quando o ano aparece. Quem decide e
# o resultado, nao a suposicao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\corrigir_formato_data.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    $db = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'DB_Resultados') { $db = $w; break } }
    if ($db -eq $null) { throw 'aba DB_Resultados ausente' }
    if ($db.ProtectContents) { try { $db.Unprotect($SENHA) } catch { } }

    $ult = $db.Cells($db.Rows.Count, 1).End(-4162).Row
    if ($ult -lt 4) { throw 'banco vazio: nao ha data para conferir' }

    "idioma de formatos do Excel: $($xl.LanguageSettings.LanguageID(3))"
    "antes: B4.Text = '$($db.Range('B4').Text)'  (Value2 = $($db.Range('B4').Value2))"

    # Candidatos, do mais neutro ao localizado. O primeiro que RENDERIZAR o ano
    # vence. Nao ha suposicao sobre qual e o idioma da maquina.
    $candidatos = @(
        @{ Prop = 'NumberFormat'; Valor = 'dd/mm/yyyy' },
        @{ Prop = 'NumberFormatLocal'; Valor = 'dd/mm/aaaa' },
        @{ Prop = 'NumberFormatLocal'; Valor = 'dd/mm/yyyy' }
    )

    $teste = $db.Range('B4')
    $anoEsperado = [DateTime]::FromOADate([double]$teste.Value2).Year.ToString()
    $vencedor = $null
    foreach ($c in $candidatos) {
        try {
            if ($c.Prop -eq 'NumberFormat') { $teste.NumberFormat = $c.Valor } else { $teste.NumberFormatLocal = $c.Valor }
            $t = "$($teste.Text)"
            "  tentativa $($c.Prop)='$($c.Valor)' -> '$t'"
            if ($t -like "*$anoEsperado*") { $vencedor = $c; break }
        }
        catch { "  tentativa $($c.Prop)='$($c.Valor)' -> erro: $($_.Exception.Message.Split([char]10)[0])" }
    }
    if ($vencedor -eq $null) { throw "nenhum formato de data renderizou o ano $anoEsperado -- nada foi alterado" }

    # Aplica na coluna de datas ate a CAPACIDADE declarada, nao ate um teto fixo.
    #
    # Era Max($ult, 15003) -- o 15003 era o provisionamento antigo, equivalente a
    # 13,5 meses. Com o ADR-025 a capacidade e CAP_LINHAS (120.000) e quem a
    # declara e o mBanco; formatar so ate 15003 deixaria toda linha alem disso
    # sem formato de data, exibindo serial cru para o usuario. Achado A5 da
    # auditoria de 12/08/2026.
    #
    # Sem mBanco no arquivo (versao anterior ao ADR-025), formata ate a ultima
    # linha real -- que e o que existe para formatar.
    $limite = $ult
    try {
        $cap = [int]$xl.Run('UltimaLinhaCapacidade')
        if ($cap -gt $limite) { $limite = $cap }
        "capacidade declarada por mBanco: linha $cap"
    }
    catch {
        "mBanco ausente -- formatando ate a ultima linha real ($ult)"
    }
    $faixa = $db.Range($db.Cells(4, 2), $db.Cells($limite, 2))
    if ($vencedor.Prop -eq 'NumberFormat') { $faixa.NumberFormat = $vencedor.Valor } else { $faixa.NumberFormatLocal = $vencedor.Valor }
    "aplicado: $($vencedor.Prop) = '$($vencedor.Valor)' em B4:B$limite"

    # ---- confere, nao confia ----
    $erros = @()
    foreach ($r in @(4, [int]([Math]::Floor(($ult + 4) / 2)), $ult)) {
        $cel = $db.Cells($r, 2)
        if ("$($cel.Value2)" -eq '') { continue }
        $ano = [DateTime]::FromOADate([double]$cel.Value2).Year.ToString()
        $txt = "$($cel.Text)"
        if ($txt -notlike "*$ano*") { $erros += "B$r exibe '$txt', sem o ano $ano" }
        if ($txt -match 'yyyy|aaaa') { $erros += "B$r ainda mostra o codigo do formato: '$txt'" }
    }
    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Correcao de formato rejeitada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }
    "depois: B4.Text = '$($db.Range('B4').Text)'   B$ult`.Text = '$($db.Cells($ult,2).Text)'"

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
