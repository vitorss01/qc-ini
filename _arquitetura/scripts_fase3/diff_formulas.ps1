# diff_formulas.ps1 - compara dois inventarios gerados por snapshot_formulas.ps1
#
# Criterio de aceite do Sprint HARDENING 1: 100% identico. Qualquer divergencia e
# LISTADA celula a celula, nunca resumida em contagem.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\diff_formulas.ps1 -Referencia baseline_producao.csv -Candidato build_nova.csv
#   .\diff_formulas.ps1 -Referencia a.csv -Candidato b.csv -OutCsv divergencias.csv

param(
    [Parameter(Mandatory = $true)][string]$Referencia,
    [Parameter(Mandatory = $true)][string]$Candidato,
    [string]$OutCsv
)

function Read-Snapshot([string]$path) {
    $h = @{}
    $primeira = $true
    foreach ($linha in [System.IO.File]::ReadLines($path, [System.Text.Encoding]::UTF8)) {
        if ($primeira) { $primeira = $false; continue }
        $p = $linha.Split(';', 4)
        if ($p.Count -lt 4) { continue }
        $h["$($p[0])!$($p[1])"] = @{ Formula = $p[2]; Valor = $p[3] }
    }
    return $h
}

$ref = Read-Snapshot $Referencia
$cnd = Read-Snapshot $Candidato

"Referencia : $Referencia  ($($ref.Count) formulas)"
"Candidato  : $Candidato  ($($cnd.Count) formulas)"
""

$divs = New-Object System.Collections.ArrayList

foreach ($k in $ref.Keys) {
    if (-not $cnd.ContainsKey($k)) {
        [void]$divs.Add([pscustomobject]@{ Tipo = 'AUSENTE'; Chave = $k; Referencia = $ref[$k].Formula; Candidato = '' })
    }
    elseif ($cnd[$k].Formula -ne $ref[$k].Formula) {
        [void]$divs.Add([pscustomobject]@{ Tipo = 'ALTERADA'; Chave = $k; Referencia = $ref[$k].Formula; Candidato = $cnd[$k].Formula })
    }
    elseif ($cnd[$k].Valor -ne $ref[$k].Valor) {
        [void]$divs.Add([pscustomobject]@{ Tipo = 'VALOR'; Chave = $k; Referencia = $ref[$k].Valor; Candidato = $cnd[$k].Valor })
    }
}
foreach ($k in $cnd.Keys) {
    if (-not $ref.ContainsKey($k)) {
        [void]$divs.Add([pscustomobject]@{ Tipo = 'EXTRA'; Chave = $k; Referencia = ''; Candidato = $cnd[$k].Formula })
    }
}

foreach ($t in @('AUSENTE', 'ALTERADA', 'VALOR', 'EXTRA')) {
    "{0,-9} {1}" -f $t, @($divs | Where-Object Tipo -eq $t).Count
}
""

if ($divs.Count -eq 0) {
    "ACEITE: OK - 100% identico."
    exit 0
}

if ($OutCsv) {
    $divs | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    "Divergencias completas em: $OutCsv"
}
else {
    $divs | Format-Table -AutoSize -Wrap
}
"ACEITE: FALHOU - $($divs.Count) divergencias."
exit 1
