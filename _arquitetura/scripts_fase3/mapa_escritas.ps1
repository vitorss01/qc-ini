# mapa_escritas.ps1 - quem escreve em que aba
#
# Insumo do parecer arquitetural: varre os modulos VBA e mapeia, por
# procedimento, quais abas sao referenciadas e quais operacoes de ESCRITA
# ocorrem. Serve para provar (ou derrubar) a afirmacao "so o motor escreve em
# Eng_Saida e ninguem escreve nas abas de interface".
#
# Heuristica, nao compilador: resolve variavel -> aba pelo Set mais recente
# dentro do mesmo procedimento. Basta para achar duplicidade de
# responsabilidade e escrita fora do lugar.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\mapa_escritas.ps1 -Pasta <dir com .bas/.cls/.frm> [-OutCsv <mapa.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Pasta,
    [string]$OutCsv
)

$enc = [System.Text.Encoding]::Default
$escrita = '(\.Value2?\s*=|\.Formula(R1C1)?\s*=|\.ClearContents|\.Delete\b|' +
           '\.Insert\b|\.Copy\b|\.PasteSpecial|\.Cut\b|\.Add\b)'

$linhas = New-Object System.Collections.ArrayList

# --- constantes de nome de aba, declaradas em qualquer modulo ---
# mDados faz Sheets(BANCO) com Public Const BANCO = "DB_Resultados". Sem
# resolver a constante, a varredura perde a camada de dados inteira e o mapa
# mente por omissao.
$constAba = @{}
foreach ($arq in Get-ChildItem $Pasta -Include *.bas, *.cls, *.frm -Recurse) {
    foreach ($t in [System.IO.File]::ReadAllLines($arq.FullName, $enc)) {
        if ($t -match '(?i)Const\s+(\w+)\s+As\s+String\s*=\s*"([^"]+)"') {
            $constAba[$Matches[1]] = $Matches[2]
        }
    }
}

foreach ($arq in Get-ChildItem $Pasta -Include *.bas, *.cls, *.frm -Recurse) {
    $L = [System.IO.File]::ReadAllLines($arq.FullName, $enc)
    $proc = '(nivel de modulo)'
    $vars = @{}
    for ($i = 0; $i -lt $L.Count; $i++) {
        $t = $L[$i]
        if ($t -match '^\s*''') { continue }

        if ($t -match '^\s*(Public |Private |Friend )?(Sub|Function|Property\s+\w+)\s+(\w+)') {
            $proc = $Matches[3]
            $vars = @{}
            continue
        }

        # Set x = ThisWorkbook.Sheets("Aba")  /  Worksheets("Aba")
        if ($t -match 'Set\s+(\w+)\s*=.*(Sheets|Worksheets)\s*\(\s*"([^"]+)"') {
            $vars[$Matches[1]] = $Matches[3]
            continue
        }
        # Set x = ThisWorkbook.Sheets(CONSTANTE)
        if ($t -match 'Set\s+(\w+)\s*=.*(Sheets|Worksheets)\s*\(\s*(\w+)\s*\)') {
            $nome = $Matches[3]
            if ($constAba.ContainsKey($nome)) { $vars[$Matches[1]] = $constAba[$nome] }
            continue
        }

        if ($t -notmatch $escrita) { continue }

        # a quem pertence a escrita?
        $alvo = $null
        if ($t -match '(Sheets|Worksheets)\s*\(\s*"([^"]+)"') { $alvo = $Matches[2] }
        else {
            foreach ($v in $vars.Keys) {
                if ($t -match "\b$v\s*\.") { $alvo = $vars[$v]; break }
            }
        }
        if (-not $alvo) { continue }

        [void]$linhas.Add([pscustomobject]@{
                Modulo       = $arq.BaseName
                Procedimento = $proc
                Aba          = $alvo
                Linha        = $i + 1
                Codigo       = $t.Trim()
            })
    }
}

"operacoes de escrita mapeadas: $($linhas.Count)"
""
"=== por aba de destino ==="
$linhas | Group-Object Aba | Sort-Object Count -Descending | ForEach-Object {
    $mods = ($_.Group | Select-Object -ExpandProperty Modulo -Unique) -join ', '
    "{0,-20} n={1,-5} <- {2}" -f $_.Name, $_.Count, $mods
}
""
"=== abas de interface (nao deveriam receber escrita de VBA) ==="
$interface = @('Calc', 'Painel', 'Estatística', 'Estatistica')
$mal = @($linhas | Where-Object { $interface -contains $_.Aba })
if ($mal.Count -eq 0) { "  nenhuma - OK" }
else {
    $mal | ForEach-Object { "  {0}!{1} <- {2}.{3} linha {4}" -f $_.Aba, '', $_.Modulo, $_.Procedimento, $_.Linha }
    $mal | Group-Object Modulo, Procedimento | ForEach-Object { "  {0,-40} n={1}" -f $_.Name, $_.Count }
}

if ($OutCsv) {
    $linhas | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    ""
    "detalhe em: $OutCsv"
}
