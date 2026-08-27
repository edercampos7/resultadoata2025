<#
Regenera o painel de custos (painel-custos-obra20.html) a partir da planilha de
custos da Obra 20. Uso:

    powershell -File scripts\atualizar_painel.ps1
    powershell -File scripts\atualizar_painel.ps1 -Xlsx "C:\caminho\planilha.xlsx"

Sem -Xlsx, o script usa o .xlsx mais recente na pasta do projeto.
#>
param(
    [string]$Xlsx,
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)

if (-not $Xlsx) {
    $Xlsx = (Get-ChildItem -Path $ProjectDir -Filter *.xlsx | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    if (-not $Xlsx) { throw "Nenhum .xlsx encontrado em $ProjectDir. Informe -Xlsx." }
}
Write-Host "Planilha: $Xlsx"

$work = Join-Path $env:TEMP ("painel_build_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null
$extractDir = Join-Path $work "xlsx_extract"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
# abre com FileShare.ReadWrite pois o Excel costuma manter o .xlsx aberto/travado
# enquanto o usuario o edita
$xlsxStream = New-Object System.IO.FileStream($Xlsx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$xlsxArchive = New-Object System.IO.Compression.ZipArchive($xlsxStream, [System.IO.Compression.ZipArchiveMode]::Read)
foreach ($entry in $xlsxArchive.Entries) {
    if ($entry.FullName.EndsWith('/')) { continue }
    $dest = Join-Path $extractDir $entry.FullName
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
}
$xlsxArchive.Dispose()
$xlsxStream.Dispose()

function Load-Xml($path) {
    $text = [System.IO.File]::ReadAllText($path, $utf8)
    $x = New-Object System.Xml.XmlDocument
    $x.LoadXml($text)
    return $x
}

$ssXml = Load-Xml "$extractDir\xl\sharedStrings.xml"
$ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
$sharedStrings = New-Object System.Collections.Generic.List[string]
foreach ($si in $ssXml.SelectNodes('//a:sst/a:si', $ns)) { $sharedStrings.Add($si.InnerText) }

function Get-CellValue($cellNode, $ns) {
    $t = $cellNode.GetAttribute('t')
    $vNode = $cellNode.SelectSingleNode('a:v', $ns)
    $isNode = $cellNode.SelectSingleNode('a:is', $ns)
    if ($t -eq 's') { if ($vNode) { return $sharedStrings[[int]$vNode.InnerText] }; return $null }
    elseif ($t -eq 'inlineStr') { if ($isNode) { return $isNode.InnerText }; return $null }
    else { if ($vNode) { return $vNode.InnerText }; return $null }
}
function Col-Letters($ref) { return ($ref -replace '[0-9]', '') }
$excelBase = Get-Date -Year 1899 -Month 12 -Day 30
function Excel-ToDate($serial) {
    if (-not $serial -or $serial -eq '-') { return $null }
    $n = 0.0
    if (-not [double]::TryParse($serial, [ref]$n)) { return $null }
    if ($n -le 0) { return $null }
    return $excelBase.AddDays($n).ToString('yyyy-MM-dd')
}
function To-Num($v) {
    if (-not $v -or $v -eq '-') { return 0.0 }
    $n = 0.0
    if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return 0.0
}

# sheet1 = OBRA 20 (ledger), sheet2 = DRF (category tree) -- matches the workbook's
# current tab order/rIds as of 2026-04. If the workbook is restructured, check
# xl/workbook.xml <sheets> order and xl/_rels/workbook.xml.rels rId mapping.
$s2 = Load-Xml "$extractDir\xl\worksheets\sheet2.xml"
$rows2 = $s2.SelectNodes('//a:sheetData/a:row', $ns)
$monthCols = @('F','G','H','I','J','K','L','M','N','O','P','Q')
$monthKeys = @('2026-01','2026-02','2026-03','2026-04','2026-05','2026-06','2026-07','2026-08','2026-09','2026-10','2026-11','2026-12')

$categories = New-Object System.Collections.Generic.List[object]
foreach ($row in $rows2) {
    $rn = [int]$row.r
    if ($rn -le 4 -or $rn -ge 109) { continue }
    $cells = @{}
    foreach ($c in $row.SelectNodes('a:c', $ns)) { $cells[(Col-Letters $c.r)] = Get-CellValue $c $ns }
    $id = $cells['B']; $desc = $cells['C']
    if ((-not $id) -and (-not $desc)) { continue }
    $months2026 = [ordered]@{}
    for ($i = 0; $i -lt $monthCols.Count; $i++) { $months2026[$monthKeys[$i]] = To-Num $cells[$monthCols[$i]] }
    # "total" = coluna E (TOTAL geral da planilha) -- e' a mesma referencia que
    # aparece na propria planilha (DRF), preferida sobre somar o razao (sheet1)
    # pois a formula da DRF tem uma pequena divergencia historica de ancoragem
    # de intervalo (poucas dezenas/centenas de reais) em relacao ao razao bruto.
    $categories.Add([pscustomobject]@{ id = $id; desc = $desc; total = To-Num $cells['E']; months2026 = $months2026 })
}

$idToDesc = @{}
foreach ($cat in $categories) { if ($cat.desc -and $cat.desc.Trim() -ne '') { $idToDesc[$cat.id] = $cat.desc.Trim() } }

function TotalFor($id) { $c = $categories | Where-Object { $_.id -eq $id } | Select-Object -First 1; return $(if ($c) { $c.total } else { 0 }) }

# ids "3." (LUCRO BRUTO) e "4." (MARGEM CANTEIRO) so aparecem no bloco-resumo
# (linhas >=109) da DRF, nao no bloco principal acima
$summaryTotals = @{}
foreach ($row in $rows2) {
    $rn = [int]$row.r
    if ($rn -lt 109) { continue }
    $cells = @{}
    foreach ($c in $row.SelectNodes('a:c', $ns)) { $cells[(Col-Letters $c.r)] = Get-CellValue $c $ns }
    $id = $cells['B']
    if ($id -and -not $summaryTotals.ContainsKey($id)) { $summaryTotals[$id] = To-Num $cells['E'] }
}

function DescFor($id) { return ($categories | Where-Object { $_.id -eq $id } | Select-Object -First 1).desc }
$macroDefs = @(
    @{ id = '2.2.1'; name = (DescFor '2.2.1'); color = '#2a78d6' },
    @{ id = '2.2.2'; name = (DescFor '2.2.2'); color = '#eb6834' },
    @{ id = '2.2.3'; name = (DescFor '2.2.3'); color = '#1baf7a' },
    @{ id = '2.2.4'; name = (DescFor '2.2.4'); color = '#eda100' },
    @{ id = '2.2.5'; name = (DescFor '2.2.5'); color = '#e87ba4' }
)

# sheet1 = OBRA 20 ledger, header row 6, data from row 7
$s1 = Load-Xml "$extractDir\xl\worksheets\sheet1.xml"
$rows1 = $s1.SelectNodes('//a:sheetData/a:row', $ns)
$mesAnoToKey = @{}
$rawTx = New-Object System.Collections.Generic.List[object]
foreach ($row in $rows1) {
    $rn = [int]$row.r
    if ($rn -le 6) { continue }
    $c = @{}
    foreach ($cell in $row.SelectNodes('a:c', $ns)) { $c[(Col-Letters $cell.r)] = Get-CellValue $cell $ns }
    $fornecedor = $c['G']; $historico = $c['H']
    $valorPago = To-Num $c['M']; $valorOrig = To-Num $c['I']
    if ([string]::IsNullOrWhiteSpace($fornecedor) -and [string]::IsNullOrWhiteSpace($historico) -and $valorPago -eq 0 -and $valorOrig -eq 0) { continue }
    $mesAno = $c['C']
    $rawTx.Add([pscustomobject]@{
        mesAno = $mesAno
        idCode = $c['D']; ref = $c['E']
        fornecedor = ($fornecedor -replace '\s+', ' ').Trim()
        historico = $historico
        pagamento = Excel-ToDate $c['L']
        valorPago = $valorPago
    })
}

# derive month keys (YYYY-MM) present in the ledger, in chronological order
$monthLabelsPt = @{ 'jan'='Jan'; 'fev'='Fev'; 'mar'='Mar'; 'abr'='Abr'; 'mai'='Mai'; 'jun'='Jun'; 'jul'='Jul'; 'ago'='Ago'; 'set'='Set'; 'out'='Out'; 'nov'='Nov'; 'dez'='Dez' }
$monthNumPt = @{ 'jan'='01'; 'fev'='02'; 'mar'='03'; 'abr'='04'; 'mai'='05'; 'jun'='06'; 'jul'='07'; 'ago'='08'; 'set'='09'; 'out'='10'; 'nov'='11'; 'dez'='12' }
$seenMonths = New-Object System.Collections.Generic.List[string]
foreach ($t in $rawTx) {
    if ($t.mesAno -match '^([a-z]{3})-(\d{4})$') {
        $mon = $Matches[1]; $yr = $Matches[2]
        if ($monthNumPt.ContainsKey($mon)) {
            $key = "$yr-$($monthNumPt[$mon])"
            $mesAnoToKey[$t.mesAno] = $key
            if (-not $seenMonths.Contains($key)) { $seenMonths.Add($key) }
        }
    }
}
$monthKeysActive = $seenMonths | Sort-Object
$monthLabelsActive = @{}
foreach ($mk in $monthKeysActive) {
    $mon = ($mesAnoToKey.GetEnumerator() | Where-Object { $_.Value -eq $mk } | Select-Object -First 1).Key
    $monPrefix = $mon.Substring(0,3)
    $monthLabelsActive[$mk] = $monthLabelsPt[$monPrefix]
}

# classify every transaction (single source of truth for all totals)
$transactions = New-Object System.Collections.Generic.List[object]
foreach ($t in $rawTx) {
    $categoriaLabel = 'Outros'; $categoriaColor = '#898781'; $macroMatch = $null
    if ($t.idCode) { foreach ($m in $macroDefs) { if ($t.idCode.StartsWith($m.id)) { $macroMatch = $m; break } } }
    if ($macroMatch) { $categoriaLabel = $macroMatch.name; $categoriaColor = $macroMatch.color }
    elseif ($t.ref -match 'APLICA') { $categoriaLabel = 'Aplica' + [char]0x00E7 + [char]0x00E3 + 'o Financeira'; $categoriaColor = '#4a3aa7' }
    elseif ($t.ref -match 'RESGATE') { $categoriaLabel = 'Resgate'; $categoriaColor = '#4a3aa7' }
    elseif ($t.ref -match 'RENDIMENTO') { $categoriaLabel = 'Rendimento Financeiro'; $categoriaColor = '#008300' }
    $isExpense = [bool]$macroMatch
    $subName = $null
    if ($isExpense -and $idToDesc.ContainsKey($t.idCode)) { $subName = $idToDesc[$t.idCode] }
    # a "categoria" exibida por lancamento segue a coluna DESCRICAO da planilha
    # (nivel do ID especifico, ex: "Salarios"); o macro-grupo (Mao de Obra etc.)
    # fica reservado para os graficos agregados, via macroId.
    $displayCategoria = $categoriaLabel
    $displayColor = $categoriaColor
    if ($isExpense) {
        $displayCategoria = $(if ($subName) { $subName } else { $macroMatch.name })
        $displayColor = $macroMatch.color
    }
    $transactions.Add([pscustomobject]@{
        data = $t.pagamento; mesAno = $t.mesAno; mesKey = $mesAnoToKey[$t.mesAno]
        fornecedor = $t.fornecedor; historico = $t.historico; idCode = $t.idCode
        subCategoria = $subName; macroId = $(if ($macroMatch) { $macroMatch.id } else { $null })
        macroName = $(if ($macroMatch) { $macroMatch.name } else { $null })
        categoria = $displayCategoria; categoriaColor = $displayColor
        isExpense = $isExpense; valor = [math]::Round($t.valorPago, 2)
    })
}
$expenseTx = $transactions | Where-Object { $_.isExpense }

# Totais agregados (macro-categoria, sub-categoria, KPIs) vem da propria DRF da
# planilha -- nao de uma soma manual do razao -- para garantir que os numeros do
# site sejam exatamente os que a controladoria ve na planilha.
$categoryById = @{}
foreach ($cat in $categories) { $categoryById[$cat.id] = $cat }

$macroCategories = New-Object System.Collections.Generic.List[object]
foreach ($m in $macroDefs) {
    $cat = $categoryById[$m.id]
    $months = [ordered]@{}
    foreach ($mk in $monthKeysActive) { $months[$mk] = $(if ($cat) { $cat.months2026[$mk] } else { 0 }) }
    $macroCategories.Add([pscustomobject]@{ id = $m.id; name = $m.name; color = $m.color; total = $(if ($cat) { $cat.total } else { 0 }); months = $months })
}
$macroTotal = ($macroCategories | Measure-Object -Property total -Sum).Sum

$subCategories = New-Object System.Collections.Generic.List[object]
foreach ($m in $macroDefs) {
    $prefix = $m.id + '.'
    $subs = $categories | Where-Object { $_.id.StartsWith($prefix) -and ($_.id.Substring($prefix.Length) -notmatch '\.') -and $_.total -ne 0 }
    foreach ($s in $subs) {
        $subCategories.Add([pscustomobject]@{ id = $s.id; name = $(if ($s.desc) { $s.desc.Trim() } else { $s.id }); macroId = $m.id; macroName = $m.name; color = $m.color; total = $s.total })
    }
}

$fornecedoresAtivos = ($expenseTx | Select-Object -ExpandProperty fornecedor -Unique).Count
$topFornecedores = $expenseTx | Group-Object fornecedor | ForEach-Object {
    [pscustomobject]@{ fornecedor = $_.Name; total = [math]::Round((($_.Group | Measure-Object -Property valor -Sum).Sum), 2); qtd = $_.Count }
} | Sort-Object total -Descending | Select-Object -First 10

$periodo = if ($monthKeysActive.Count -gt 0) {
    $first = $monthLabelsActive[$monthKeysActive[0]]; $last = $monthLabelsActive[$monthKeysActive[-1]]
    $yr = $monthKeysActive[0].Substring(0,4)
    "$first a $last de $yr"
} else { "-" }

$site = [ordered]@{
    meta = [ordered]@{
        generatedAt = (Get-Date).ToString('dd/MM/yyyy HH:mm')
        sourceFile = [System.IO.Path]::GetFileName($Xlsx)
        periodo = $periodo
        monthKeys = $monthKeysActive
        monthLabels = $monthLabelsActive
    }
    kpis = [ordered]@{
        receitaLiquida = TotalFor '10.1'
        custosReceitaCanteiro = TotalFor '2.1'
        custosOperacionais = [math]::Round($macroTotal, 2)
        custoCanteiro = TotalFor '2.'
        resultado = $(if ($summaryTotals.ContainsKey('3.')) { $summaryTotals['3.'] } else { -$macroTotal })
        margemCanteiro = $(if ($summaryTotals.ContainsKey('4.')) { $summaryTotals['4.'] } else { 0 })
    }
    macroCategories = $macroCategories
    subCategories = $subCategories
    topFornecedores = $topFornecedores
    transactions = $transactions
}

$json = $site | ConvertTo-Json -Depth 6
$json = $json.Replace('</script', '<\/script')

$templatePath = Join-Path $PSScriptRoot "template.html"
$tpl = [System.IO.File]::ReadAllText($templatePath, $utf8)
$logoB64Path = Join-Path $PSScriptRoot "logo_base64.txt"
$logoB64 = [System.IO.File]::ReadAllText($logoB64Path, [System.Text.Encoding]::ASCII)
$final = $tpl.Replace('__SITE_DATA_JSON__', $json).Replace('__LOGO_BASE64__', $logoB64)
$outPath = Join-Path $ProjectDir "painel-custos-obra20.html"
[System.IO.File]::WriteAllText($outPath, $final, $utf8)

Remove-Item -Recurse -Force $work
Write-Host ("Painel atualizado: " + $outPath)
Write-Host ("Receita Liquida: R$ " + $site.kpis.receitaLiquida)
Write-Host ("Custos Operacionais: R$ " + $site.kpis.custosOperacionais + " | Custos Receita Canteiro: R$ " + $site.kpis.custosReceitaCanteiro)
Write-Host ("Resultado: R$ " + $site.kpis.resultado + " | Margem Canteiro: " + ([math]::Round($site.kpis.margemCanteiro * 100, 2)) + "%")
Write-Host ("Lancamentos: " + $expenseTx.Count + " | Fornecedores: " + $fornecedoresAtivos)
