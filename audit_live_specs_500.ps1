param(
    [string]$LinkAuditCsv = 'data/catalog_500_live_link_audit_2026-07-18.csv',
    [string]$HtmlDirectory = (Join-Path $env:TEMP 'ozon_catalog_link_audit'),
    [string]$OutputCsv = 'data/catalog_500_direct_source_specs_2026-07-18.csv'
)

function Plain-Text([string]$Html) {
    $text = $Html -replace '(?is)<script.*?</script>|<style.*?</style>', ' '
    $text = $text -replace '(?i)</?(p|div|li|br|tr|td|th|h[1-6])[^>]*>', "`n"
    $text = $text -replace '(?s)<[^>]+>', ' '
    $text = [Net.WebUtility]::HtmlDecode($text)
    ($text -replace '[\u00a0\t ]+', ' ' -replace '(?m)^\s+|\s+$', '')
}

function Extract-LabeledValues([string]$Text, [string]$Labels) {
    $values = foreach ($match in [regex]::Matches(
        $Text,
        "(?im)(?:$Labels)\s*[:\-]\s*([^\n]{1,160})"
    )) {
        ($match.Groups[1].Value -replace '\s+', ' ').Trim(' ', '.', ';')
    }
    @($values | Where-Object { $_ } | Select-Object -Unique) -join ' | '
}

function Extract-LabeledFacts([string]$Text, [string]$Labels) {
    $facts = foreach ($match in [regex]::Matches(
        $Text,
        "(?im)($Labels)\s*[:\-]\s*([^\n]{1,160})"
    )) {
        $label = ($match.Groups[1].Value -replace '\s+', ' ').Trim()
        $value = ($match.Groups[2].Value -replace '\s+', ' ').Trim(' ', '.', ';')
        if ($value) { "${label}: $value" }
    }
    @($facts | Select-Object -Unique) -join ' | '
}

$links = Import-Csv $LinkAuditCsv
$report = foreach ($link in $links) {
    $index = [int]$link.index
    $path = Join-Path $HtmlDirectory ('{0:D3}.html' -f $index)
    if (!(Test-Path $path)) {
        [pscustomobject]@{index=$index;offer_id=$link.offer_id;title=$link.catalog_title;dimensions='';weight='';page_status='missing';url=$link.url}
        continue
    }
    $text = Plain-Text (Get-Content $path -Raw -Encoding UTF8)
    $visibleWeight = Extract-LabeledValues $text 'вес(?:\s+(?:нетто|изделия))?|масса'
    $hiddenWeights = [regex]::Matches((Get-Content $path -Raw -Encoding UTF8), '&quot;weight&quot;:([0-9.]+)')
    $weight = $visibleWeight
    $weightSource = if ($visibleWeight) { 'visible_page_text' } else { '' }
    if (!$weight -and $hiddenWeights.Count -ge 2) {
        $weight = "$($hiddenWeights[1].Groups[1].Value) г"
        $weightSource = 'main_product_data_offer'
    }
    $dimensions = Extract-LabeledValues $text 'габарит(?:ные)?\s+размеры|габариты|размеры?|длина|ширина|высота|диаметр'
    $dimensionFacts = Extract-LabeledFacts $text 'габарит(?:ные)?\s+размеры|габариты|размеры?|длина|ширина|высота|диаметр'
    $titleDimensions = ''
    if (!$dimensions -and $link.catalog_title -match '(?i)\d+(?:[.,]\d+)?\s*[xх×]\s*\d+(?:[.,]\d+)?(?:\s*[xх×]\s*\d+(?:[.,]\d+)?)?\s*(?:мм|см|м)\b') {
        $titleDimensions = $matches[0]
    }
    [pscustomobject][ordered]@{
        index=$index
        offer_id=$link.offer_id
        title=$link.catalog_title
        dimensions=$dimensions
        dimension_facts=$dimensionFacts
        title_dimensions=$titleDimensions
        weight=$weight
        weight_source=$weightSource
        page_status='checked'
        url=$link.url
    }
}

$report | Export-Csv $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Output "Pages: $($report.Count)"
Write-Output "With direct dimension facts: $(($report | Where-Object dimensions).Count)"
Write-Output "With direct weight facts: $(($report | Where-Object weight).Count)"
