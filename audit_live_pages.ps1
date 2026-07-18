param(
    [string]$InputCsv = "data/catalog_500_link_audit_input.csv",
    [string]$PagesDir = (Join-Path $env:TEMP "ozon_catalog_link_audit"),
    [string]$OutputCsv = "data/catalog_500_live_link_audit_2026-07-18.csv"
)

function Repair-Utf8Text([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    return [Text.Encoding]::UTF8.GetString(
        [Text.Encoding]::GetEncoding(1251).GetBytes($Text)
    )
}

function ConvertFrom-HtmlText([string]$Text) {
    $value = [Net.WebUtility]::HtmlDecode($Text)
    $value = [regex]::Replace($value, '<[^>]+>', ' ')
    return [regex]::Replace($value, '\s+', ' ').Trim()
}

function Normalize-Title([string]$Text) {
    return (ConvertFrom-HtmlText $Text).ToLowerInvariant().Replace([char]0x451, [char]0x435)
}

$results = foreach ($row in (Import-Csv $InputCsv)) {
    $pagePath = Join-Path $PagesDir ('{0:D3}.html' -f [int]$row.index)
    $html = [IO.File]::ReadAllText($pagePath, [Text.Encoding]::UTF8)
    $site = ([uri]$row.external_url).Host
    $currentTitle = ''
    $currentPrice = ''

    if ($site -eq 'moreman.ru') {
        $match = [regex]::Match($html, '<h1[^>]*class="h2"[^>]*>(.*?)</h1>', 'Singleline')
        if ($match.Success) { $currentTitle = ConvertFrom-HtmlText $match.Groups[1].Value }

        $match = [regex]::Match($html, '<span[^>]*class="card__price"[^>]*>(.*?)</span>', 'Singleline')
        if ($match.Success) {
            $currentPrice = [regex]::Replace(
                (ConvertFrom-HtmlText $match.Groups[1].Value),
                '[^0-9,.]',
                ''
            ).Replace(',', '.')
        }
    }
    else {
        $namePosition = $html.IndexOf('id="nomenclature_name"')
        $start = if ($namePosition -ge 0) {
            $html.LastIndexOf('<h1', $namePosition)
        }
        else {
            0
        }
        if ($start -lt 0) { $start = 0 }
        $productHtml = $html.Substring($start)

        $match = [regex]::Match($productHtml, '<h1[^>]*itemprop="name"[^>]*>(.*?)</h1>', 'Singleline')
        if ($match.Success) { $currentTitle = ConvertFrom-HtmlText $match.Groups[1].Value }

        $match = [regex]::Match(
            $productHtml,
            'itemprop=\\?"price\\?"\s+content=\\?"([0-9.,]+)',
            'Singleline'
        )
        if ($match.Success) { $currentPrice = $match.Groups[1].Value.Replace(',', '.') }
    }

    $catalogTitle = Repair-Utf8Text $row.live_title
    $expectedSku = Repair-Utf8Text $row.live_sku
    $catalogPrice = [string]$row.live_price
    $titleMatches = (Normalize-Title $currentTitle) -eq (Normalize-Title $catalogTitle)
    $skuMatches = !$expectedSku -or $html.Contains($expectedSku)
    $priceMatches = if ($currentPrice -and $catalogPrice) {
        [decimal]$currentPrice -eq [decimal]$catalogPrice
    }
    else {
        !$currentPrice -and !$catalogPrice
    }

    $status = if ($html.Length -lt 1000) { 'page_error' }
        elseif (!$currentTitle) { 'parse_error' }
        elseif (!$skuMatches) { 'sku_mismatch' }
        elseif (!$titleMatches -and !$priceMatches) { 'title_and_price_mismatch' }
        elseif (!$titleMatches) { 'title_mismatch' }
        elseif (!$currentPrice) { 'price_unavailable' }
        elseif (!$priceMatches) { 'price_mismatch' }
        else { 'ok' }

    [pscustomobject][ordered]@{
        index = $row.index
        offer_id = Repair-Utf8Text $row.offer_id
        code = Repair-Utf8Text $row.code
        site = $site
        status = $status
        catalog_title = $catalogTitle
        live_title_now = $currentTitle
        title_match = $titleMatches
        catalog_source_price = $catalogPrice
        live_price_now = $currentPrice
        price_match = $priceMatches
        expected_sku = $expectedSku
        sku_found = $skuMatches
        url = $row.external_url
    }
}

$results | Export-Csv $OutputCsv -NoTypeInformation -Encoding UTF8
$results | Group-Object status | Sort-Object Count -Descending | Format-Table Count, Name -AutoSize
