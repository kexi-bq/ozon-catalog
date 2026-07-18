param(
    [string]$CatalogPath = 'data/catalog_500_exact_match.xlsx',
    [string]$SpecsCsv = 'data/catalog_500_direct_source_specs_2026-07-18.csv',
    [string]$ReportCsv = 'data/catalog_500_direct_source_apply_2026-07-18.csv'
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ColumnNumber([string]$Reference) {
    $number = 0
    foreach ($letter in ([regex]::Match($Reference, '^[A-Z]+')).Value.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$letter - 64)
    }
    $number
}

function Get-OrCreateCell($Sheet, $Namespace, [int]$Row, [int]$Column) {
    $rowNode = $Sheet.SelectSingleNode("//x:sheetData/x:row[@r='$Row']", $Namespace)
    $cell = $rowNode.SelectNodes('./x:c', $Namespace) |
        Where-Object { (Get-ColumnNumber $_.r) -eq $Column } | Select-Object -First 1
    if ($cell) { return $cell }
    $cell = $Sheet.CreateElement('c', $Sheet.DocumentElement.NamespaceURI)
    $columnName = ''
    $value = $Column
    while ($value -gt 0) {
        $value--
        $columnName = [char][int](65 + ($value % 26)) + $columnName
        $value = [math]::Floor($value / 26)
    }
    $cell.SetAttribute('r', "$columnName$Row")
    [void]$rowNode.AppendChild($cell)
    $cell
}

function Clear-Cell($Sheet, $Namespace, [int]$Row, [int]$Column) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $reference = $cell.r
    $cell.RemoveAll()
    $cell.SetAttribute('r', $reference)
}

function Set-NumericCell($Sheet, $Namespace, [int]$Row, [int]$Column, [decimal]$Value) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $reference = $cell.r
    $cell.RemoveAll()
    $cell.SetAttribute('r', $reference)
    $node = $Sheet.CreateElement('v', $Sheet.DocumentElement.NamespaceURI)
    $node.InnerText = $Value.ToString([Globalization.CultureInfo]::InvariantCulture)
    [void]$cell.AppendChild($node)
}

function Set-TextCell($Sheet, $Namespace, [int]$Row, [int]$Column, [string]$Value) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $reference = $cell.r
    $cell.RemoveAll()
    $cell.SetAttribute('r', $reference)
    $cell.SetAttribute('t', 'inlineStr')
    $inline = $Sheet.CreateElement('is', $Sheet.DocumentElement.NamespaceURI)
    $text = $Sheet.CreateElement('t', $Sheet.DocumentElement.NamespaceURI)
    $text.InnerText = $Value
    [void]$inline.AppendChild($text)
    [void]$cell.AppendChild($inline)
}

function Parse-Decimal([string]$Value) {
    [decimal]::Parse(($Value -replace '\s', '' -replace ',', '.'), [Globalization.CultureInfo]::InvariantCulture)
}

function Unit-Factor([string]$Unit) {
    switch ($Unit.ToLowerInvariant()) {
        'м' { return 1000 }
        'см' { return 10 }
        default { return 1 }
    }
}

function Parse-Weight([string]$Text) {
    $match = [regex]::Match($Text, '(?i)(\d+(?:[.,]\d+)?)\s*(кг|г)\b')
    if (!$match.Success) { return $null }
    $weight = Parse-Decimal $match.Groups[1].Value
    if ($match.Groups[2].Value.ToLowerInvariant() -eq 'кг') { $weight *= 1000 }
    $weight
}

function Parse-Dimensions($Spec) {
    $result = @{}
    $evidence = if ($Spec.dimension_facts) { $Spec.dimension_facts } else { $Spec.title_dimensions }
    if (!$evidence) { return $result }

    $triple = [regex]::Match($evidence, '(?i)(\d+(?:[.,]\d+)?)\s*[xх×*]\s*(\d+(?:[.,]\d+)?)\s*[xх×*]\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м)')
    if ($triple.Success) {
        $factor = Unit-Factor $triple.Groups[4].Value
        $result.length = (Parse-Decimal $triple.Groups[1].Value) * $factor
        $result.width = (Parse-Decimal $triple.Groups[2].Value) * $factor
        $result.height = (Parse-Decimal $triple.Groups[3].Value) * $factor
        return $result
    }

    $pair = [regex]::Match($evidence, '(?i)(\d+(?:[.,]\d+)?)\s*[xх×*]\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м)')
    if ($pair.Success) {
        $factor = Unit-Factor $pair.Groups[3].Value
        $result.length = (Parse-Decimal $pair.Groups[1].Value) * $factor
        $result.width = (Parse-Decimal $pair.Groups[2].Value) * $factor
    }

    foreach ($field in @('длина','ширина','высота')) {
        $match = [regex]::Match($evidence, "(?i)$field\s*:\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м)")
        if ($match.Success) {
            $name = switch ($field) { 'длина' {'length'}; 'ширина' {'width'}; 'высота' {'height'} }
            $result[$name] = (Parse-Decimal $match.Groups[1].Value) * (Unit-Factor $match.Groups[2].Value)
        }
    }
    $result
}

$specs = Import-Csv $SpecsCsv
$resolved = (Resolve-Path $CatalogPath).Path
$backup = Join-Path (Split-Path $resolved) 'catalog_500_exact_match.before_direct_specs_2026-07-18.xlsx'
Copy-Item $resolved $backup -Force
$changes = @()

$archive = [IO.Compression.ZipFile]::Open($resolved, [IO.Compression.ZipArchiveMode]::Update)
try {
    $sharedEntry = $archive.GetEntry('xl/sharedStrings.xml')
    $reader = [IO.StreamReader]::new($sharedEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$shared = $reader.ReadToEnd(); $reader.Dispose()
    $sharedNs = [Xml.XmlNamespaceManager]::new($shared.NameTable)
    $sharedNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $strings = @($shared.SelectNodes('//x:si', $sharedNs) | ForEach-Object {
        ($_.SelectNodes('.//x:t', $sharedNs) | ForEach-Object InnerText) -join ''
    })

    $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    $reader = [IO.StreamReader]::new($sheetEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$sheet = $reader.ReadToEnd(); $reader.Dispose()
    $sheetNs = [Xml.XmlNamespaceManager]::new($sheet.NameTable)
    $sheetNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $columns = @{}
    foreach ($cell in $sheet.SelectNodes('//x:sheetData/x:row[@r="1"]/x:c', $sheetNs)) {
        $value = $cell.SelectSingleNode('./x:v', $sheetNs).InnerText
        $name = if ($cell.t -eq 's') { $strings[[int]$value] } else { $value }
        $columns[$name] = Get-ColumnNumber $cell.r
    }

    foreach ($spec in $specs) {
        $row = [int]$spec.index + 1
        foreach ($field in @('weight','source_weight_text','length','width','height','source_dimensions_text')) {
            Clear-Cell $sheet $sheetNs $row $columns[$field]
        }

        $weight = Parse-Weight $spec.weight
        if ($null -ne $weight -and $weight -gt 0) {
            Set-NumericCell $sheet $sheetNs $row $columns.weight $weight
            Set-TextCell $sheet $sheetNs $row $columns.source_weight_text "$weight г"
        }

        $dimensions = Parse-Dimensions $spec
        foreach ($field in @('length','width','height')) {
            if ($dimensions.ContainsKey($field) -and $dimensions[$field] -gt 0) {
                Set-NumericCell $sheet $sheetNs $row $columns[$field] $dimensions[$field]
            }
        }
        $dimensionEvidence = if ($spec.dimension_facts) { $spec.dimension_facts } else { $spec.title_dimensions }
        if ($dimensionEvidence) {
            Set-TextCell $sheet $sheetNs $row $columns.source_dimensions_text $dimensionEvidence
        }

        $changes += [pscustomobject][ordered]@{
            index=$spec.index; offer_id=$spec.offer_id; weight_g=$weight
            length_mm=$dimensions['length']; width_mm=$dimensions['width']; height_mm=$dimensions['height']
            dimension_evidence=$dimensionEvidence; url=$spec.url
        }
    }

    $sheetEntry.Delete()
    $newEntry = $archive.CreateEntry('xl/worksheets/sheet1.xml', [IO.Compression.CompressionLevel]::Optimal)
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $writer = [Xml.XmlWriter]::Create($newEntry.Open(), $settings)
    $sheet.Save($writer); $writer.Dispose()
}
finally {
    $archive.Dispose()
}

$changes | Export-Csv $ReportCsv -NoTypeInformation -Encoding UTF8
Write-Output "Weights retained from direct source: $(($changes | Where-Object weight_g).Count)"
Write-Output "Rows with at least one sourced dimension: $(($changes | Where-Object { $_.length_mm -or $_.width_mm -or $_.height_mm }).Count)"
