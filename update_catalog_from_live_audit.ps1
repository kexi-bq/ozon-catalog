param(
    [string]$CatalogPath = "data/catalog_500_exact_match.xlsx",
    [string]$AuditPath = "data/catalog_500_live_link_audit_2026-07-18.csv",
    [decimal]$OzonMultiplier = 1.352
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$resolvedCatalog = (Resolve-Path $CatalogPath).Path
$backupPath = Join-Path (Split-Path $resolvedCatalog) "catalog_500_exact_match.before_live_price_update.xlsx"
Copy-Item $resolvedCatalog $backupPath -Force

function Get-ColumnNumber([string]$CellReference) {
    $letters = ([regex]::Match($CellReference, '^[A-Z]+')).Value
    $number = 0
    foreach ($letter in $letters.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$letter - [int][char]'A' + 1)
    }
    return $number
}

function Get-ColumnName([int]$Column) {
    $name = ''
    while ($Column -gt 0) {
        $Column--
        $name = [char](65 + ($Column % 26)) + $name
        $Column = [math]::Floor($Column / 26)
    }
    return $name
}

function Set-NumericCell($Sheet, $Namespace, [int]$Row, [int]$Column, [decimal]$Value) {
    $rowNode = $Sheet.SelectSingleNode("//x:sheetData/x:row[@r='$Row']", $Namespace)
    if (!$rowNode) { throw "Missing worksheet row $Row" }

    $cell = $rowNode.SelectNodes('./x:c', $Namespace) |
        Where-Object { (Get-ColumnNumber $_.r) -eq $Column } |
        Select-Object -First 1
    if (!$cell) {
        $cell = $Sheet.CreateElement('c', $Sheet.DocumentElement.NamespaceURI)
        $cell.SetAttribute('r', "$(Get-ColumnName $Column)$Row")
        $nextCell = $rowNode.SelectNodes('./x:c', $Namespace) |
            Where-Object { (Get-ColumnNumber $_.r) -gt $Column } |
            Select-Object -First 1
        if ($nextCell) {
            [void]$rowNode.InsertBefore($cell, $nextCell)
        }
        else {
            [void]$rowNode.AppendChild($cell)
        }
    }

    if ($cell.HasAttribute('t')) { $cell.RemoveAttribute('t') }
    $valueNode = $cell.SelectSingleNode('./x:v', $Namespace)
    if (!$valueNode) {
        $valueNode = $Sheet.CreateElement('v', $Sheet.DocumentElement.NamespaceURI)
        [void]$cell.AppendChild($valueNode)
    }
    $valueNode.InnerText = $Value.ToString([Globalization.CultureInfo]::InvariantCulture)
}

$archive = [IO.Compression.ZipFile]::Open($resolvedCatalog, [IO.Compression.ZipArchiveMode]::Update)
try {
    $sharedEntry = $archive.GetEntry('xl/sharedStrings.xml')
    $reader = New-Object IO.StreamReader($sharedEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$shared = $reader.ReadToEnd()
    $reader.Close()

    $sharedNs = New-Object Xml.XmlNamespaceManager($shared.NameTable)
    $sharedNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $strings = @($shared.SelectNodes('//x:si', $sharedNs) | ForEach-Object {
        ($_.SelectNodes('.//x:t', $sharedNs) | ForEach-Object { $_.InnerText }) -join ''
    })

    $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    $reader = New-Object IO.StreamReader($sheetEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$sheet = $reader.ReadToEnd()
    $reader.Close()

    $sheetNs = New-Object Xml.XmlNamespaceManager($sheet.NameTable)
    $sheetNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $columns = @{}
    foreach ($cell in $sheet.SelectNodes('//x:sheetData/x:row[@r="1"]/x:c', $sheetNs)) {
        $value = $cell.SelectSingleNode('./x:v', $sheetNs).InnerText
        $name = if ($cell.t -eq 's') { $strings[[int]$value] } else { $value }
        $columns[$name] = Get-ColumnNumber $cell.r
    }

    $updatedPrices = 0
    $disabledWithoutPrice = 0
    foreach ($item in (Import-Csv $AuditPath)) {
        $row = [int]$item.index + 1
        if ($item.live_price_now) {
            $sourcePrice = [decimal]$item.live_price_now
            $ozonPrice = [math]::Floor($sourcePrice * $OzonMultiplier)
            Set-NumericCell $sheet $sheetNs $row $columns['retail_price'] $sourcePrice
            Set-NumericCell $sheet $sheetNs $row $columns['live_price'] $sourcePrice
            Set-NumericCell $sheet $sheetNs $row $columns['ozon_price'] $ozonPrice
            $updatedPrices++
        }
        elseif ($item.status -eq 'price_unavailable') {
            Set-NumericCell $sheet $sheetNs $row $columns['stock_qty'] 0
            $disabledWithoutPrice++
        }
    }

    $sheetEntry.Delete()
    $newEntry = $archive.CreateEntry('xl/worksheets/sheet1.xml', [IO.Compression.CompressionLevel]::Optimal)
    $settings = New-Object Xml.XmlWriterSettings
    $settings.Encoding = New-Object Text.UTF8Encoding($false)
    $settings.Indent = $false
    $writer = [Xml.XmlWriter]::Create($newEntry.Open(), $settings)
    $sheet.Save($writer)
    $writer.Close()

    Write-Output "Updated prices: $updatedPrices"
    Write-Output "Disabled without live price: $disabledWithoutPrice"
    Write-Output "Backup: $backupPath"
}
finally {
    $archive.Dispose()
}
