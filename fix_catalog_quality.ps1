param(
    [string]$CatalogPath = "data/catalog_500_exact_match.xlsx",
    [string]$InputCsv = "data/catalog_500_classification_audit_input.csv",
    [string]$ReportCsv = "data/catalog_500_quality_fix_2026-07-18.csv"
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Repair-Text([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    return [Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding(1251).GetBytes($Text))
}

function Get-ColumnNumber([string]$Reference) {
    $number = 0
    foreach ($letter in ([regex]::Match($Reference, '^[A-Z]+')).Value.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$letter - 64)
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

function Get-OrCreateCell($Sheet, $Namespace, [int]$Row, [int]$Column) {
    $rowNode = $Sheet.SelectSingleNode("//x:sheetData/x:row[@r='$Row']", $Namespace)
    $cell = $rowNode.SelectNodes('./x:c', $Namespace) |
        Where-Object { (Get-ColumnNumber $_.r) -eq $Column } |
        Select-Object -First 1
    if ($cell) { return $cell }

    $cell = $Sheet.CreateElement('c', $Sheet.DocumentElement.NamespaceURI)
    $cell.SetAttribute('r', "$(Get-ColumnName $Column)$Row")
    $nextCell = $rowNode.SelectNodes('./x:c', $Namespace) |
        Where-Object { (Get-ColumnNumber $_.r) -gt $Column } |
        Select-Object -First 1
    if ($nextCell) { [void]$rowNode.InsertBefore($cell, $nextCell) }
    else { [void]$rowNode.AppendChild($cell) }
    return $cell
}

function Set-NumericCell($Sheet, $Namespace, [int]$Row, [int]$Column, [decimal]$Value) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $cell.RemoveAll()
    $cell.SetAttribute('r', "$(Get-ColumnName $Column)$Row")
    $valueNode = $Sheet.CreateElement('v', $Sheet.DocumentElement.NamespaceURI)
    $valueNode.InnerText = $Value.ToString([Globalization.CultureInfo]::InvariantCulture)
    [void]$cell.AppendChild($valueNode)
}

function Set-TextCell($Sheet, $Namespace, [int]$Row, [int]$Column, [string]$Value) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $cell.RemoveAll()
    $cell.SetAttribute('r', "$(Get-ColumnName $Column)$Row")
    $cell.SetAttribute('t', 'inlineStr')
    $inline = $Sheet.CreateElement('is', $Sheet.DocumentElement.NamespaceURI)
    $text = $Sheet.CreateElement('t', $Sheet.DocumentElement.NamespaceURI)
    $text.InnerText = $Value
    [void]$inline.AppendChild($text)
    [void]$cell.AppendChild($inline)
}

function Clear-Cell($Sheet, $Namespace, [int]$Row, [int]$Column) {
    $cell = Get-OrCreateCell $Sheet $Namespace $Row $Column
    $cell.RemoveAll()
    $cell.SetAttribute('r', "$(Get-ColumnName $Column)$Row")
}

function Get-TnvedAssignment([string]$Title, [string]$TypeId) {
    $text = $Title.ToLowerInvariant().Replace([char]0x451, [char]0x435)
    $rules = @(
        @('filter', '8421230000', 'high', 'engine oil/fuel filter', 'фильтр|фильтрующая сетка'),
        @('fastener_bolt', '7318158900', 'high', 'iron or steel bolt/screw', 'болт|винт '),
        @('fastener_nut', '7318169109', 'high', 'iron or steel nut', 'гайка'),
        @('fastener_washer', '7318220008', 'high', 'iron or steel washer', 'шайба'),
        @('fastener_pin', '7318240008', 'high', 'cotter or retaining pin', 'шплинт|штифт'),
        @('switch', '8536508008', 'high', 'electrical switch', 'выключатель|переключатель|тумблер'),
        @('fuse_holder', '8536908500', 'medium', 'other low-voltage electrical apparatus', 'держатель предохранителя|блок предохранителей'),
        @('fuse', '8536109000', 'high', 'electrical fuse', 'предохранитель'),
        @('terminal_block', '8536901000', 'high', 'connection and contact element for wires', 'колодка распределительная|клеммная колодка'),
        @('control_panel', '8537109800', 'medium', 'electrical control panel under 1000 V', 'панель контроля заряда'),
        @('socket', '8536699008', 'medium', 'electrical socket or connector', 'разъем|прикуриватель|розетк'),
        @('solar_fan', '8414598000', 'medium', 'other fan', 'вентилятор на солнечной'),
        @('wiper_motor', '8512400009', 'high', 'vehicle windscreen wiper equipment', 'электропривод стеклоочистителя'),
        @('led_lamp', '9405420039', 'medium', 'LED luminaire', 'светодиодн.*светильник|светильник.*светодиодн'),
        @('lamp', '9405490039', 'medium', 'other electric luminaire', 'светильник|плафон|лампа'),
        @('gloves', '6216000000', 'high', 'gloves', 'перчатк'),
        @('rope', '5607491100', 'medium', 'synthetic rope', 'намотка|конец александрова|строп'),
        @('bell', '8306100000', 'high', 'bell', 'рында'),
        @('horn', '8512309009', 'medium', 'electric sound signalling equipment', 'горн|мегафон'),
        @('whistle', '9208900000', 'medium', 'mouth-blown sound instrument', 'свисток|боцманская дудка'),
        @('clock', '9104000000', 'medium', 'instrument-panel clock', 'часы'),
        @('voltmeter', '9030339900', 'medium', 'voltmeter', 'вольтметр'),
        @('level_gauge', '9026108900', 'medium', 'liquid level indicator', 'указатель уровня|датчик уровня|сигнализатор уровня'),
        @('propeller', '8487109000', 'medium', 'ship or boat propeller', 'гребной винт'),
        @('impeller', '8409910008', 'medium', 'part for spark-ignition marine engine', 'крыльчатка|детал.*двигател|вал гребной|ремень грм'),
        @('gasket', '8484900000', 'medium', 'gasket or seal', 'прокладка'),
        @('oil_seal', '8484200000', 'medium', 'mechanical seal', 'сальник'),
        @('pump', '8413603900', 'medium', 'liquid pump', 'помпа|насос'),
        @('powered_sump', '8413810000', 'medium', 'other liquid pump', 'сточный бак для душа'),
        @('plastic_tank', '3923301090', 'medium', 'plastic tank or container', 'бак топливный|водяной бак|сточной бак'),
        @('paint', '3208909109', 'medium', 'paint or coating', 'краска|грунт'),
        @('keychain', '3926909709', 'medium', 'plastic or composite accessory', 'брелок'),
        @('plastic_tableware', '3924100000', 'medium', 'plastic tableware', 'набор бокалов|набор кружек|набор мисок'),
        @('plastic_hatch', '3926909709', 'medium', 'plastic article', 'утка пластиковая|лючок|люк |органайзер|пробк|кожух|держатель емкости|крышка защитная'),
        @('trailer_roller', '3926909709', 'medium', 'plastic trailer roller', 'ролик килевой|ролик носовой|ролик подкильный|ролик боковой|запасной роульс'),
        @('trim_actuator', '8479899707', 'medium', 'other mechanical appliance', 'привод транцевых плит'),
        @('oar', '7616999008', 'low', 'aluminium article', 'весло|багор'),
        @('metal_fitting', '7326909807', 'low', 'other iron or steel article', 'кронштейн|держатель|крепление|ролик|упор|ручка|решетка|застежк|хомут'),
        @('engine_part', '8409910008', 'low', 'other marine engine part', 'suzuki|yamaha|honda|mercury|marine rocket|tohatsu|recmar')
    )
    foreach ($rule in $rules) {
        if ($text -match $rule[4]) {
            return [pscustomobject]@{ code=$rule[1]; confidence=$rule[2]; rule=$rule[3] }
        }
    }

    $fallback = switch ($TypeId) {
        '94373' { '8536508008' }
        '95674' { '8536109000' }
        '94698' { '9031809800' }
        '971322244' { '3926909709' }
        '970827605' { '3926909709' }
        default { '7326909807' }
    }
    return [pscustomobject]@{ code=$fallback; confidence='low'; rule="type fallback $TypeId" }
}

$weights = @{
    11=50; 13=2190; 14=2680; 24=150; 29=100; 36=1270; 37=1260; 38=850;
    39=350; 40=30; 41=410; 42=30; 51=300; 52=1500; 91=11700; 98=730; 99=120
}
$dimensions = @{
    14=@(510,365,230); 15=@(375,280,70); 16=@(600,360,70); 17=@(380,380,70);
    18=@(510,460,70); 24=@(129,134,115); 25=@(500,350,300); 36=@(340,135,190);
    37=@(340,135,190); 38=@(1800,190,61); 48=@(1100,240,240); 52=@(380,280,130);
    69=@(140,140,90); 98=@(48,39,48)
}

$resolved = (Resolve-Path $CatalogPath).Path
$backup = Join-Path (Split-Path $resolved) 'catalog_500_exact_match.before_quality_fix_2026-07-18.xlsx'
Copy-Item $resolved $backup -Force
$sourceRows = Import-Csv $InputCsv
$report = @()

$archive = [IO.Compression.ZipFile]::Open($resolved, [IO.Compression.ZipArchiveMode]::Update)
try {
    $sharedEntry = $archive.GetEntry('xl/sharedStrings.xml')
    $reader = New-Object IO.StreamReader($sharedEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$shared = $reader.ReadToEnd(); $reader.Close()
    $sharedNs = New-Object Xml.XmlNamespaceManager($shared.NameTable)
    $sharedNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $strings = @($shared.SelectNodes('//x:si', $sharedNs) | ForEach-Object {
        ($_.SelectNodes('.//x:t', $sharedNs) | ForEach-Object { $_.InnerText }) -join ''
    })

    $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    $reader = New-Object IO.StreamReader($sheetEntry.Open(), [Text.Encoding]::UTF8)
    [xml]$sheet = $reader.ReadToEnd(); $reader.Close()
    $sheetNs = New-Object Xml.XmlNamespaceManager($sheet.NameTable)
    $sheetNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $columns = @{}
    foreach ($cell in $sheet.SelectNodes('//x:sheetData/x:row[@r="1"]/x:c', $sheetNs)) {
        $value = $cell.SelectSingleNode('./x:v', $sheetNs).InnerText
        $name = if ($cell.t -eq 's') { $strings[[int]$value] } else { $value }
        $columns[$name] = Get-ColumnNumber $cell.r
    }

    foreach ($index in $weights.Keys) {
        Set-NumericCell $sheet $sheetNs ([int]$index + 1) $columns['weight'] $weights[$index]
    }
    foreach ($index in $dimensions.Keys) {
        $values = $dimensions[$index]
        Set-NumericCell $sheet $sheetNs ([int]$index + 1) $columns['length'] $values[0]
        Set-NumericCell $sheet $sheetNs ([int]$index + 1) $columns['width'] $values[1]
        Set-NumericCell $sheet $sheetNs ([int]$index + 1) $columns['height'] $values[2]
    }

    # Product 511 has no supplier-published package dimensions.
    foreach ($field in @('length','width','height')) {
        Clear-Cell $sheet $sheetNs 8 $columns[$field]
    }

    foreach ($item in ($sourceRows | Where-Object { !$_.tnved })) {
        $assignment = Get-TnvedAssignment (Repair-Text $item.title) $item.type_id
        Set-TextCell $sheet $sheetNs ([int]$item.index + 1) $columns['tnved'] $assignment.code
        $report += [pscustomobject][ordered]@{
            index=$item.index; offer_id=(Repair-Text $item.offer_id); title=(Repair-Text $item.title)
            tnved=$assignment.code; confidence=$assignment.confidence; rule=$assignment.rule
        }
    }

    $sheetEntry.Delete()
    $newEntry = $archive.CreateEntry('xl/worksheets/sheet1.xml', [IO.Compression.CompressionLevel]::Optimal)
    $settings = New-Object Xml.XmlWriterSettings
    $settings.Encoding = New-Object Text.UTF8Encoding($false)
    $writer = [Xml.XmlWriter]::Create($newEntry.Open(), $settings)
    $sheet.Save($writer); $writer.Close()
}
finally {
    $archive.Dispose()
}

$report | Export-Csv $ReportCsv -NoTypeInformation -Encoding UTF8
Write-Output "Corrected weights: $($weights.Count)"
Write-Output "Corrected dimensions: $($dimensions.Count)"
Write-Output "Filled TN VED: $($report.Count)"
$report | Group-Object confidence | Sort-Object Name | Format-Table Count, Name -AutoSize

