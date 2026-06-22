Add-Type -AssemblyName System.IO.Compression.FileSystem

function Safe-DeleteFolder {
    param([string]$Path)
    try {
        if (-not [string]::IsNullOrWhiteSpace($Path) -and [System.IO.Directory]::Exists($Path)) {
            [System.IO.Directory]::Delete($Path, $true)
        }
    } catch {}
}

function Decode-Text {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    $t = $Text
    $t = $t -replace '&amp;', '&'
    $t = $t -replace '&lt;', '<'
    $t = $t -replace '&gt;', '>'
    $t = $t -replace '&quot;', '"'
    $t = $t -replace '&apos;', "'"
    $t = $t -replace "`r", " "
    $t = $t -replace "`n", " "
    $t = $t -replace "`t", " "
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

function Get-CellValue {
    param(
        [System.Xml.XmlElement]$Shape,
        [string]$CellName
    )

    $cell = $Shape.SelectSingleNode("./*[local-name()='Cell'][@N='$CellName']")
    if ($null -ne $cell) {
        return $cell.GetAttribute("V")
    }

    return ""
}

function Get-BacnetObjects {
    param([string]$Text)

    $matches = [regex]::Matches($Text, '\b(AI|AO|AV|BI|BO|BV|MI|MO|MV)-?\d+\b', 'IgnoreCase')
    $objects = @()

    foreach ($m in $matches) {
        $obj = $m.Value.ToUpper()

        if ($obj -notmatch '-') {
            $obj = $obj -replace '^([A-Z]+)(\d+)$', '$1-$2'
        }

        $objects += $obj
    }

    return $objects | Select-Object -Unique
}

function Is-BadDescription {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $true }

    $v = $Value.Trim().ToUpper()

    if ($v.Length -lt 2) { return $true }
    if ($v -match '^\d+(\.\d+)?$') { return $true }
    if ($v -match '^(AI|AO|AV|BI|BO|BV|MI|MO|MV)-?\d+$') { return $true }
    if ($v -match '^BR-\d+$') { return $true }

    $badWords = @(
        "RGB",
        "INH",
        "WIDTH",
        "HEIGHT",
        "GUARD",
        "PAR",
        "PNT",
        "USER",
        "THEME",
        "THEMEGUARD",
        "COLOR",
        "FILL",
        "LINE",
        "BEGINX",
        "BEGINY",
        "ENDX",
        "ENDY",
        "CONTROLS",
        "GEOMETRY",
        "CONNECTION",
        "CONNECTIONS"
    )

    if ($badWords -contains $v) { return $true }
    if ($v -match '^RGB[A-Z0-9_]*$') { return $true }
    if ($v -match 'SHEET\.|WIDTH|HEIGHT|GUARD|PAR|PNT|XFTRIGGER|THEMEGUARD|RGB') { return $true }

    return $false
}

function Get-BestDescriptionFromText {
    param([string]$Text)

    $clean = Decode-Text $Text

    $clean = $clean -replace '\bInh\b.*$', ''
    $clean = $clean -replace 'Sheet\.\d+!.*$', ''
    $clean = $clean -replace 'GUARD\(.*$', ''
    $clean = $clean -replace 'PAR\(.*$', ''
    $clean = $clean -replace 'PNT\(.*$', ''
    $clean = $clean -replace '_XFTRIGGER\(.*$', ''

    $clean = $clean -replace '\b(AI|AO|AV|BI|BO|BV|MI|MO|MV)-?\d+\b', ' '
    $clean = $clean -replace '\bBR-\d+\b', ' '

    # Descriptions are single words, usually all caps, often with underscores.
    $matches = [regex]::Matches($clean, '\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*\b')

    $candidates = @()

    foreach ($m in $matches) {
        $value = $m.Value.Trim().ToUpper()

        if (-not (Is-BadDescription $value)) {
            $candidates += $value
        }
    }

    if ($candidates.Count -eq 0) {
        return ""
    }

    $best = $candidates |
        Sort-Object @{
            Expression = { if ($_ -match '_') { 0 } else { 1 } }
            Ascending = $true
        }, @{
            Expression = { $_.Length }
            Ascending = $false
        } |
        Select-Object -First 1

    return $best
}

function Get-ShapeSearchText {
    param([System.Xml.XmlElement]$Shape)

    $parts = @()

    $textNode = $Shape.SelectSingleNode("./*[local-name()='Text']")
    if ($null -ne $textNode) {
        $parts += $textNode.InnerText
    }

    # Use cell V values only. Do not use F formulas.
    $cells = $Shape.SelectNodes(".//*[local-name()='Cell']")
    foreach ($cell in $cells) {
        $v = $cell.GetAttribute("V")
        if (-not [string]::IsNullOrWhiteSpace($v)) {
            $parts += $v
        }
    }

    return Decode-Text ($parts -join " ")
}

Write-Host ""
Write-Host "VisualLogic BACnet Object Exporter - Direct XML Version" -ForegroundColor Cyan
Write-Host "------------------------------------------------------" -ForegroundColor Cyan

$vsdxPath = Read-Host "Enter the full path to the Visio .vsdx file"

if (-not (Test-Path $vsdxPath)) {
    Write-Host "File not found: $vsdxPath" -ForegroundColor Red
    exit
}

$defaultOutputFolder = Join-Path ([System.IO.Path]::GetDirectoryName($vsdxPath)) "BacnetExport"
$outputFolder = Read-Host "Enter output folder, or press Enter for $defaultOutputFolder"

if ([string]::IsNullOrWhiteSpace($outputFolder)) {
    $outputFolder = $defaultOutputFolder
}

if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$tempFolder = Join-Path $env:TEMP "VisioBacnetExtract_$timestamp"

$excelPath = Join-Path $outputFolder "BacnetObjects_$timestamp.xlsx"
$xmlPath = Join-Path $outputFolder "BacnetObjects_$timestamp.xml"
$diagPath = Join-Path $outputFolder "BacnetDiagnostic_$timestamp.txt"

Safe-DeleteFolder $tempFolder
New-Item -ItemType Directory -Path $tempFolder | Out-Null

Write-Host ""
Write-Host "Extracting .vsdx XML..." -ForegroundColor Cyan

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($vsdxPath, $tempFolder)
}
catch {
    Write-Host "Could not extract .vsdx file." -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit
}

$pageXmlFolder = Join-Path $tempFolder "visio\pages"

if (-not (Test-Path $pageXmlFolder)) {
    Write-Host "Could not find visio\pages." -ForegroundColor Red
    Safe-DeleteFolder $tempFolder
    exit
}

$pageFiles = Get-ChildItem -Path $pageXmlFolder -Filter "*.xml" -File

$rawRecords = @()

foreach ($pageFile in $pageFiles) {
    Write-Host "Scanning $($pageFile.Name)..." -ForegroundColor Yellow

    try {
        [xml]$pageXml = Get-Content $pageFile.FullName -Raw
    }
    catch {
        continue
    }

    $shapes = $pageXml.SelectNodes("//*[local-name()='Shape']")

    foreach ($shape in $shapes) {
        $searchText = Get-ShapeSearchText -Shape $shape

        if ([string]::IsNullOrWhiteSpace($searchText)) {
            continue
        }

        $objects = Get-BacnetObjects -Text $searchText

        if ($objects.Count -gt 0) {
            $description = Get-BestDescriptionFromText -Text $searchText

            foreach ($object in $objects) {
                $rawRecords += [PSCustomObject]@{
                    Object      = $object
                    Description = $description
                    Page        = $pageFile.BaseName
                    ShapeID     = $shape.ID
                    RawText     = $searchText
                }
            }
        }
    }
}

$finalPairs = $rawRecords |
    Group-Object Object |
    ForEach-Object {
        $_.Group |
            Sort-Object @{
                Expression = {
                    if ([string]::IsNullOrWhiteSpace($_.Description)) { 1 } else { 0 }
                }
                Ascending = $true
            }, @{
                Expression = {
                    if ($_.Description -match '_') { 0 } else { 1 }
                }
                Ascending = $true
            }, @{
                Expression = { $_.Description.Length }
                Ascending = $false
            } |
            Select-Object -First 1
    } |
    Sort-Object {
        if ($_.Object -match '^([A-Z]+)-(\d+)$') {
            "{0}-{1:D6}" -f $matches[1], [int]$matches[2]
        }
        else {
            $_.Object
        }
    }

$rawRecords |
    Sort-Object Object, Page, ShapeID |
    ForEach-Object {
        "Object: $($_.Object) | Description: $($_.Description) | Page: $($_.Page) | ShapeID: $($_.ShapeID) | RawText: $($_.RawText)"
    } |
    Set-Content -Path $diagPath -Encoding UTF8

Write-Host ""
Write-Host "Creating XML..." -ForegroundColor Cyan

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.Encoding = [System.Text.Encoding]::UTF8

$writer = [System.Xml.XmlWriter]::Create($xmlPath, $settings)

$writer.WriteStartDocument()
$writer.WriteStartElement("BacnetObjects")

foreach ($item in $finalPairs) {
    $writer.WriteStartElement("BacnetObject")
    $writer.WriteElementString("Object", $item.Object)
    $writer.WriteElementString("Description", $item.Description)
    $writer.WriteEndElement()
}

$writer.WriteEndElement()
$writer.WriteEndDocument()
$writer.Close()

Write-Host "Creating Excel..." -ForegroundColor Cyan

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "BACnet Objects"

    $worksheet.Cells.Item(1, 1) = "Object"
    $worksheet.Cells.Item(1, 2) = "Description"

    $worksheet.Range("A1:B1").Font.Bold = $true

    $row = 2

    foreach ($item in $finalPairs) {
        $worksheet.Cells.Item($row, 1) = $item.Object
        $worksheet.Cells.Item($row, 2) = $item.Description
        $row++
    }

    $worksheet.UsedRange.Columns.AutoFit() | Out-Null
    $workbook.SaveAs($excelPath, 51)

    $workbook.Close($true)
    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Write-Host "Excel failed. Creating CSV instead." -ForegroundColor Yellow

    $csvPath = Join-Path $outputFolder "BacnetObjects_$timestamp.csv"

    $finalPairs |
        Select-Object Object, Description |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "CSV saved to: $csvPath" -ForegroundColor Green
}

Safe-DeleteFolder $tempFolder

Write-Host ""
Write-Host "Export complete." -ForegroundColor Cyan
Write-Host "Unique objects found: $($finalPairs.Count)"
Write-Host "Excel saved to: $excelPath" -ForegroundColor Green
Write-Host "XML saved to: $xmlPath" -ForegroundColor Green
Write-Host "Diagnostic saved to: $diagPath" -ForegroundColor Green
