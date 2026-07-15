[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [ValidateSet('Preview', 'Generate')] [string]$Mode,
    [Parameter(Mandatory = $true)] [string]$InputPath,
    [string]$OutputPath,
    [Parameter(Mandatory = $true)] [string]$SourceSheet,
    [string]$FinalSheetName = 'MergedMappedPool',
    [Parameter(Mandatory = $true)] [string]$SourceKeyColumn,
    [Parameter(Mandatory = $true)] [string]$MappingIdColumn,
    [Parameter(Mandatory = $true)] [string]$EntityColumn,
    [Parameter(Mandatory = $true)] [string]$IndicationColumn,
    [Parameter(Mandatory = $true)] [string]$DiseaseColumn,
    [Parameter(Mandatory = $true)] [string]$StatusColumn,
    [Parameter(Mandatory = $true)] [string]$RationaleColumn,
    [Parameter(Mandatory = $true)] [string]$GenericColumn,
    [Parameter(Mandatory = $true)] [string]$DosageFormColumn,
    [Parameter(Mandatory = $true)] [string]$BrandColumn,
    [string]$MappedValue = 'Mapped',
    [switch]$Confirmed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Delimiter = [string][char]0xFF1B
$GroupSeparator = [char]31

function Join-Unique {
    param([object[]]$Values)
    $seen = @{}
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        if ($null -eq $value) { continue }
        $text = ([string]$value).Trim()
        if ($text -and -not $seen.ContainsKey($text)) {
            $seen[$text] = $true
            [void]$result.Add($text)
        }
    }
    return [string]::Join($Delimiter, $result.ToArray())
}

function Get-UniqueArray {
    param([object[]]$Values)
    $seen = @{}
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        if ($null -eq $value) { continue }
        $text = ([string]$value).Trim()
        if ($text -and -not $seen.ContainsKey($text)) {
            $seen[$text] = $true
            [void]$result.Add($text)
        }
    }
    return @($result.ToArray())
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input workbook was not found: $InputPath" }
if ([IO.Path]::GetExtension($InputPath) -notin '.xlsx', '.xlsm') { throw 'Input workbook must be .xlsx or .xlsm.' }
if ($Mode -eq 'Generate') {
    if (-not $Confirmed) { throw 'Generate mode requires explicit human approval. Run Preview first, then rerun with -Confirmed.' }
    if (-not $OutputPath) { throw 'Generate mode requires OutputPath.' }
    if ([IO.Path]::GetFullPath($InputPath) -eq [IO.Path]::GetFullPath($OutputPath)) { throw 'OutputPath must differ from InputPath.' }
    if (Test-Path -LiteralPath $OutputPath) { throw "Output workbook already exists and will not be overwritten: $OutputPath" }
    $outputDirectory = [IO.Path]::GetDirectoryName($OutputPath)
    if (-not (Test-Path -LiteralPath $outputDirectory)) { throw "Output directory does not exist: $outputDirectory" }
}

$excel = $null
$sourceBook = $null
$outputBook = $null
$createdOutput = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    try { $excel.AutomationSecurity = 3 } catch { }

    $sourceBook = $excel.Workbooks.Open($InputPath, 0, $true)
    try { $source = $sourceBook.Worksheets.Item($SourceSheet) }
    catch { throw "Source sheet was not found: $SourceSheet" }

    $rowCount = [int]$source.UsedRange.Rows.Count
    $columnCount = [int]$source.UsedRange.Columns.Count
    if ($rowCount -lt 2) { throw 'Source sheet has no data rows.' }
    $sourceRange = $source.Range($source.Cells.Item(1, 1), $source.Cells.Item($rowCount, $columnCount))
    $sourceValues = $sourceRange.Value2
    $sourceFormulas = $sourceRange.Formula

    $headers = @{}
    $headerNames = @{}
    for ($column = 1; $column -le $columnCount; $column++) {
        $header = ([string]$sourceValues[1, $column]).Trim()
        if (-not $header -or $headers.ContainsKey($header)) { throw "Header is blank or duplicated at column $column." }
        $headers[$header] = $column
        $headerNames[$column] = $header
    }
    $required = @($SourceKeyColumn, $MappingIdColumn, $EntityColumn, $IndicationColumn, $DiseaseColumn, $StatusColumn, $RationaleColumn, $GenericColumn, $DosageFormColumn, $BrandColumn)
    foreach ($header in $required) { if (-not $headers.ContainsKey($header)) { throw "Source sheet does not contain required field: $header" } }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    function Get-CellText {
        param([int]$Row, [int]$Column)
        $value = $sourceValues[$Row, $Column]
        if ($null -eq $value) { return '' }
        $header = $headerNames[$Column]
        if ($value -is [double] -or $value -is [decimal] -or $value -is [int] -or $value -is [long]) {
            if ($header -match 'Date|date|DATE|日期|有效期') {
                try { return [datetime]::FromOADate([double]$value).ToString('yyyy-MM-dd') } catch { }
            }
            return [convert]::ToString($value, $culture)
        }
        return ([string]$value).Trim()
    }
    function Get-Field {
        param([int]$Row, [string]$Header)
        return Get-CellText -Row $Row -Column ([int]$headers[$Header])
    }

    $statusCounts = @{}
    $mappingIds = @{}
    $formulaCount = 0
    $externalFormulaCount = 0
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($column = 1; $column -le $columnCount; $column++) {
            $formula = [string]$sourceFormulas[$row, $column]
            if ($formula.StartsWith('=')) {
                $formulaCount++
                if ($formula.Contains('[')) { $externalFormulaCount++ }
            }
        }
    }
    for ($row = 2; $row -le $rowCount; $row++) {
        $mappingId = Get-Field -Row $row -Header $MappingIdColumn
        if (-not $mappingId -or $mappingIds.ContainsKey($mappingId)) { throw "Mapping ID is blank or duplicated: $mappingId" }
        $mappingIds[$mappingId] = $row
        $status = Get-Field -Row $row -Header $StatusColumn
        if (-not $statusCounts.ContainsKey($status)) { $statusCounts[$status] = 0 }
        $statusCounts[$status]++
        if ($status -eq $MappedValue -and -not (Get-Field -Row $row -Header $DiseaseColumn)) { throw "Mapped row has blank disease at source row $row." }
    }

    $layerOneByKey = @{}
    $layerOne = New-Object 'System.Collections.Generic.List[object]'
    for ($row = 2; $row -le $rowCount; $row++) {
        $sourceKey = Get-Field -Row $row -Header $SourceKeyColumn
        $entity = Get-Field -Row $row -Header $EntityColumn
        $disease = Get-Field -Row $row -Header $DiseaseColumn
        $status = Get-Field -Row $row -Header $StatusColumn
        $key = if ($status -eq $MappedValue -and $disease) { @($sourceKey, $entity, $disease) -join $GroupSeparator } else { "ROW$GroupSeparator$row" }
        if (-not $layerOneByKey.ContainsKey($key)) {
            $entry = [pscustomobject]@{ Anchor = $row; Rows = (New-Object 'System.Collections.Generic.List[int]'); SourceKey = $sourceKey; Entity = $entity; Disease = $disease; Status = $status; MappingIds = ''; Indications = ''; Rationales = '' }
            $layerOneByKey[$key] = $entry
            [void]$layerOne.Add($entry)
        }
        [void]$layerOneByKey[$key].Rows.Add($row)
    }
    foreach ($entry in $layerOne) {
        $entry.MappingIds = Join-Unique @($entry.Rows | ForEach-Object { Get-Field -Row $_ -Header $MappingIdColumn })
        $entry.Indications = Join-Unique @($entry.Rows | ForEach-Object { Get-Field -Row $_ -Header $IndicationColumn })
        $entry.Rationales = Join-Unique @($entry.Rows | ForEach-Object { Get-Field -Row $_ -Header $RationaleColumn })
    }

    $conflicts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($layerOne | Where-Object { $_.Rows.Count -gt 1 })) {
        foreach ($field in @($EntityColumn, $GenericColumn, $DosageFormColumn, $BrandColumn)) {
            $values = @(Get-UniqueArray @($entry.Rows | ForEach-Object { Get-Field -Row $_ -Header $field }))
            if ($values.Count -gt 1) {
                [void]$conflicts.Add([pscustomobject]@{ Layer = 1; Entity = $entry.Entity; Disease = $entry.Disease; MappingIds = $entry.MappingIds; Field = $field; Values = ($values -join ' | ') })
            }
        }
    }

    $layerTwoByKey = @{}
    $finalEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in $layerOne) {
        if ($entry.Status -eq $MappedValue -and $entry.Disease) {
            $key = @($entry.Entity, (Get-Field -Row $entry.Anchor -Header $GenericColumn), (Get-Field -Row $entry.Anchor -Header $DosageFormColumn), (Get-Field -Row $entry.Anchor -Header $BrandColumn), $entry.Indications, $entry.Disease, $entry.Status) -join $GroupSeparator
        } else {
            $key = "ROW$GroupSeparator$($entry.Anchor)"
        }
        if (-not $layerTwoByKey.ContainsKey($key)) {
            $final = [pscustomobject]@{ Anchor = $entry.Anchor; Children = (New-Object 'System.Collections.Generic.List[object]'); Status = $entry.Status; Disease = $entry.Disease }
            $layerTwoByKey[$key] = $final
            [void]$finalEntries.Add($final)
        }
        [void]$layerTwoByKey[$key].Children.Add($entry)
    }
    $finalEntries = @($finalEntries | Sort-Object Anchor)
    $layerOneGroups = @($layerOne | Where-Object { $_.Rows.Count -gt 1 })
    $layerTwoGroups = @($finalEntries | Where-Object { $_.Children.Count -gt 1 })
    $layerOneReduction = ($rowCount - 1) - $layerOne.Count
    $layerTwoReduction = 0
    foreach ($group in $layerTwoGroups) { $layerTwoReduction += ($group.Children.Count - 1) }

    $previewLayerOne = @($layerOneGroups | ForEach-Object {
        [pscustomobject]@{ Entity = $_.Entity; Disease = $_.Disease; SourceRows = ($_.Rows -join ','); MappingIds = $_.MappingIds; Indications = $_.Indications; Rationale = $_.Rationales }
    })
    $previewLayerTwo = @($layerTwoGroups | ForEach-Object {
        $differenceFields = New-Object 'System.Collections.Generic.List[string]'
        for ($column = 1; $column -le $columnCount; $column++) {
            $values = @(Get-UniqueArray @($_.Children | ForEach-Object { Get-CellText -Row $_.Anchor -Column $column }))
            if ($values.Count -gt 1) { [void]$differenceFields.Add($headerNames[$column]) }
        }
        $rationales = @(Get-UniqueArray @($_.Children | ForEach-Object { $_.Rationales }))
        [pscustomobject]@{ Entity = (Get-Field -Row $_.Anchor -Header $EntityColumn); Disease = $_.Disease; SourceKeys = (Join-Unique @($_.Children | ForEach-Object { $_.SourceKey })); MappingIds = (Join-Unique @($_.Children | ForEach-Object { $_.MappingIds })); Indications = (Join-Unique @($_.Children | ForEach-Object { $_.Indications })); RegistrationDifferenceFields = ($differenceFields -join ', '); ReviewDistinctRationale = ($rationales.Count -gt 1) }
    })

    $preview = [ordered]@{
        Mode = 'Preview'
        InputRows = $rowCount - 1
        InputColumns = $columnCount
        MappingStatusCounts = $statusCounts
        FormulaCells = $formulaCount
        PotentialExternalFormulaCells = $externalFormulaCount
        FirstLayerGroups = $previewLayerOne
        FirstLayerReduction = $layerOneReduction
        SecondLayerGroups = $previewLayerTwo
        SecondLayerReduction = $layerTwoReduction
        ExpectedOutputRows = ($rowCount - 1 - $layerOneReduction - $layerTwoReduction)
        Conflicts = @($conflicts | ForEach-Object { $_ })
    }
    if ($Mode -eq 'Preview') {
        $preview | ConvertTo-Json -Depth 8
        return
    }
    if ($conflicts.Count -gt 0) { throw 'Layer 1 core-field conflicts were found. Resolve them before generation.' }

    function Get-OutputValue {
        param([object]$Final, [int]$Column)
        $header = $headerNames[$Column]
        if ($header -eq $SourceKeyColumn) { return Join-Unique @($Final.Children | ForEach-Object { $_.SourceKey }) }
        if ($header -eq $MappingIdColumn) { return Join-Unique @($Final.Children | ForEach-Object { $_.MappingIds }) }
        if ($header -eq $IndicationColumn) { return Join-Unique @($Final.Children | ForEach-Object { $_.Indications }) }
        if ($header -eq $RationaleColumn) { return Join-Unique @($Final.Children | ForEach-Object { $_.Rationales }) }
        if ($Final.Children.Count -eq 1) { return Get-CellText -Row $Final.Children[0].Anchor -Column $Column }
        return Join-Unique @($Final.Children | ForEach-Object { Get-CellText -Row $_.Anchor -Column $Column })
    }

    [void]$sourceBook.SaveCopyAs($OutputPath)
    $createdOutput = $true
    $sourceBook.Close($false)
    $sourceBook = $null
    $outputBook = $excel.Workbooks.Open($OutputPath, 0, $false)
    $copiedSource = $outputBook.Worksheets.Item($SourceSheet)
    $copiedFormulas = $copiedSource.Range($copiedSource.Cells.Item(1, 1), $copiedSource.Cells.Item($rowCount, $columnCount)).Formula
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($column = 1; $column -le $columnCount; $column++) {
            if ([string]$sourceFormulas[$row, $column] -ne [string]$copiedFormulas[$row, $column]) { throw "Copied source sheet changed at R$row C$column." }
        }
    }
    for ($index = $outputBook.Worksheets.Count; $index -ge 1; $index--) {
        if ($outputBook.Worksheets.Item($index).Name -eq $FinalSheetName) { $outputBook.Worksheets.Item($index).Delete() }
    }
    $finalSheet = $outputBook.Worksheets.Add($outputBook.Worksheets.Item($outputBook.Worksheets.Count))
    $finalSheet.Name = $FinalSheetName
    for ($column = 1; $column -le $columnCount; $column++) {
        $finalSheet.Cells.Item(1, $column).NumberFormat = '@'
        $finalSheet.Cells.Item(1, $column).Value2 = [string]$sourceValues[1, $column]
        $finalSheet.Cells.Item(1, $column).Font.Bold = $true
        $finalSheet.Columns.Item($column).ColumnWidth = $copiedSource.Columns.Item($column).ColumnWidth
    }
    $outputRows = $finalEntries.Count
    $matrix = [Array]::CreateInstance([object], [int[]]@($outputRows, $columnCount), [int[]]@(1, 1))
    $finalStatusCounts = @{}
    for ($outputRow = 1; $outputRow -le $outputRows; $outputRow++) {
        $entry = $finalEntries[$outputRow - 1]
        if (-not $finalStatusCounts.ContainsKey($entry.Status)) { $finalStatusCounts[$entry.Status] = 0 }
        $finalStatusCounts[$entry.Status]++
        for ($column = 1; $column -le $columnCount; $column++) { $matrix.SetValue(([string](Get-OutputValue -Final $entry -Column $column)), $outputRow, $column) }
    }
    $outputRange = $finalSheet.Range($finalSheet.Cells.Item(2, 1), $finalSheet.Cells.Item($outputRows + 1, $columnCount))
    $outputRange.NumberFormat = '@'
    $outputRange.Value2 = $matrix
    $outputRange.WrapText = $true

    if ($outputRows -ne ($rowCount - 1 - $layerOneReduction - $layerTwoReduction)) { throw 'Output row-count verification failed.' }
    foreach ($status in $statusCounts.Keys) {
        if ($status -ne $MappedValue -and [int]$finalStatusCounts[$status] -ne [int]$statusCounts[$status]) { throw "Non-mapped status count changed: $status" }
    }
    if ([int]$finalStatusCounts[$MappedValue] -ne ([int]$statusCounts[$MappedValue] - $layerOneReduction - $layerTwoReduction)) { throw 'Mapped row-count verification failed.' }
    $finalFormulaCount = 0
    $finalFormulas = $finalSheet.Range($finalSheet.Cells.Item(1, 1), $finalSheet.Cells.Item($outputRows + 1, $columnCount)).Formula
    for ($row = 1; $row -le ($outputRows + 1); $row++) {
        for ($column = 1; $column -le $columnCount; $column++) { if (([string]$finalFormulas[$row, $column]).StartsWith('=')) { $finalFormulaCount++ } }
    }
    if ($finalFormulaCount -ne 0) { throw "Merged sheet contains $finalFormulaCount formula cells." }

    $outputBook.Save()
    $outputBook.Close($true)
    $outputBook = $null
    $excel.Quit()
    $excel = $null
    [pscustomobject]@{
        Mode = 'Generate'
        OutputPath = $OutputPath
        InputRows = $rowCount - 1
        OutputRows = $outputRows
        InputColumns = $columnCount
        FirstLayerReduction = $layerOneReduction
        SecondLayerReduction = $layerTwoReduction
        FirstLayerGroupCount = $layerOneGroups.Count
        SecondLayerGroupCount = $layerTwoGroups.Count
        FinalStatusCounts = $finalStatusCounts
        SourceSheetPreserved = $true
        FinalFormulaCells = $finalFormulaCount
    } | ConvertTo-Json -Depth 6
}
catch {
    if ($createdOutput -and (Test-Path -LiteralPath $OutputPath)) { Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue }
    throw
}
finally {
    if ($null -ne $outputBook) { try { $outputBook.Close($false) } catch { } }
    if ($null -ne $sourceBook) { try { $sourceBook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
