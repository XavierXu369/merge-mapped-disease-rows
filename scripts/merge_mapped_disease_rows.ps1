[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [ValidateSet('Preview', 'Generate')] [string]$Mode,
    [Parameter(Mandatory = $true)] [string]$ConfigPath,
    [string]$DecisionPath,
    [string]$OutputPath,
    [switch]$Confirmed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Version = '2.0.0'
$MappedStatus = 'Mapped'
$AllowedStatuses = @(
    'Mapped',
    'Manual Review Required - No TA Match',
    'Manual Review Required - Multiple Candidates',
    'Unmapped - Other TA',
    'Unmapped - Invalid Information'
)
$ExpectedMappingHeaders = @(
    '序号', 'Mapping ID', '药品', '适应症拆分结果', '标准化疾病实体',
    'ICD-10代码', 'ICD-10疾病名称', 'Disease CN', 'Mapping Status', 'Rationale'
)
$Culture = [System.Globalization.CultureInfo]::InvariantCulture

function Get-RequiredProperty {
    param([object]$Object, [string]$Name, [string]$Context)
    if ($null -eq $Object) { throw "$Context is missing." }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { throw "$Context is missing required property: $Name" }
    return $property.Value
}

function Read-JsonFile {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label was not found: $Path" }
    try { return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { throw "$Label is not valid JSON: $Path. $($_.Exception.Message)" }
}

function Resolve-RunPath {
    param([object]$Value, [string]$BaseDirectory, [string]$Label)
    $text = ([string]$Value).Trim()
    if (-not $text) { throw "$Label is blank." }
    $expanded = [Environment]::ExpandEnvironmentVariables($text)
    if (-not [IO.Path]::IsPathRooted($expanded)) { $expanded = Join-Path $BaseDirectory $expanded }
    return [IO.Path]::GetFullPath($expanded)
}

function Get-TextHash {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Convert-StableText {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Runtime.InteropServices.ErrorWrapper]) { return "ERROR:$($Value.ErrorCode)" }
    if ($Value -is [double]) { return $Value.ToString('R', $Culture) }
    if ($Value -is [single]) { return $Value.ToString('R', $Culture) }
    if ($Value -is [decimal]) { return $Value.ToString($Culture) }
    if ($Value -is [datetime]) { return $Value.ToString('o', $Culture) }
    return ([string]$Value).Trim()
}

function Normalize-Text {
    param([object]$Value)
    $text = Convert-StableText $Value
    $text = [regex]::Replace($text.Replace([char]0x00A0, ' '), '\s+', ' ').Trim()
    return $text.ToLowerInvariant()
}

function Get-MatrixValue {
    param([object]$Matrix, [int]$Rows, [int]$Columns, [int]$Row, [int]$Column)
    if ($Rows -eq 1 -and $Columns -eq 1) { return $Matrix }
    return $Matrix[$Row, $Column]
}

function Join-UniqueText {
    param([object[]]$Values, [string]$Delimiter)
    $seen = @{}
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        $text = (Convert-StableText $value).Trim()
        if (-not $text) { continue }
        $key = Normalize-Text $text
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            [void]$result.Add($text)
        }
    }
    return [string]::Join($Delimiter, $result.ToArray())
}

function Get-UniqueComparisonValues {
    param([object[]]$Values)
    $seen = @{}
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        $display = Convert-StableText $value
        $key = Normalize-Text $value
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            [void]$result.Add($(if ($display) { $display } else { '<blank>' }))
        }
    }
    return @($result.ToArray())
}

function Get-SheetFingerprint {
    param([object]$Worksheet)
    $used = $Worksheet.UsedRange
    $rows = [int]$used.Rows.Count
    $columns = [int]$used.Columns.Count
    $startRow = [int]$used.Row
    $startColumn = [int]$used.Column
    $values = $used.Value2
    $formulas = $used.Formula
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append("$startRow|$startColumn|$rows|$columns|")
    for ($row = 1; $row -le $rows; $row++) {
        for ($column = 1; $column -le $columns; $column++) {
            $value = Convert-StableText (Get-MatrixValue $values $rows $columns $row $column)
            $formula = Convert-StableText (Get-MatrixValue $formulas $rows $columns $row $column)
            [void]$builder.Append($value.Length).Append(':').Append($value).Append('|')
            [void]$builder.Append($formula.Length).Append(':').Append($formula).Append(';')
        }
    }
    return Get-TextHash $builder.ToString()
}

function Get-WorkbookFingerprints {
    param([object]$Workbook)
    $result = [ordered]@{}
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $sheet = $Workbook.Worksheets.Item($index)
        $result[[string]$sheet.Name] = Get-SheetFingerprint $sheet
    }
    return $result
}

function Get-FormulaAnalysis {
    param([object]$Values, [object]$Formulas, [int]$Rows, [int]$Columns)
    $formulaCount = 0
    $externalCount = 0
    $errorCount = 0
    $errorCells = New-Object 'System.Collections.Generic.List[string]'
    for ($row = 1; $row -le $Rows; $row++) {
        for ($column = 1; $column -le $Columns; $column++) {
            $formula = Convert-StableText (Get-MatrixValue $Formulas $Rows $Columns $row $column)
            if (-not $formula.StartsWith('=')) { continue }
            $formulaCount++
            if ($formula.Contains('[')) { $externalCount++ }
            $value = Get-MatrixValue $Values $Rows $Columns $row $column
            $valueText = Convert-StableText $value
            if ($value -is [System.Runtime.InteropServices.ErrorWrapper] -or $valueText -match '^#(NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!)') {
                $errorCount++
                if ($errorCells.Count -lt 20) { [void]$errorCells.Add("R$row C$column") }
            }
        }
    }
    return [pscustomobject]@{
        FormulaCells = $formulaCount
        ExternalFormulaCells = $externalCount
        FormulaErrorCells = $errorCount
        FormulaErrorLocations = @($errorCells.ToArray())
    }
}

function Compare-StringSets {
    param([string[]]$Left, [string[]]$Right)
    if ($Left.Count -ne $Right.Count) { return $false }
    for ($index = 0; $index -lt $Left.Count; $index++) {
        if ($Left[$index] -ne $Right[$index]) { return $false }
    }
    return $true
}

$ConfigPath = [IO.Path]::GetFullPath($ConfigPath)
$config = Read-JsonFile $ConfigPath 'Run config'
$configDirectory = [IO.Path]::GetDirectoryName($ConfigPath)
$inputConfig = Get-RequiredProperty $config 'input' 'Run config'
$fieldConfig = Get-RequiredProperty $config 'fields' 'Run config'
$mergeConfig = Get-RequiredProperty $config 'merge' 'Run config'
$formulaPolicy = Get-RequiredProperty $config 'formula_policy' 'Run config'
$outputConfig = Get-RequiredProperty $config 'output' 'Run config'
$workflow = Get-RequiredProperty $config 'workflow' 'Run config'

$InputPath = Resolve-RunPath (Get-RequiredProperty $inputConfig 'workbook' 'input') $configDirectory 'input.workbook'
if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input workbook was not found: $InputPath" }
$inputExtension = [IO.Path]::GetExtension($InputPath).ToLowerInvariant()
if ($inputExtension -notin @('.xlsx', '.xlsm')) { throw 'Input workbook must be .xlsx or .xlsm.' }
$sourceSheetName = ([string](Get-RequiredProperty $inputConfig 'source_sheet' 'input')).Trim()
$mappingSheetName = ([string](Get-RequiredProperty $inputConfig 'mapping_result_sheet' 'input')).Trim()
$headerRow = [int](Get-RequiredProperty $inputConfig 'header_row' 'input')
if ($headerRow -lt 1) { throw 'input.header_row must be at least 1.' }

$fieldNames = [ordered]@{
    SourceKey = ([string](Get-RequiredProperty $fieldConfig 'source_key' 'fields')).Trim()
    MappingId = ([string](Get-RequiredProperty $fieldConfig 'mapping_id' 'fields')).Trim()
    Entity = ([string](Get-RequiredProperty $fieldConfig 'entity' 'fields')).Trim()
    Indication = ([string](Get-RequiredProperty $fieldConfig 'indication' 'fields')).Trim()
    SourceDisease = ([string](Get-RequiredProperty $fieldConfig 'source_disease' 'fields')).Trim()
    Generic = ([string](Get-RequiredProperty $fieldConfig 'generic' 'fields')).Trim()
    DosageForm = ([string](Get-RequiredProperty $fieldConfig 'dosage_form' 'fields')).Trim()
    Brand = ([string](Get-RequiredProperty $fieldConfig 'brand' 'fields')).Trim()
}
foreach ($name in $fieldNames.Values) { if (-not $name) { throw 'Every fields value must be nonblank.' } }

$delimiter = ([string](Get-RequiredProperty $mergeConfig 'delimiter' 'merge')).Trim()
if (-not $delimiter) { throw 'merge.delimiter must be nonblank.' }
$layer1Allowed = @((Get-RequiredProperty $mergeConfig 'layer1_allowed_varying_fields' 'merge') | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
$layer2Identity = @((Get-RequiredProperty $mergeConfig 'layer2_identity_fields' 'merge') | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
$layer2Allowed = @((Get-RequiredProperty $mergeConfig 'layer2_allowed_difference_fields' 'merge') | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
$allowExternalFormulas = [bool](Get-RequiredProperty $formulaPolicy 'allow_external_formulas' 'formula_policy')
$finalSheetName = ([string](Get-RequiredProperty $outputConfig 'final_sheet' 'output')).Trim()
if (-not $finalSheetName -or $finalSheetName.Length -gt 31 -or $finalSheetName -match '[:\\/\?\*\[\]]') { throw 'output.final_sheet is blank or invalid for Excel.' }

$protectedLayer1 = @($fieldNames.SourceKey, $fieldNames.MappingId, $fieldNames.Entity, $fieldNames.Indication, $fieldNames.SourceDisease, $fieldNames.Generic, $fieldNames.DosageForm, $fieldNames.Brand)
foreach ($field in $layer1Allowed) {
    if ($protectedLayer1 -contains $field) { throw "Layer 1 protected field cannot be configured as varying: $field" }
}
$requiredLayer2Identity = @($fieldNames.Entity, $fieldNames.Generic, $fieldNames.DosageForm, $fieldNames.Brand, $fieldNames.Indication, $fieldNames.SourceDisease)
foreach ($field in $requiredLayer2Identity) {
    if ($layer2Identity -notcontains $field) { throw "merge.layer2_identity_fields must include: $field" }
}
foreach ($field in @($fieldNames.SourceKey, $fieldNames.MappingId)) {
    if ($layer2Identity -contains $field) { throw "Layer 2 identity fields may not include: $field" }
}
foreach ($field in $layer2Allowed) {
    if ($layer2Identity -contains $field) { throw "A Layer 2 field cannot be both identity and allowed difference: $field" }
}

$contractObject = [ordered]@{
    version = $Version
    input = $inputConfig
    fields = $fieldConfig
    merge = $mergeConfig
    formula_policy = $formulaPolicy
    output = $outputConfig
}
$contractHash = Get-TextHash ($contractObject | ConvertTo-Json -Depth 20 -Compress)
$inputHashBefore = Get-FileSha256 $InputPath

if ($Mode -eq 'Generate') {
    $requiredWorkflowFlags = @('input_approved', 'merge_keys_approved', 'second_layer_decisions_closed', 'final_execution_approved')
    $missingWorkflow = New-Object 'System.Collections.Generic.List[string]'
    foreach ($flag in $requiredWorkflowFlags) {
        $property = $workflow.PSObject.Properties[$flag]
        if ($null -eq $property -or $property.Value -ne $true) { [void]$missingWorkflow.Add($flag) }
    }
    if (-not $Confirmed -or $missingWorkflow.Count -gt 0) {
        throw "Generate requires -Confirmed and true workflow flags. Missing: $([string]::Join(', ', $missingWorkflow.ToArray()))"
    }
    if (-not $OutputPath) { throw 'Generate requires OutputPath.' }
    $OutputPath = [IO.Path]::GetFullPath($OutputPath)
    if ([IO.Path]::GetFullPath($InputPath) -eq $OutputPath) { throw 'OutputPath must differ from InputPath.' }
    if (Test-Path -LiteralPath $OutputPath) { throw "Output workbook already exists and will not be overwritten: $OutputPath" }
    if ([IO.Path]::GetExtension($OutputPath).ToLowerInvariant() -ne $inputExtension) { throw 'Output workbook extension must match the input extension.' }
    $outputDirectory = [IO.Path]::GetDirectoryName($OutputPath)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) { throw "Output directory does not exist: $outputDirectory" }
}

$excel = $null
$sourceBook = $null
$outputBook = $null
$reopenBook = $null
$createdOutput = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    try { $excel.AutomationSecurity = 3 } catch { }
    try { $excel.Calculation = -4135 } catch { }
    try { $excel.CalculateBeforeSave = $false } catch { }

    $sourceBook = $excel.Workbooks.Open($InputPath, 0, $true)
    try { $source = $sourceBook.Worksheets.Item($sourceSheetName) }
    catch { throw "Source sheet was not found: $sourceSheetName" }
    try { $mapping = $sourceBook.Worksheets.Item($mappingSheetName) }
    catch { throw "Mapping-result sheet was not found: $mappingSheetName" }
    for ($index = 1; $index -le $sourceBook.Worksheets.Count; $index++) {
        if ([string]$sourceBook.Worksheets.Item($index).Name -eq $finalSheetName) { throw "Final sheet already exists in the input workbook: $finalSheetName" }
    }

    $lastColumn = [int]$source.Cells.Item($headerRow, $source.Columns.Count).End(-4159).Column
    if ($lastColumn -lt 1) { throw 'Source header row is empty.' }
    $headers = @{}
    $headerNames = @{}
    for ($column = 1; $column -le $lastColumn; $column++) {
        $header = ([string]$source.Cells.Item($headerRow, $column).Value2).Trim()
        if (-not $header) { throw "Source header is blank at column $column." }
        if ($headers.ContainsKey($header)) { throw "Source header is duplicated: $header" }
        $headers[$header] = $column
        $headerNames[$column] = $header
    }
    foreach ($header in @($fieldNames.Values + $layer1Allowed + $layer2Identity + $layer2Allowed)) {
        if (-not $headers.ContainsKey($header)) { throw "Source sheet does not contain configured field: $header" }
    }

    $sourceLastRow = $headerRow
    foreach ($header in @($fieldNames.SourceKey, $fieldNames.MappingId, $fieldNames.Entity, $fieldNames.Indication, $fieldNames.SourceDisease)) {
        $candidateLast = [int]$source.Cells.Item($source.Rows.Count, [int]$headers[$header]).End(-4162).Row
        if ($candidateLast -gt $sourceLastRow) { $sourceLastRow = $candidateLast }
    }
    if ($sourceLastRow -le $headerRow) { throw 'Source sheet has no effective data rows.' }
    $sourceRowCount = $sourceLastRow - $headerRow + 1
    $sourceRange = $source.Range($source.Cells.Item($headerRow, 1), $source.Cells.Item($sourceLastRow, $lastColumn))
    $sourceValues = $sourceRange.Value2
    $sourceFormulas = $sourceRange.Formula

    $mappingLastColumn = [int]$mapping.Cells.Item(1, $mapping.Columns.Count).End(-4159).Column
    if ($mappingLastColumn -ne $ExpectedMappingHeaders.Count) { throw "Mapping-result sheet must contain exactly $($ExpectedMappingHeaders.Count) columns." }
    $mappingHeaders = @{}
    for ($column = 1; $column -le $mappingLastColumn; $column++) {
        $header = ([string]$mapping.Cells.Item(1, $column).Value2).Trim()
        if ($header -ne $ExpectedMappingHeaders[$column - 1]) { throw "Mapping-result header mismatch at column $column. Expected '$($ExpectedMappingHeaders[$column - 1])', found '$header'." }
        $mappingHeaders[$header] = $column
    }
    $mappingLastRow = 1
    foreach ($header in @('Mapping ID', 'Mapping Status', 'Rationale')) {
        $candidateLast = [int]$mapping.Cells.Item($mapping.Rows.Count, [int]$mappingHeaders[$header]).End(-4162).Row
        if ($candidateLast -gt $mappingLastRow) { $mappingLastRow = $candidateLast }
    }
    if ($mappingLastRow -le 1) { throw 'Mapping-result sheet has no data rows.' }
    $mappingRowCount = $mappingLastRow
    $mappingRange = $mapping.Range($mapping.Cells.Item(1, 1), $mapping.Cells.Item($mappingLastRow, $mappingLastColumn))
    $mappingValues = $mappingRange.Value2
    $mappingFormulas = $mappingRange.Formula

    function Get-SourceRaw {
        param([int]$ActualRow, [string]$Header)
        $matrixRow = $ActualRow - $headerRow + 1
        return Get-MatrixValue $sourceValues $sourceRowCount $lastColumn $matrixRow ([int]$headers[$Header])
    }
    function Get-SourceText {
        param([int]$ActualRow, [string]$Header)
        return Convert-StableText (Get-SourceRaw $ActualRow $Header)
    }
    function Get-MappingText {
        param([int]$ActualRow, [string]$Header)
        return Convert-StableText (Get-MatrixValue $mappingValues $mappingLastRow $mappingLastColumn $ActualRow ([int]$mappingHeaders[$Header]))
    }

    $mappingById = @{}
    $mappingOrder = New-Object 'System.Collections.Generic.List[string]'
    $statusCounts = [ordered]@{}
    foreach ($status in $AllowedStatuses) { $statusCounts[$status] = 0 }
    for ($row = 2; $row -le $mappingLastRow; $row++) {
        $mappingId = Get-MappingText $row 'Mapping ID'
        $status = Get-MappingText $row 'Mapping Status'
        $disease = Get-MappingText $row 'Disease CN'
        $rationale = Get-MappingText $row 'Rationale'
        if (-not $mappingId) { throw "Mapping-result ID is blank at row $row." }
        if ($mappingById.ContainsKey($mappingId)) { throw "Mapping-result ID is duplicated: $mappingId" }
        if ($AllowedStatuses -notcontains $status) { throw "Mapping ID $mappingId uses an invalid Mapping Status: $status" }
        if (-not $rationale) { throw "Rationale is blank for Mapping ID $mappingId." }
        if ($status -eq $MappedStatus -and -not $disease) { throw "Mapped result has blank Disease CN for Mapping ID $mappingId." }
        if ($status -ne $MappedStatus -and $disease) { throw "Non-Mapped result has Disease CN for Mapping ID $mappingId." }
        $mappingById[$mappingId] = [pscustomobject]@{ Disease = $disease; Status = $status; Rationale = $rationale; Row = $row }
        [void]$mappingOrder.Add($mappingId)
        $statusCounts[$status]++
    }

    $sourceRows = New-Object 'System.Collections.Generic.List[object]'
    $sourceIds = @{}
    $sourceOrder = New-Object 'System.Collections.Generic.List[string]'
    for ($row = $headerRow + 1; $row -le $sourceLastRow; $row++) {
        $mappingId = Get-SourceText $row $fieldNames.MappingId
        $sourceKey = Get-SourceText $row $fieldNames.SourceKey
        $entity = Get-SourceText $row $fieldNames.Entity
        $indication = Get-SourceText $row $fieldNames.Indication
        $backfill = Get-SourceText $row $fieldNames.SourceDisease
        if (-not $mappingId -and -not $sourceKey -and -not $entity -and -not $indication -and -not $backfill) { continue }
        if (-not $mappingId) { throw "Source Mapping ID is blank at row $row." }
        if ($sourceIds.ContainsKey($mappingId)) { throw "Source Mapping ID is duplicated: $mappingId" }
        if (-not $mappingById.ContainsKey($mappingId)) { throw "Source Mapping ID is absent from Mapping results: $mappingId" }
        $mapped = $mappingById[$mappingId]
        $expectedBackfill = if ($mapped.Status -eq $MappedStatus) { $mapped.Disease } else { '' }
        if ((Normalize-Text $backfill) -ne (Normalize-Text $expectedBackfill)) { throw "Source disease backfill disagrees with Mapping results for Mapping ID $mappingId." }
        if ($mapped.Status -eq $MappedStatus) {
            if (-not $sourceKey) { throw "Mapped row has blank source key for Mapping ID $mappingId." }
            if (-not $entity) { throw "Mapped row has blank Entity for Mapping ID $mappingId." }
            if (-not $indication) { throw "Mapped row has blank split indication for Mapping ID $mappingId." }
        }
        $record = [pscustomobject]@{
            Row = $row
            SourceKey = $sourceKey
            MappingId = $mappingId
            Entity = $entity
            Indication = $indication
            Disease = $mapped.Disease
            Status = $mapped.Status
            Rationale = $mapped.Rationale
        }
        [void]$sourceRows.Add($record)
        $sourceIds[$mappingId] = $record
        [void]$sourceOrder.Add($mappingId)
    }
    if ($sourceRows.Count -ne $mappingById.Count) { throw 'Source and Mapping-result effective row counts differ.' }
    if (-not (Compare-StringSets $sourceOrder.ToArray() $mappingOrder.ToArray())) { throw 'Source and Mapping-result Mapping ID order differs.' }

    $sourceFormulaAnalysis = Get-FormulaAnalysis $sourceValues $sourceFormulas $sourceRowCount $lastColumn
    $mappingFormulaAnalysis = Get-FormulaAnalysis $mappingValues $mappingFormulas $mappingLastRow $mappingLastColumn
    $workbookExternalLinks = 0
    try {
        $links = @($sourceBook.LinkSources(1))
        if ($links.Count -eq 1 -and $null -eq $links[0]) { $workbookExternalLinks = 0 } else { $workbookExternalLinks = $links.Count }
    }
    catch { $workbookExternalLinks = 0 }

    $blockingIssues = New-Object 'System.Collections.Generic.List[string]'
    if ($sourceFormulaAnalysis.FormulaErrorCells -gt 0) { [void]$blockingIssues.Add("Source formulas contain $($sourceFormulaAnalysis.FormulaErrorCells) error cells.") }
    if ($mappingFormulaAnalysis.FormulaErrorCells -gt 0) { [void]$blockingIssues.Add("Mapping formulas contain $($mappingFormulaAnalysis.FormulaErrorCells) error cells.") }
    if (-not $allowExternalFormulas -and ($sourceFormulaAnalysis.ExternalFormulaCells -gt 0 -or $mappingFormulaAnalysis.ExternalFormulaCells -gt 0 -or $workbookExternalLinks -gt 0)) {
        [void]$blockingIssues.Add('External formulas or workbook links are present but formula_policy.allow_external_formulas is false.')
    }

    $layerOneByKey = @{}
    $layerOneEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($record in $sourceRows) {
        $key = if ($record.Status -eq $MappedStatus) {
            "M|$(Normalize-Text $record.SourceKey)|$(Normalize-Text $record.Disease)"
        }
        else { "R|$($record.Row)" }
        if (-not $layerOneByKey.ContainsKey($key)) {
            $entry = [pscustomobject]@{
                Anchor = $record.Row
                Rows = (New-Object 'System.Collections.Generic.List[object]')
                SourceKey = $record.SourceKey
                Disease = $record.Disease
                Status = $record.Status
                MappingIds = ''
                Indications = ''
                GroupId = ''
                AllowedVaryingFields = @()
                ConflictingFields = @()
            }
            $layerOneByKey[$key] = $entry
            [void]$layerOneEntries.Add($entry)
        }
        [void]$layerOneByKey[$key].Rows.Add($record)
    }

    $layerOneGroups = New-Object 'System.Collections.Generic.List[object]'
    $layerOneReduction = 0
    foreach ($entry in $layerOneEntries) {
        $entry.MappingIds = Join-UniqueText @($entry.Rows | ForEach-Object { $_.MappingId }) $delimiter
        $entry.Indications = Join-UniqueText @($entry.Rows | ForEach-Object { $_.Indication }) $delimiter
        $entry.GroupId = 'L1-' + (Get-TextHash "$($entry.SourceKey)|$($entry.Disease)|$($entry.MappingIds)").Substring(0, 16)
        if ($entry.Rows.Count -le 1) { continue }
        $layerOneReduction += $entry.Rows.Count - 1
        $allowedDifferences = New-Object 'System.Collections.Generic.List[object]'
        $conflicts = New-Object 'System.Collections.Generic.List[object]'
        for ($column = 1; $column -le $lastColumn; $column++) {
            $header = [string]$headerNames[$column]
            if ($header -in @($fieldNames.MappingId, $fieldNames.Indication)) { continue }
            $values = @(Get-UniqueComparisonValues @($entry.Rows | ForEach-Object { Get-SourceRaw $_.Row $header }))
            if ($values.Count -le 1) { continue }
            $detail = [pscustomobject]@{ Field = $header; Values = $values }
            if ($layer1Allowed -contains $header) { [void]$allowedDifferences.Add($detail) }
            else { [void]$conflicts.Add($detail) }
        }
        $entry.AllowedVaryingFields = @($allowedDifferences.ToArray())
        $entry.ConflictingFields = @($conflicts.ToArray())
        if ($conflicts.Count -gt 0) { [void]$blockingIssues.Add("Layer 1 group $($entry.GroupId) contains unapproved field differences.") }
        [void]$layerOneGroups.Add([pscustomobject]@{
            GroupId = $entry.GroupId
            SourceKey = $entry.SourceKey
            Disease = $entry.Disease
            SourceRows = @($entry.Rows | ForEach-Object { $_.Row })
            MappingIds = $entry.MappingIds
            Indications = $entry.Indications
            AllowedVaryingFields = $entry.AllowedVaryingFields
            ConflictingFields = $entry.ConflictingFields
        })
    }

    function Get-EntryFieldText {
        param([object]$Entry, [string]$Header)
        if ($Header -eq $fieldNames.SourceKey) { return $Entry.SourceKey }
        if ($Header -eq $fieldNames.MappingId) { return $Entry.MappingIds }
        if ($Header -eq $fieldNames.Indication) { return $Entry.Indications }
        if ($Header -eq $fieldNames.SourceDisease) { return $Entry.Disease }
        if ($layer1Allowed -contains $Header -and $Entry.Rows.Count -gt 1) {
            return Join-UniqueText @($Entry.Rows | ForEach-Object { Get-SourceRaw $_.Row $Header }) $delimiter
        }
        return Get-SourceText $Entry.Anchor $Header
    }

    $layerTwoByKey = @{}
    foreach ($entry in $layerOneEntries) {
        if ($entry.Status -ne $MappedStatus -or -not $entry.Disease) { continue }
        $identityNormalized = @($layer2Identity | ForEach-Object { Normalize-Text (Get-EntryFieldText $entry $_) })
        $key = [string]::Join([char]31, $identityNormalized)
        if (-not $layerTwoByKey.ContainsKey($key)) { $layerTwoByKey[$key] = New-Object 'System.Collections.Generic.List[object]' }
        [void]$layerTwoByKey[$key].Add($entry)
    }

    $layerTwoCandidates = New-Object 'System.Collections.Generic.List[object]'
    foreach ($key in $layerTwoByKey.Keys) {
        $children = $layerTwoByKey[$key]
        if ($children.Count -le 1) { continue }
        $mappingIds = Join-UniqueText @($children | ForEach-Object { $_.MappingIds }) $delimiter
        $candidateId = 'L2-' + (Get-TextHash "$key|$mappingIds").Substring(0, 16)
        $differenceDetails = New-Object 'System.Collections.Generic.List[object]'
        $prohibited = New-Object 'System.Collections.Generic.List[string]'
        for ($column = 1; $column -le $lastColumn; $column++) {
            $header = [string]$headerNames[$column]
            $values = @(Get-UniqueComparisonValues @($children | ForEach-Object { Get-EntryFieldText $_ $header }))
            if ($values.Count -le 1) { continue }
            [void]$differenceDetails.Add([pscustomobject]@{ Field = $header; Values = $values })
            if ($header -notin @($fieldNames.SourceKey, $fieldNames.MappingId) -and $layer2Allowed -notcontains $header) {
                [void]$prohibited.Add($header)
            }
        }
        $identityValues = [ordered]@{}
        foreach ($field in $layer2Identity) { $identityValues[$field] = Get-EntryFieldText $children[0] $field }
        [void]$layerTwoCandidates.Add([pscustomobject]@{
            CandidateId = $candidateId
            Anchor = ($children | Measure-Object -Property Anchor -Minimum).Minimum
            Children = $children
            SourceKeys = (Join-UniqueText @($children | ForEach-Object { $_.SourceKey }) $delimiter)
            MappingIds = $mappingIds
            IdentityValues = $identityValues
            DifferenceFields = @($differenceDetails.ToArray())
            ProhibitedDifferenceFields = @($prohibited.ToArray())
            MergeEligible = ($prohibited.Count -eq 0)
            PotentialReduction = $children.Count - 1
        })
    }
    $layerTwoCandidates = @($layerTwoCandidates | Sort-Object Anchor, CandidateId)

    $planObject = [ordered]@{
        version = $Version
        input_sha256 = $inputHashBefore
        contract_hash = $contractHash
        layer1 = @($layerOneGroups | ForEach-Object {
            [ordered]@{ id = $_.GroupId; mapping_ids = $_.MappingIds; conflicts = @($_.ConflictingFields | ForEach-Object { $_.Field }) }
        })
        layer2 = @($layerTwoCandidates | ForEach-Object {
            [ordered]@{ id = $_.CandidateId; mapping_ids = $_.MappingIds; differences = @($_.DifferenceFields | ForEach-Object { $_.Field }); prohibited = $_.ProhibitedDifferenceFields }
        })
    }
    $runFingerprint = Get-TextHash ($planObject | ConvertTo-Json -Depth 20 -Compress)
    $eligibleLayerTwoReduction = 0
    foreach ($candidate in $layerTwoCandidates) { if ($candidate.MergeEligible) { $eligibleLayerTwoReduction += $candidate.PotentialReduction } }

    $previewCandidates = @($layerTwoCandidates | ForEach-Object {
        [pscustomobject]@{
            CandidateId = $_.CandidateId
            SourceKeys = $_.SourceKeys
            MappingIds = $_.MappingIds
            IdentityValues = $_.IdentityValues
            DifferenceFields = $_.DifferenceFields
            ProhibitedDifferenceFields = $_.ProhibitedDifferenceFields
            MergeEligible = $_.MergeEligible
            PotentialReduction = $_.PotentialReduction
        }
    })
    $preview = [ordered]@{
        Status = $(if ($blockingIssues.Count -gt 0) { 'blocked' } elseif ($layerTwoCandidates.Count -gt 0) { 'ready_for_layer2_decisions' } else { 'ready_for_generation_confirmation' })
        SkillVersion = $Version
        Mode = 'Preview'
        InputPath = $InputPath
        InputSha256 = $inputHashBefore
        ContractHash = $contractHash
        RunFingerprint = $runFingerprint
        SourceSheet = $sourceSheetName
        SourceRows = $sourceRows.Count
        SourceColumns = $lastColumn
        MappingResultSheet = $mappingSheetName
        MappingResultRows = $mappingById.Count
        MappingJoinValidated = $true
        MappingStatusCounts = $statusCounts
        SourceFormulaAnalysis = $sourceFormulaAnalysis
        MappingFormulaAnalysis = $mappingFormulaAnalysis
        WorkbookExternalLinks = $workbookExternalLinks
        Layer1Groups = @($layerOneGroups.ToArray())
        Layer1Reduction = $layerOneReduction
        RowsAfterLayer1 = $sourceRows.Count - $layerOneReduction
        Layer2Candidates = $previewCandidates
        Layer2CandidateCount = $layerTwoCandidates.Count
        EligibleLayer2PotentialReduction = $eligibleLayerTwoReduction
        MinimumRowsIfAllEligibleLayer2CandidatesMerge = $sourceRows.Count - $layerOneReduction - $eligibleLayerTwoReduction
        DecisionFileRequired = ($layerTwoCandidates.Count -gt 0)
        FinalSheetName = $finalSheetName
        FinalSheetAvailable = $true
        BlockingIssues = @($blockingIssues.ToArray())
    }
    if ($Mode -eq 'Preview') {
        $preview | ConvertTo-Json -Depth 20
        return
    }

    if ($blockingIssues.Count -gt 0) { throw "Generation is blocked: $([string]::Join(' | ', $blockingIssues.ToArray()))" }

    $decisionById = @{}
    if ($layerTwoCandidates.Count -gt 0) {
        if (-not $DecisionPath) { throw 'Layer 2 candidates exist; Generate requires DecisionPath.' }
        $DecisionPath = [IO.Path]::GetFullPath($DecisionPath)
        $decisionData = Read-JsonFile $DecisionPath 'Layer 2 decision file'
        $decisionFingerprint = ([string](Get-RequiredProperty $decisionData 'run_fingerprint' 'Decision file')).Trim()
        if ($decisionFingerprint -ne $runFingerprint) { throw 'Decision run_fingerprint does not match the current Preview.' }
        $decisions = @((Get-RequiredProperty $decisionData 'layer2_decisions' 'Decision file'))
        foreach ($decision in $decisions) {
            $candidateId = ([string](Get-RequiredProperty $decision 'candidate_id' 'Layer 2 decision')).Trim()
            $action = ([string](Get-RequiredProperty $decision 'decision' 'Layer 2 decision')).Trim().ToUpperInvariant()
            $rationale = ([string](Get-RequiredProperty $decision 'rationale' 'Layer 2 decision')).Trim()
            if (-not $candidateId -or $decisionById.ContainsKey($candidateId)) { throw "Layer 2 decision candidate is blank or duplicated: $candidateId" }
            if ($action -notin @('MERGE', 'KEEP SEPARATE')) { throw "Invalid Layer 2 decision for ${candidateId}: $action" }
            if (-not $rationale) { throw "Layer 2 decision rationale is blank for $candidateId." }
            $decisionById[$candidateId] = [pscustomobject]@{ Decision = $action; Rationale = $rationale }
        }
        foreach ($candidate in $layerTwoCandidates) {
            if (-not $decisionById.ContainsKey($candidate.CandidateId)) { throw "Layer 2 candidate is not closed: $($candidate.CandidateId)" }
            if ($decisionById[$candidate.CandidateId].Decision -eq 'MERGE' -and -not $candidate.MergeEligible) {
                throw "Layer 2 candidate $($candidate.CandidateId) cannot be merged because it has prohibited differences: $([string]::Join(', ', $candidate.ProhibitedDifferenceFields))"
            }
        }
        foreach ($candidateId in $decisionById.Keys) {
            if (-not ($layerTwoCandidates.CandidateId -contains $candidateId)) { throw "Decision file contains unknown Layer 2 candidate: $candidateId" }
        }
    }
    elseif ($DecisionPath) {
        $DecisionPath = [IO.Path]::GetFullPath($DecisionPath)
        $decisionData = Read-JsonFile $DecisionPath 'Layer 2 decision file'
        $decisions = @((Get-RequiredProperty $decisionData 'layer2_decisions' 'Decision file'))
        if ($decisions.Count -gt 0) { throw 'Decision file contains Layer 2 decisions, but Preview found no candidates.' }
    }

    $candidateByEntryGroupId = @{}
    foreach ($candidate in $layerTwoCandidates) {
        foreach ($child in $candidate.Children) { $candidateByEntryGroupId[$child.GroupId] = $candidate }
    }
    $handledCandidates = @{}
    $finalEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($layerOneEntries | Sort-Object Anchor)) {
        if (-not $candidateByEntryGroupId.ContainsKey($entry.GroupId)) {
            [void]$finalEntries.Add([pscustomobject]@{ Anchor = $entry.Anchor; Children = @($entry); Status = $entry.Status; Disease = $entry.Disease; Layer2Decision = '' })
            continue
        }
        $candidate = $candidateByEntryGroupId[$entry.GroupId]
        if ($handledCandidates.ContainsKey($candidate.CandidateId)) { continue }
        $handledCandidates[$candidate.CandidateId] = $true
        $decision = $decisionById[$candidate.CandidateId].Decision
        if ($decision -eq 'MERGE') {
            $candidateChildren = @($candidate.Children | ForEach-Object { $_ })
            [void]$finalEntries.Add([pscustomobject]@{ Anchor = $candidate.Anchor; Children = $candidateChildren; Status = $MappedStatus; Disease = $candidateChildren[0].Disease; Layer2Decision = 'MERGE' })
        }
        else {
            foreach ($child in @($candidate.Children | Sort-Object Anchor)) {
                [void]$finalEntries.Add([pscustomobject]@{ Anchor = $child.Anchor; Children = @($child); Status = $child.Status; Disease = $child.Disease; Layer2Decision = 'KEEP SEPARATE' })
            }
        }
    }
    $finalEntries = @($finalEntries | Sort-Object Anchor)

    function Get-LayerOneOutputValue {
        param([object]$Entry, [string]$Header)
        if ($Entry.Rows.Count -eq 1) { return Get-SourceRaw $Entry.Anchor $Header }
        if ($Header -eq $fieldNames.MappingId) { return $Entry.MappingIds }
        if ($Header -eq $fieldNames.Indication) { return $Entry.Indications }
        if ($Header -eq $fieldNames.SourceDisease) { return $Entry.Disease }
        if ($layer1Allowed -contains $Header) { return Join-UniqueText @($Entry.Rows | ForEach-Object { Get-SourceRaw $_.Row $Header }) $delimiter }
        return Get-SourceRaw $Entry.Anchor $Header
    }
    function Get-FinalOutputValue {
        param([object]$Final, [string]$Header)
        if ($Final.Children.Count -eq 1) { return Get-LayerOneOutputValue $Final.Children[0] $Header }
        if ($Header -eq $fieldNames.SourceKey) { return Join-UniqueText @($Final.Children | ForEach-Object { $_.SourceKey }) $delimiter }
        if ($Header -eq $fieldNames.MappingId) { return Join-UniqueText @($Final.Children | ForEach-Object { $_.MappingIds }) $delimiter }
        if ($Header -eq $fieldNames.Indication) { return Join-UniqueText @($Final.Children | ForEach-Object { $_.Indications }) $delimiter }
        if ($Header -eq $fieldNames.SourceDisease) { return $Final.Disease }
        if ($layer2Allowed -contains $Header) { return Join-UniqueText @($Final.Children | ForEach-Object { Get-EntryFieldText $_ $Header }) $delimiter }
        return Get-LayerOneOutputValue $Final.Children[0] $Header
    }

    $originalSheetNames = @()
    for ($index = 1; $index -le $sourceBook.Worksheets.Count; $index++) { $originalSheetNames += [string]$sourceBook.Worksheets.Item($index).Name }
    $originalFingerprints = Get-WorkbookFingerprints $sourceBook

    [void]$sourceBook.SaveCopyAs($OutputPath)
    $createdOutput = $true
    $sourceBook.Close($false)
    $sourceBook = $null
    $outputBook = $excel.Workbooks.Open($OutputPath, 0, $false)
    for ($index = 1; $index -le $outputBook.Worksheets.Count; $index++) {
        if ([string]$outputBook.Worksheets.Item($index).Name -eq $finalSheetName) { throw "Final sheet unexpectedly exists in the output copy: $finalSheetName" }
    }
    $finalSheet = $outputBook.Worksheets.Add($null, $outputBook.Worksheets.Item($outputBook.Worksheets.Count))
    $finalSheet.Name = $finalSheetName

    $sourceOut = $outputBook.Worksheets.Item($sourceSheetName)
    [void]$sourceOut.Range($sourceOut.Cells.Item($headerRow, 1), $sourceOut.Cells.Item($headerRow, $lastColumn)).Copy($finalSheet.Range($finalSheet.Cells.Item(1, 1), $finalSheet.Cells.Item(1, $lastColumn)))
    for ($column = 1; $column -le $lastColumn; $column++) { $finalSheet.Columns.Item($column).ColumnWidth = $sourceOut.Columns.Item($column).ColumnWidth }

    $outputRows = $finalEntries.Count
    $matrix = [Array]::CreateInstance([object], [int[]]@($outputRows, $lastColumn), [int[]]@(1, 1))
    for ($outputRow = 1; $outputRow -le $outputRows; $outputRow++) {
        $final = $finalEntries[$outputRow - 1]
        [void]$sourceOut.Range($sourceOut.Cells.Item($final.Anchor, 1), $sourceOut.Cells.Item($final.Anchor, $lastColumn)).Copy($finalSheet.Range($finalSheet.Cells.Item($outputRow + 1, 1), $finalSheet.Cells.Item($outputRow + 1, $lastColumn)))
        for ($column = 1; $column -le $lastColumn; $column++) {
            $header = [string]$headerNames[$column]
            $matrix.SetValue((Get-FinalOutputValue $final $header), $outputRow, $column)
        }
    }
    $outputRange = $finalSheet.Range($finalSheet.Cells.Item(2, 1), $finalSheet.Cells.Item($outputRows + 1, $lastColumn))
    $outputRange.Value2 = $matrix
    $outputRange.WrapText = $true
    $finalSheet.Range($finalSheet.Cells.Item(1, 1), $finalSheet.Cells.Item($outputRows + 1, $lastColumn)).AutoFilter() | Out-Null

    $approvedLayerTwoReduction = 0
    $mergedCandidateIds = New-Object 'System.Collections.Generic.List[string]'
    $keptCandidateIds = New-Object 'System.Collections.Generic.List[string]'
    foreach ($candidate in $layerTwoCandidates) {
        if ($decisionById[$candidate.CandidateId].Decision -eq 'MERGE') {
            $approvedLayerTwoReduction += $candidate.PotentialReduction
            [void]$mergedCandidateIds.Add($candidate.CandidateId)
        }
        else { [void]$keptCandidateIds.Add($candidate.CandidateId) }
    }
    $expectedOutputRows = $sourceRows.Count - $layerOneReduction - $approvedLayerTwoReduction
    if ($outputRows -ne $expectedOutputRows) { throw "Output row arithmetic failed: expected $expectedOutputRows, found $outputRows." }

    $finalStatusCounts = [ordered]@{}
    foreach ($status in $AllowedStatuses) { $finalStatusCounts[$status] = 0 }
    foreach ($final in $finalEntries) { $finalStatusCounts[$final.Status]++ }
    foreach ($status in $AllowedStatuses) {
        $expected = if ($status -eq $MappedStatus) { [int]$statusCounts[$status] - $layerOneReduction - $approvedLayerTwoReduction } else { [int]$statusCounts[$status] }
        if ([int]$finalStatusCounts[$status] -ne $expected) { throw "Final status count changed unexpectedly for $status." }
    }

    $finalFormulaValues = $finalSheet.Range($finalSheet.Cells.Item(1, 1), $finalSheet.Cells.Item($outputRows + 1, $lastColumn)).Formula
    $finalFormulaCount = 0
    for ($row = 1; $row -le $outputRows + 1; $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) {
            $formula = Convert-StableText (Get-MatrixValue $finalFormulaValues ($outputRows + 1) $lastColumn $row $column)
            if ($formula.StartsWith('=')) { $finalFormulaCount++ }
        }
    }
    if ($finalFormulaCount -ne 0) { throw "Final sheet contains $finalFormulaCount formula cells." }

    $outputBook.Save()
    $outputBook.Close($true)
    $outputBook = $null

    $reopenBook = $excel.Workbooks.Open($OutputPath, 0, $true)
    $reopenedNames = @()
    for ($index = 1; $index -le $reopenBook.Worksheets.Count; $index++) { $reopenedNames += [string]$reopenBook.Worksheets.Item($index).Name }
    $expectedNames = @($originalSheetNames + $finalSheetName)
    if (-not (Compare-StringSets $reopenedNames $expectedNames)) { throw 'Original worksheet names or order changed in the output.' }
    foreach ($sheetName in $originalSheetNames) {
        $actualFingerprint = Get-SheetFingerprint $reopenBook.Worksheets.Item($sheetName)
        if ($actualFingerprint -ne $originalFingerprints[$sheetName]) { throw "Preserved worksheet changed: $sheetName" }
    }
    $reopenedFinal = $reopenBook.Worksheets.Item($finalSheetName)
    $reopenedLastColumn = [int]$reopenedFinal.Cells.Item(1, $reopenedFinal.Columns.Count).End(-4159).Column
    $reopenedLastRow = [int]$reopenedFinal.Cells.Item($reopenedFinal.Rows.Count, [int]$headers[$fieldNames.MappingId]).End(-4162).Row
    if ($reopenedLastColumn -ne $lastColumn -or $reopenedLastRow -ne $outputRows + 1) { throw 'Reopened final-sheet dimensions differ from the expected output.' }
    $reopenedFormulas = $reopenedFinal.Range($reopenedFinal.Cells.Item(1, 1), $reopenedFinal.Cells.Item($reopenedLastRow, $reopenedLastColumn)).Formula
    for ($row = 1; $row -le $reopenedLastRow; $row++) {
        for ($column = 1; $column -le $reopenedLastColumn; $column++) {
            if ((Convert-StableText (Get-MatrixValue $reopenedFormulas $reopenedLastRow $reopenedLastColumn $row $column)).StartsWith('=')) { throw 'Reopened final sheet contains a formula.' }
        }
    }
    $reopenBook.Close($false)
    $reopenBook = $null

    if ((Get-FileSha256 $InputPath) -ne $inputHashBefore) { throw 'Input workbook changed during generation.' }
    [ordered]@{
        Status = 'generated'
        SkillVersion = $Version
        OutputPath = $OutputPath
        OutputSha256 = Get-FileSha256 $OutputPath
        InputPath = $InputPath
        InputSha256Unchanged = $true
        ContractHash = $contractHash
        RunFingerprint = $runFingerprint
        InputRows = $sourceRows.Count
        OutputRows = $outputRows
        OutputColumns = $lastColumn
        Layer1GroupCount = $layerOneGroups.Count
        Layer1Reduction = $layerOneReduction
        Layer2MergedCandidateIds = @($mergedCandidateIds.ToArray())
        Layer2KeptSeparateCandidateIds = @($keptCandidateIds.ToArray())
        Layer2Reduction = $approvedLayerTwoReduction
        MappingStatusCountsBefore = $statusCounts
        MappingStatusCountsAfter = $finalStatusCounts
        SourceFormulaAnalysis = $sourceFormulaAnalysis
        MappingFormulaAnalysis = $mappingFormulaAnalysis
        OriginalSheetsPreserved = $true
        FinalFormulaCells = 0
        ReopenQc = 'passed'
    } | ConvertTo-Json -Depth 20
}
catch {
    if ($createdOutput -and $OutputPath -and (Test-Path -LiteralPath $OutputPath)) {
        try { [IO.File]::Delete($OutputPath) } catch { }
    }
    throw
}
finally {
    if ($null -ne $reopenBook) { try { $reopenBook.Close($false) } catch { } }
    if ($null -ne $outputBook) { try { $outputBook.Close($false) } catch { } }
    if ($null -ne $sourceBook) { try { $sourceBook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
