#Requires -Version 7.2
<#
.SYNOPSIS
    Runs Provision-Lists.ps1 against the in-memory MockPnP module and checks that it
    builds everything in the schema, is idempotent, and repairs drift.
    Usage: pwsh provision/tests/Test-Provisioning.ps1
#>
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$script = Join-Path $here '..' 'Provision-Lists.ps1'
$schema = Get-Content -Raw (Join-Path $here '..' 'list-schema.json') -Encoding utf8 | ConvertFrom-Json -AsHashtable
Import-Module (Join-Path $here 'MockPnP.psm1') -Force

$failures = 0
function Assert([bool]$Condition, [string]$Message) {
    if ($Condition) { Write-Host "  PASS $Message" -ForegroundColor Green }
    else { Write-Host "  FAIL $Message" -ForegroundColor Red; $script:failures++ }
}
function Invoke-Provision([switch]$WhatIf) {
    $out = & $script -SiteUrl 'https://contoso.sharepoint.com/sites/test' -ClientId 'x' -SkipConnect -SetRegionalSettings -WhatIf:$WhatIf 6>&1 3>&1
    return ($out | Out-String)
}

Write-Host 'Test 1: -WhatIf on an empty site changes nothing'
Reset-MockPnP
$o = Invoke-Provision -WhatIf
Assert ((Get-MockPnPState).Lists.Count -eq 0) 'no lists created'
Assert ($o -match 'WhatIf mode') 'reports WhatIf mode'

Write-Host 'Test 2: first run builds every list, column, index and view'
Reset-MockPnP
$o = Invoke-Provision
$state = Get-MockPnPState
Assert ($state.Lists.Count -eq $schema.lists.Count) "$($schema.lists.Count) lists created"
Assert ($state.Regional.LocaleId -eq 3081 -and $state.Regional.TimeZoneId -eq 18 -and $state.Regional.FirstDayOfWeek -eq 1) 'regional settings en-AU / Brisbane / Monday'
foreach ($def in $schema.lists) {
    $l = $state.Lists[$def.url]
    Assert ($l.EnableVersioning -and $l.MajorVersionLimit -eq 500) "$($def.title): versioning on, 500 versions"
    Assert ($l.EnableAttachments -eq [bool]$def.attachments) "$($def.title): attachments = $([bool]$def.attachments)"
    Assert ($l.Fields['Title'].Title -eq $def.titleField.displayName) "$($def.title): Title renamed to '$($def.titleField.displayName)'"
    $missing = @($def.fields | Where-Object { -not $l.Fields.Contains($_.name) })
    Assert ($missing.Count -eq 0) "$($def.title): all $($def.fields.Count) columns exist"
    $wantIdx = @($def.fields | Where-Object { $_['indexed'] -or $_.type -eq 'Lookup' }).Count + $(if ($def.titleField.indexed) { 1 } else { 0 })
    $haveIdx = @($l.Fields.Values | Where-Object Indexed).Count
    Assert ($haveIdx -eq $wantIdx) "$($def.title): $wantIdx indexed columns (found $haveIdx)"
    foreach ($f in $def.fields | Where-Object { $_['unique'] }) { Assert $l.Fields[$f.name].EnforceUniqueValues "$($def.title).$($f.name) unique" }
    if ($def.titleField.unique) { Assert $l.Fields['Title'].EnforceUniqueValues "$($def.title).Title unique" }
    foreach ($v in $def.views) {
        $mv = $l.Views | Where-Object Title -eq $v.title
        Assert ($null -ne $mv -and (($mv.ViewFields -join ',') -eq ($v.fields -join ','))) "$($def.title) view '$($v.title)' fields"
    }
    $defaults = @($l.Views | Where-Object DefaultView)
    Assert ($defaults.Count -eq 1 -and $defaults[0].Title -eq ($def.views | Where-Object { $_['default'] }).title) "$($def.title): one default view"
}
# Spot-check column XML
$ob = $state.Lists['Lists/Obligations']
$status = [xml]$ob.Fields['Status'].SchemaXml
Assert (@($status.Field.CHOICES.CHOICE).Count -eq 5 -and $status.Field.Default -eq 'Not started') 'Obligations.Status choices and default'
Assert (@($status.Field.CHOICES.CHOICE)[4] -eq "N/A $([char]0x2013) Info only") 'en dash preserved in N/A – Info only'
$lk = ([xml]$state.Lists['Lists/Occurrences'].Fields['Obligation'].SchemaXml).Field
Assert ($lk.List.Trim('{}') -eq $ob.Id.ToString() -and $lk.RelationshipDeleteBehavior -eq 'Restrict') 'Occurrences.Obligation looks up Obligations (restrict delete)'
Assert (([xml]$ob.Fields['DueDate'].SchemaXml).Field.Format -eq 'DateOnly') 'Obligations.DueDate is date-only'

Write-Host 'Test 3: second run makes no changes'
$callsBefore = $state.Calls.Count
$o = Invoke-Provision
Assert ($o -match 'No changes needed') 'reports no changes'
$writes = @($state.Calls | Select-Object -Skip $callsBefore | Where-Object { $_ -notmatch '^Invoke-PnPQuery' })
Assert ($writes.Count -eq 0) "no write calls on re-run (got: $($writes -join '; '))"

Write-Host 'Test 4: drift is repaired'
$f = $ob.Fields['Frequency']
$f.SchemaXml = $f.SchemaXml -replace '<CHOICE>Milestone</CHOICE>', '' -replace 'DisplayName="Frequency"', 'DisplayName="Freq"'
$f.Update()
$ob.Fields['Owner'].SchemaXml = $ob.Fields['Owner'].SchemaXml -replace 'Indexed="TRUE"', 'Indexed="FALSE"'; $ob.Fields['Owner'].Update()
$ob.EnableVersioning = $false
($ob.Views | Where-Object Title -eq 'Open').ViewFields = @('LinkTitle')
$o = Invoke-Provision
Assert ((([xml]$f.SchemaXml).Field.CHOICES.CHOICE) -contains 'Milestone') 'missing choice restored'
Assert ($f.Title -eq 'Frequency') 'display name restored'
Assert ($ob.Fields['Owner'].Indexed) 'index restored'
Assert ($ob.EnableVersioning) 'versioning restored'
Assert ((($ob.Views | Where-Object Title -eq 'Open').ViewFields).Count -gt 1) 'view fields restored'
$o = Invoke-Provision
Assert ($o -match 'No changes needed') 'idempotent after repair'

Write-Host 'Test 5: a type clash is reported, not forced'
$ob.Fields['ConditionNo'].SchemaXml = $ob.Fields['ConditionNo'].SchemaXml -replace 'Type="Text"', 'Type="Number"'; $ob.Fields['ConditionNo'].Update()
$o = Invoke-Provision
Assert ($o -match "Column 'ConditionNo' is Number but the schema says Text") 'warns about type clash'

Write-Host ''
if ($failures) { Write-Host "$failures assertion(s) failed" -ForegroundColor Red; exit 1 }
Write-Host 'All provisioning tests passed' -ForegroundColor Green
