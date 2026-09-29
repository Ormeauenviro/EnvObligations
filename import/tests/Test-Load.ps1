#Requires -Version 7.2
<#
.SYNOPSIS
    Runs import/Load-Data.ps1 against the in-memory MockPnP module (after provisioning
    the lists with Provision-Lists.ps1) and checks: dry run, full load, re-run without
    duplicates, protection of team-owned fields, repair of source fields, unresolved people.
    Usage: pwsh import/tests/Test-Load.ps1   (run `python import/transform.py` first)
#>
$ErrorActionPreference = 'Stop'
$repo = Resolve-Path (Join-Path $PSScriptRoot '..' '..')
$provision = Join-Path $repo 'provision' 'Provision-Lists.ps1'
$load = Join-Path $repo 'import' 'Load-Data.ps1'
$source = Join-Path $repo 'import' 'output'
Import-Module (Join-Path $repo 'provision' 'tests' 'MockPnP.psm1') -Force

# Work on a copy of the transform output so load-log.csv and test edits stay out of the repo.
$data = Join-Path ([IO.Path]::GetTempPath()) "envob-load-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $data | Out-Null
Copy-Item (Join-Path $source '*.json') $data

$failures = 0
function Assert([bool]$Condition, [string]$Message) {
    if ($Condition) { Write-Host "  PASS $Message" -ForegroundColor Green }
    else { Write-Host "  FAIL $Message" -ForegroundColor Red; $script:failures++ }
}
function Invoke-Load([switch]$DryRun, [switch]$OverwriteLiveFields, [string[]]$Lists) {
    $p = @{ SiteUrl = 'https://contoso.sharepoint.com/sites/test'; ClientId = 'x'; SkipConnect = $true; DataPath = $data }
    if ($DryRun) { $p.DryRun = $true }
    if ($OverwriteLiveFields) { $p.OverwriteLiveFields = $true }
    if ($Lists) { $p.Lists = $Lists }
    $out = & $load @p 6>&1 3>&1 | Out-String
    return $out
}
function Get-Items([string]$Url) { $l = (Get-MockPnPState).Lists[$Url]; if ($l.PSObject.Properties.Name -contains 'Items') { @($l.Items) } else { @() } }
function Get-Item([string]$Url, [string]$Title) { Get-Items $Url | Where-Object { $_.FieldValues['Title'] -eq $Title } }
function Get-Json([string]$Key) { Get-Content -Raw (Join-Path $data "$Key.json") | ConvertFrom-Json -AsHashtable }

Reset-MockPnP
& $provision -SiteUrl 'https://contoso.sharepoint.com/sites/test' -ClientId 'x' -SkipConnect 6>&1 | Out-Null
$keys = 'SourceDocuments', 'ProjectMilestones', 'RoleAssignments', 'Obligations', 'Occurrences', 'Tasks', 'Evidence', 'FollowUps', 'StatusHistory', 'ReportComponents'

Write-Host 'Test 1: dry run writes nothing'
$o = Invoke-Load -DryRun
$total = ($keys | ForEach-Object { @(Get-Items (Get-Json $_).url).Count } | Measure-Object -Sum).Sum
Assert ($total -eq 0) 'no items created'
Assert ($o -match 'Dry run: nothing was changed') 'reports dry run'
Assert ($o -notmatch 'WARNING') 'no warnings in dry run'
$plan = Import-Csv (Join-Path $data 'load-plan.csv')
Assert (@($plan | Where-Object Action -eq 'Create').Count -eq 641) "load-plan.csv lists 641 planned creates (got $(@($plan).Count))"
Assert ($o -notmatch 'What if:') 'dry run is quiet (no per-item What if lines)'

Write-Host 'Test 2: full load creates every item'
$o = Invoke-Load
Assert ($o -match 'Load complete') 'load completes without failures'
foreach ($k in $keys) {
    $j = Get-Json $k
    Assert (@(Get-Items $j.url).Count -eq $j.count) "$k`: $($j.count) items"
}
$docs = Get-Items 'Lists/SourceDocuments'
$ob = Get-Item 'Lists/Obligations' 'MP-NAC-003'
$doc = $docs | Where-Object { $_.FieldValues['DocKey'] -eq 'SDA-2309-36966' }
Assert ($ob.FieldValues['SourceDocument'].LookupId -eq $doc.Id) 'Obligation -> Source Document lookup resolves to the right item'
$occ = Get-Item 'Lists/Occurrences' 'MP-NAC-003|COMMENCE'
Assert ($occ.FieldValues['Obligation'].LookupId -eq $ob.Id) 'Occurrence -> Obligation lookup resolves'
Assert ($occ.FieldValues['EvidenceURL'].Url -like 'https://dtinfrastructurecomau.sharepoint.com/*') 'hyperlink stored'
Assert ((Get-Item 'Lists/Obligations' 'CEMP-001').FieldValues['Owner'].Email -eq 'DanielBlunt@dtinfrastructure.com.au') 'owner stored as person'
$nn10 = (Get-Item 'Lists/Obligations' 'NN-DA-010').FieldValues['DueDate']
Assert ($nn10 -eq [datetime]::new(2026, 7, 30, 14, 0, 0, [DateTimeKind]::Utc)) "date-only 31/07/2026 stored as Brisbane midnight (got $($nn10.ToString('o')))"
$ra = Get-Item 'Lists/RoleAssignments' 'Environmental and Sustainability Team'
Assert (@($ra.FieldValues['AdditionalPeople']).Email -contains 'EthanLittle@dtinfrastructure.com.au') 'multi-person column stored'

Write-Host 'Test 3: re-run creates no duplicates and changes nothing'
$calls = (Get-MockPnPState).Calls.Count
$o = Invoke-Load
$writes = @((Get-MockPnPState).Calls | Select-Object -Skip $calls | Where-Object { $_ -match '^(Add|Set)-PnPListItem' })
Assert ($writes.Count -eq 0) "no writes on re-run (got $($writes.Count): $($writes | Select-Object -First 3))"
foreach ($k in 'Obligations', 'Evidence', 'StatusHistory') { $j = Get-Json $k; Assert (@(Get-Items $j.url).Count -eq $j.count) "$k still $($j.count) items" }

Write-Host 'Test 4: team-owned fields are kept, source fields are repaired, blanks are filled'
$c1 = Get-Item 'Lists/Obligations' 'CEMP-001'
$c1.FieldValues['Status'] = 'Complete'
$c1.FieldValues['ComplianceNotes'] = 'Team note'
$c1.FieldValues['Requirement'] = 'accidentally edited'
$c2 = Get-Item 'Lists/Obligations' 'CEMP-002'
$c2.FieldValues['Owner'] = $null
$o = Invoke-Load -Lists Obligations
Assert ($c1.FieldValues['Status'] -eq 'Complete') 'Status set by the team is kept'
Assert ($c1.FieldValues['ComplianceNotes'] -eq 'Team note') 'notes set by the team are kept'
Assert ($c1.FieldValues['Requirement'] -like 'Groundwater (ambient)*') 'source field restored from the register'
Assert ($c2.FieldValues['Owner'].Email -eq 'DanielBlunt@dtinfrastructure.com.au') 'blank owner filled'

Write-Host 'Test 5: -OverwriteLiveFields resets team-owned fields'
$o = Invoke-Load -Lists Obligations -OverwriteLiveFields
Assert ($c1.FieldValues['Status'] -eq 'Not started') 'Status reset to the transform value'

Write-Host 'Test 6: loading one child list alone still resolves lookups'
$tasks = Get-Items 'Lists/ObligationTasks'
(Get-MockPnPState).Lists['Lists/ObligationTasks'].Items.Clear()
$o = Invoke-Load -Lists Tasks
$t = Get-Items 'Lists/ObligationTasks'
Assert ($t.Count -eq 4 -and $t[0].FieldValues['Obligation'].LookupId -eq (Get-Item 'Lists/Obligations' 'NN-DA-001').Id) 'Tasks reloaded with Obligation lookup'

Write-Host 'Test 7: an unknown person is reported and left blank, the rest still loads'
$ra = Get-Json 'RoleAssignments'
$ra.items[4].PrimaryPerson = 'someone.unknown@dtinfrastructure.com.au'   # Construction Manager
$ra | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $data 'RoleAssignments.json') -Encoding utf8
(Get-MockPnPState).Lists['Lists/RoleAssignments'].Items.Clear()
$o = Invoke-Load -Lists RoleAssignments
Assert ($o -match "someone.unknown@dtinfrastructure.com.au' could not be resolved") 'warning shown'
Assert (@(Get-Items 'Lists/RoleAssignments').Count -eq 10) 'all 10 roles loaded'
Assert ($null -eq (Get-Item 'Lists/RoleAssignments' 'Construction Manager').FieldValues['PrimaryPerson']) 'unresolved person left blank'

Remove-Item -Recurse -Force $data
Write-Host ''
if ($failures) { Write-Host "$failures assertion(s) failed" -ForegroundColor Red; exit 1 }
Write-Host 'All load tests passed' -ForegroundColor Green
