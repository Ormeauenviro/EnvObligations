#Requires -Version 7.2
<#
.SYNOPSIS
    Loads the transformed register (import/output/*.json) into the SharePoint lists.
    Re-runnable: upserts on each list's key, so it never creates duplicates.

.DESCRIPTION
    Lists are loaded in dependency order (Source Documents and Obligations first so
    lookups resolve). For each item the key (e.g. Obligation ID) is matched against
    what is already in the list:

      * not found  -> created with every non-blank value
      * found      -> source fields (requirement text, references, Extra Fields...) are
                      updated where they differ from the transform output;
                      live fields (Status, Owner, notes, dates...) are only filled when
                      blank in SharePoint, so the team's work is never overwritten.
                      Use -OverwriteLiveFields to force them (e.g. a reload before go-live).
      * createOnly lists (Status History) are never updated.

    Nothing is ever deleted. Each run writes load-log.csv next to the data.

.PARAMETER SiteUrl
    Full URL of the site the lists were provisioned in.

.PARAMETER ClientId
    Entra ID app registration (client) ID used by PnP.PowerShell.

.PARAMETER DryRun
    Report what would be created and updated without changing anything (same as -WhatIf).
    The plan is written to load-plan.csv next to the data.

.PARAMETER OverwriteLiveFields
    Also overwrite team-owned fields that already have a value. Use only before go-live.

.PARAMETER Lists
    Load only these list keys (e.g. Obligations, Evidence). Default: all, in order.

.EXAMPLE
    ./import/Load-Data.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QTMP-Env -ClientId <id> -DryRun
.EXAMPLE
    ./import/Load-Data.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QTMP-Env -ClientId <id>
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/.+')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$Tenant,
    [switch]$DeviceLogin,
    [switch]$SkipConnect,
    [switch]$DryRun,
    [switch]$OverwriteLiveFields,
    [string[]]$Lists,
    [string]$DataPath = (Join-Path $PSScriptRoot 'output'),
    [string]$SchemaPath = (Join-Path $PSScriptRoot '..' 'provision' 'list-schema.json')
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
if ($DryRun) { $WhatIfPreference = $true }

$Order = 'SourceDocuments', 'ProjectMilestones', 'RoleAssignments', 'Obligations', 'Occurrences',
         'Tasks', 'Evidence', 'FollowUps', 'StatusHistory', 'ReportComponents'
# Lists other lists look up to, and the key column their lookups use.
$LookupKeyField = @{ SourceDocuments = 'DocKey'; Obligations = 'Title' }

$script:Log = [System.Collections.Generic.List[object]]::new()
$script:IdMaps = @{}        # list key -> @{ key value -> SharePoint item ID }
$script:People = @{}        # email -> $true (resolved) / $false (not found)
$script:Invariant = [cultureinfo]::InvariantCulture
$script:Brisbane = $null
foreach ($tzId in 'Australia/Brisbane', 'E. Australia Standard Time') {
    try { $script:Brisbane = [TimeZoneInfo]::FindSystemTimeZoneById($tzId); break } catch { }
}
if (-not $script:Brisbane) { throw 'Could not find the Brisbane time zone on this machine.' }

function Write-Step([string]$Message) { Write-Host "==> $Message" -ForegroundColor Cyan }

# ---------------------------------------------------------------------------
# Value conversion
# ---------------------------------------------------------------------------
function Get-FieldTypes([hashtable]$Schema, [string]$ListKey) {
    $def = $Schema.lists | Where-Object { $_.key -eq $ListKey }
    $types = @{ Title = 'Text' }
    foreach ($f in $def.fields) { $types[$f.name] = $f.type }
    return $types
}

function ConvertTo-NumberText($Value) { ([double]$Value).ToString('0.##########', $script:Invariant) }

function ConvertTo-Comparable {
    <# Normalises a transform (JSON) value so it can be compared with SharePoint's. #>
    param($Value, [string]$Type)
    if ($null -eq $Value -or ($Value -is [string] -and $Value -eq '')) { return '' }
    switch ($Type) {
        'Lookup' {
            $map = $script:IdMaps[$Value.lookup]
            if ($map -and $map.ContainsKey($Value.key)) { return [string]$map[$Value.key] }
            return "new:$($Value.key)"
        }
        'URL' { return [string]$Value.url }
        'Boolean' { return [string][bool]$Value }
        'Number' { return ConvertTo-NumberText $Value }
        'User' { return ([string]$Value).ToLowerInvariant() }
        'UserMulti' { return (@($Value) | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ';' }
        'DateTime' { return ([DateTimeOffset]::Parse($Value, $script:Invariant)).UtcDateTime.ToString('yyyy-MM-ddTHH:mm', $script:Invariant) }
        default { return ([string]$Value -replace "`r`n", "`n").Trim() }
    }
}

function ConvertFrom-SpValue {
    <# Normalises a value read from SharePoint (CSOM FieldValues) for comparison. #>
    param($Value, [string]$Type)
    if ($null -eq $Value) { return '' }
    switch ($Type) {
        'Lookup' { return [string]$Value.LookupId }
        'URL' { return [string]$Value.Url }
        'Boolean' { return [string][bool]$Value }
        'Number' { return ConvertTo-NumberText $Value }
        'User' { return ([string]$Value.Email).ToLowerInvariant() }
        'UserMulti' { return (@($Value) | ForEach-Object { ([string]$_.Email).ToLowerInvariant() } | Sort-Object) -join ';' }
        'DateOnly' {
            # SharePoint returns UTC; a date-only value is midnight in the site's time zone (Brisbane).
            $d = [datetime]$Value
            if ($d.Kind -eq [DateTimeKind]::Utc) { $d = [TimeZoneInfo]::ConvertTimeFromUtc($d, $script:Brisbane) }
            return $d.ToString('yyyy-MM-dd', $script:Invariant)
        }
        'DateTime' { return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm', $script:Invariant) }
        default { return ([string]$Value -replace "`r`n", "`n").Trim() }
    }
}

function ConvertTo-SpValue {
    <# Converts a transform value into what Add-/Set-PnPListItem expects. #>
    param($Value, [string]$Type)
    if ($null -eq $Value) { return $null }
    switch ($Type) {
        'Lookup' {
            $map = $script:IdMaps[$Value.lookup]
            if (-not $map -or -not $map.ContainsKey($Value.key)) {
                if ($WhatIfPreference) { return -1 }   # target would be created earlier in this run
                throw "Lookup target '$($Value.key)' not found in $($Value.lookup)."
            }
            return [int]$map[$Value.key]
        }
        'URL' {
            # PnP splits "url, description" on the comma, so commas in the description are replaced.
            $desc = ([string]$Value.description) -replace ',', ';'
            return "$($Value.url), $desc"
        }
        'DateOnly' {
            # Noon avoids any day shift from time-zone conversion; the column stores the date only.
            return [datetime]::ParseExact($Value, 'yyyy-MM-dd', $script:Invariant).AddHours(12)
        }
        'DateTime' { return ([DateTimeOffset]::Parse($Value, $script:Invariant)).UtcDateTime }
        'Boolean' { return [bool]$Value }
        'Number' { return [double]$Value }
        'User' { if ($script:People[$Value] -eq $false) { return $null } return [string]$Value }
        'UserMulti' {
            $ok = @(@($Value) | Where-Object { $script:People[$_] -ne $false })
            if ($ok.Count -eq 0) { return $null } return [string[]]$ok
        }
        default { return [string]$Value }
    }
}

function Get-ItemKey([hashtable]$Item, [string[]]$KeyFields, [hashtable]$Types) {
    ($KeyFields | ForEach-Object { ConvertTo-Comparable $Item[$_] $Types[$_] }) -join '|'
}

function Get-SpItemKey($SpItem, [string[]]$KeyFields, [hashtable]$Types) {
    ($KeyFields | ForEach-Object { ConvertFrom-SpValue $SpItem.FieldValues[$_] $Types[$_] }) -join '|'
}

# ---------------------------------------------------------------------------
# People
# ---------------------------------------------------------------------------
function Resolve-People([hashtable[]]$Items, [hashtable]$Types) {
    $emails = foreach ($item in $Items) {
        foreach ($name in $item.Keys) {
            if ($Types[$name] -in 'User', 'UserMulti' -and $null -ne $item[$name]) { @($item[$name]) }
        }
    }
    foreach ($email in ($emails | Sort-Object -Unique)) {
        if ($script:People.ContainsKey($email)) { continue }
        if ($WhatIfPreference) { $script:People[$email] = $true; continue }  # not checked in a dry run
        try {
            New-PnPUser -LoginName $email | Out-Null   # ensures the user exists on the site
            $script:People[$email] = $true
        }
        catch {
            $script:People[$email] = $false
            Write-Warning "Person '$email' could not be resolved; that column is left blank. ($($_.Exception.Message))"
            $script:Log.Add([pscustomobject]@{ Action = 'Warning'; List = ''; Key = $email; Fields = ''; Error = 'Person not found' })
        }
    }
}

# ---------------------------------------------------------------------------
# Load one list
# ---------------------------------------------------------------------------
function Import-ListData {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([hashtable]$Data, [hashtable]$Types)
    $listKey = $Data.key
    $list = Get-PnPList -Identity $Data.url -ErrorAction SilentlyContinue
    if (-not $list) { throw "List $($Data.url) not found. Run provision/Provision-Lists.ps1 first." }

    $keyFields = [string[]]$Data.keyFields
    $live = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Data.liveFields))
    $existing = @{}
    foreach ($sp in @(Get-PnPListItem -List $list -PageSize 500)) {
        $existing[(Get-SpItemKey $sp $keyFields $Types)] = $sp
    }
    $items = @($Data.items)
    Resolve-People -Items $items -Types $Types

    $stats = [ordered]@{ Created = 0; Updated = 0; Unchanged = 0; Failed = 0 }
    $idMap = @{}
    $lookupField = $LookupKeyField[$listKey]
    foreach ($sp in $existing.Values) { if ($lookupField) { $idMap[(ConvertFrom-SpValue $sp.FieldValues[$lookupField] 'Text')] = $sp.Id } }

    foreach ($item in $items) {
        $key = Get-ItemKey $item $keyFields $Types
        $sp = $existing[$key]
        try {
            if (-not $sp) {
                $values = @{}
                foreach ($name in $item.Keys) {
                    $v = ConvertTo-SpValue $item[$name] $Types[$name]
                    if ($null -ne $v) { $values[$name] = $v }
                }
                if (-not $WhatIfPreference) {
                    $new = Add-PnPListItem -List $list -Values $values
                    if ($lookupField) { $idMap[[string]$item[$lookupField]] = $new.Id }
                }
                elseif ($lookupField) { $idMap[[string]$item[$lookupField]] = -1 }
                $stats.Created++
                $script:Log.Add([pscustomobject]@{ Action = 'Create'; List = $listKey; Key = $key; Fields = ($values.Keys | Sort-Object) -join ' '; Error = '' })
                continue
            }
            if ($Data.createOnly) { $stats.Unchanged++; continue }
            $changes = @{}
            foreach ($name in $item.Keys) {
                if ($name -in $keyFields) { continue }
                $type = $Types[$name]
                $want = ConvertTo-Comparable $item[$name] $type
                $have = ConvertFrom-SpValue $sp.FieldValues[$name] $type
                if ($want -ceq $have) { continue }
                if ($live.Contains($name)) {
                    if ($want -eq '') { continue }                              # never clear team data
                    if ($have -ne '' -and -not $OverwriteLiveFields) { continue } # team has set it
                }
                $changes[$name] = ConvertTo-SpValue $item[$name] $type
            }
            if ($changes.Count -eq 0) { $stats.Unchanged++; continue }
            if (-not $WhatIfPreference) {
                Set-PnPListItem -List $list -Identity $sp.Id -Values $changes | Out-Null
            }
            $stats.Updated++
            $script:Log.Add([pscustomobject]@{ Action = 'Update'; List = $listKey; Key = $key; Fields = ($changes.Keys | Sort-Object) -join ' '; Error = '' })
        }
        catch {
            $stats.Failed++
            Write-Warning "$listKey ${key}: $($_.Exception.Message)"
            $script:Log.Add([pscustomobject]@{ Action = 'Failed'; List = $listKey; Key = $key; Fields = ''; Error = $_.Exception.Message })
        }
    }
    if ($lookupField) { $script:IdMaps[$listKey] = $idMap }
    return [pscustomobject]@{ List = $Data.list; Created = $stats.Created; Updated = $stats.Updated; Unchanged = $stats.Unchanged; Failed = $stats.Failed }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$schema = Get-Content -Raw -Path $SchemaPath -Encoding utf8 | ConvertFrom-Json -AsHashtable
$selected = if ($Lists) { $Order | Where-Object { $_ -in $Lists } } else { $Order }
if ($Lists -and @($selected).Count -ne @($Lists).Count) { throw "Unknown list key(s). Valid keys: $($Order -join ', ')" }

if (-not $SkipConnect) {
    Write-Step "Connecting to $SiteUrl"
    $connect = @{ Url = $SiteUrl; ClientId = $ClientId }
    if ($DeviceLogin) {
        if (-not $Tenant) { throw '-Tenant is required with -DeviceLogin.' }
        $connect.DeviceLogin = $true; $connect.Tenant = $Tenant
    }
    else { $connect.Interactive = $true; if ($Tenant) { $connect.Tenant = $Tenant } }
    Connect-PnPOnline @connect
}

# Lookups need the target lists' IDs even when only a subset is loaded.
foreach ($target in $LookupKeyField.Keys) {
    if ($target -in $selected) { continue }
    $def = $schema.lists | Where-Object { $_.key -eq $target }
    $list = Get-PnPList -Identity $def.url -ErrorAction SilentlyContinue
    if (-not $list) { continue }
    $map = @{}
    foreach ($sp in @(Get-PnPListItem -List $list -PageSize 500)) { $map[[string]$sp.FieldValues[$LookupKeyField[$target]]] = $sp.Id }
    $script:IdMaps[$target] = $map
}

$summary = foreach ($key in $selected) {
    $path = Join-Path $DataPath "$key.json"
    if (-not (Test-Path $path)) { throw "Missing $path. Run: python import/transform.py" }
    $data = Get-Content -Raw -Path $path -Encoding utf8 | ConvertFrom-Json -AsHashtable
    Write-Step "$($data.list): $($data.count) items"
    Import-ListData -Data $data -Types (Get-FieldTypes $schema $key)
}

Write-Step 'Summary'
$summary | Format-Table -AutoSize | Out-String | Write-Host
# A dry run writes the plan (every create/update and the fields that would change); a real run writes the log.
$logPath = Join-Path $DataPath $(if ($WhatIfPreference) { 'load-plan.csv' } else { 'load-log.csv' })
$script:Log | Export-Csv -Path $logPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
$failed = ($summary | Measure-Object -Property Failed -Sum).Sum
if ($WhatIfPreference) { Write-Host "Dry run: nothing was changed. Counts above show what would happen; details in $logPath" -ForegroundColor Yellow }
elseif ($failed) { Write-Host "$failed item(s) failed - see $logPath" -ForegroundColor Red; exit 1 }
else { Write-Host "Load complete. Log: $logPath" -ForegroundColor Green }
