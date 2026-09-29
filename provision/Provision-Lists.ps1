#Requires -Version 7.2
<#
.SYNOPSIS
    Creates or updates the SharePoint lists for the Environmental Compliance
    Obligations Register (QTMP OMF). Safe to re-run.

.DESCRIPTION
    Reads provision/list-schema.json and, for each list:
      1. creates the list if it is missing (by URL, e.g. Lists/Obligations)
      2. turns on versioning (major versions, limit -MajorVersionLimit) and sets attachments
      3. renames the Title column and creates or updates every column
      4. creates the indexes (and unique constraints) in the schema
      5. creates or updates every view

    Re-running only changes what differs from the schema. It never deletes
    lists, columns, views or items, and never removes an index.

.PARAMETER SiteUrl
    Full URL of the target site, e.g. https://dtinfrastructurecomau.sharepoint.com/sites/QTMP-EnvCompliance

.PARAMETER ClientId
    Entra ID app registration (client) ID used by PnP.PowerShell to sign in.

.PARAMETER Tenant
    Tenant name or ID (e.g. dtinfrastructurecomau.onmicrosoft.com). Needed with -DeviceLogin.

.PARAMETER DeviceLogin
    Use device-code sign-in instead of an interactive browser window.

.PARAMETER MajorVersionLimit
    Number of major versions kept per item (default 500).

.PARAMETER SetRegionalSettings
    Also set the site's regional settings to en-AU (DD/MM/YYYY), Australia/Brisbane
    time zone and a Monday week start. Requires site owner permission.

.PARAMETER SkipConnect
    Use the existing PnP connection (Connect-PnPOnline already run in this session).

.EXAMPLE
    ./Provision-Lists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QTMP-Env -ClientId 00000000-0000-0000-0000-000000000000 -SetRegionalSettings

.EXAMPLE
    ./Provision-Lists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QTMP-Env -ClientId <id> -WhatIf
    Shows what would be created or changed without changing anything.
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

    [ValidateRange(100, 50000)]
    [int]$MajorVersionLimit = 500,

    [switch]$SetRegionalSettings,

    [switch]$SkipConnect,

    [string]$SchemaPath = (Join-Path $PSScriptRoot 'list-schema.json')
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:Changes = 0

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------
function Write-Step([string]$Message) { Write-Host "==> $Message" -ForegroundColor Cyan }
function Write-Change([string]$Message) { $script:Changes++; Write-Host "    + $Message" -ForegroundColor Green }
function Write-Info([string]$Message) { Write-Host "      $Message" -ForegroundColor DarkGray }

function Get-Opt([hashtable]$Table, [string]$Key, $Default = $null) {
    if ($Table.ContainsKey($Key)) { return $Table[$Key] } else { return $Default }
}

# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------
function Read-Schema([string]$Path) {
    if (-not (Test-Path $Path)) { throw "Schema file not found: $Path" }
    $schema = Get-Content -Raw -Path $Path -Encoding utf8 | ConvertFrom-Json -AsHashtable
    foreach ($list in $schema.lists) {
        $indexCount = @($list.fields | Where-Object { (Get-Opt $_ 'indexed' $false) -or $_.type -eq 'Lookup' }).Count
        if ($list.titleField.indexed) { $indexCount++ }
        if ($indexCount -gt 20) { throw "List '$($list.title)' defines $indexCount indexes; SharePoint allows 20." }
    }
    return $schema
}

function Resolve-Choices([hashtable]$Schema, [hashtable]$Field) {
    $c = Get-Opt $Field 'choices'
    if ($c -is [string] -and $c.StartsWith('@')) { return @($Schema.choiceSets[$c.Substring(1)]) }
    return @($c)
}

# SharePoint TypeAsString for each schema type, used to detect a type clash on re-run.
$script:TypeAsString = @{
    Text = 'Text'; Note = 'Note'; Choice = 'Choice'; DateOnly = 'DateTime'; DateTime = 'DateTime'
    Number = 'Number'; Boolean = 'Boolean'; User = 'User'; UserMulti = 'UserMulti'; URL = 'URL'; Lookup = 'Lookup'
}

function New-FieldXml([hashtable]$Schema, [hashtable]$Field, [hashtable]$ListIds) {
    $doc = [System.Xml.XmlDocument]::new()
    $el = $doc.CreateElement('Field')
    [void]$doc.AppendChild($el)
    $attrs = [ordered]@{
        Name        = $Field.name
        StaticName  = $Field.name
        DisplayName = $Field.displayName
        Required    = $(if (Get-Opt $Field 'required' $false) { 'TRUE' } else { 'FALSE' })
    }
    $description = Get-Opt $Field 'description'
    if ($description) { $attrs.Description = $description }
    switch ($Field.type) {
        'Text' { $attrs.Type = 'Text'; $attrs.MaxLength = '255' }
        'Note' { $attrs.Type = 'Note'; $attrs.RichText = 'FALSE'; $attrs.NumLines = '6'; $attrs.AppendOnly = 'FALSE' }
        'Choice' { $attrs.Type = 'Choice'; $attrs.Format = 'Dropdown'; $attrs.FillInChoice = 'FALSE' }
        'DateOnly' { $attrs.Type = 'DateTime'; $attrs.Format = 'DateOnly' }
        'DateTime' { $attrs.Type = 'DateTime'; $attrs.Format = 'DateTime' }
        'Number' { $attrs.Type = 'Number'; $attrs.Decimals = '0' }
        'Boolean' { $attrs.Type = 'Boolean' }
        'User' { $attrs.Type = 'User'; $attrs.UserSelectionMode = 'PeopleOnly'; $attrs.UserSelectionScope = '0'; $attrs.Mult = 'FALSE' }
        'UserMulti' { $attrs.Type = 'UserMulti'; $attrs.UserSelectionMode = 'PeopleOnly'; $attrs.UserSelectionScope = '0'; $attrs.Mult = 'TRUE' }
        'URL' { $attrs.Type = 'URL'; $attrs.Format = 'Hyperlink' }
        'Lookup' {
            $attrs.Type = 'Lookup'
            $attrs.List = '{' + $ListIds[$Field.lookupList] + '}'
            $attrs.ShowField = 'Title'
            # Restrict: an obligation with child items cannot be deleted (audit trail).
            $attrs.RelationshipDeleteBehavior = 'Restrict'
        }
        default { throw "Unsupported field type '$($Field.type)' on '$($Field.name)'." }
    }
    if ((Get-Opt $Field 'indexed' $false) -or $Field.type -eq 'Lookup') { $attrs.Indexed = 'TRUE' }
    if (Get-Opt $Field 'unique' $false) { $attrs.EnforceUniqueValues = 'TRUE' }
    foreach ($k in $attrs.Keys) { $el.SetAttribute($k, [string]$attrs[$k]) }

    if ($Field.type -eq 'Choice') {
        $choicesEl = $doc.CreateElement('CHOICES')
        foreach ($c in (Resolve-Choices $Schema $Field)) {
            $ce = $doc.CreateElement('CHOICE'); $ce.InnerText = $c; [void]$choicesEl.AppendChild($ce)
        }
        [void]$el.AppendChild($choicesEl)
    }
    if ($Field.ContainsKey('default')) {
        $d = $doc.CreateElement('Default')
        $d.InnerText = if ($Field.type -eq 'Boolean') { if ($Field.default) { '1' } else { '0' } } else { [string]$Field.default }
        [void]$el.AppendChild($d)
    }
    return $doc
}

# ---------------------------------------------------------------------------
# Site
# ---------------------------------------------------------------------------
function Set-SiteRegionalSettings {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    Write-Step 'Regional settings (en-AU, Brisbane, Monday week start)'
    $web = Get-PnPWeb -Includes RegionalSettings.LocaleId, RegionalSettings.FirstDayOfWeek, RegionalSettings.TimeZone, RegionalSettings.TimeZones
    $rs = $web.RegionalSettings
    $tz = $rs.TimeZones | Where-Object { $_.Description -like '*Brisbane*' } | Select-Object -First 1
    if (-not $tz) { Write-Warning 'Brisbane time zone not found; set it manually (see docs/manual-provisioning.md).'; return }
    $needed = ($rs.LocaleId -ne 3081) -or ($rs.TimeZone.Id -ne $tz.Id) -or ($rs.FirstDayOfWeek -ne 1)
    if (-not $needed) { Write-Info 'Already set.'; return }
    if ($PSCmdlet.ShouldProcess($SiteUrl, 'Set regional settings')) {
        try {
            $rs.LocaleId = 3081          # English (Australia): DD/MM/YYYY
            $rs.TimeZone = $tz           # (UTC+10:00) Brisbane, no daylight saving
            $rs.FirstDayOfWeek = 1       # Monday
            $web.Update()
            Invoke-PnPQuery
            Write-Change "Regional settings set to en-AU / $($tz.Description)"
        }
        catch {
            Write-Warning "Could not set regional settings ($($_.Exception.Message)). Set them manually: Site settings > Regional settings."
        }
    }
}

# ---------------------------------------------------------------------------
# Lists
# ---------------------------------------------------------------------------
function Confirm-List {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([hashtable]$Def)
    $list = Get-PnPList -Identity $Def.url -ErrorAction SilentlyContinue
    if (-not $list) {
        if (-not $PSCmdlet.ShouldProcess($Def.title, 'Create list')) { return $null }
        $list = New-PnPList -Title $Def.title -Url $Def.url -Template GenericList -OnQuickLaunch
        Write-Change "Created list '$($Def.title)' at $($Def.url)"
        $list = Get-PnPList -Identity $Def.url
    }
    else {
        Write-Info "List '$($Def.title)' exists."
    }
    Get-PnPProperty -ClientObject $list -Property EnableVersioning, MajorVersionLimit, EnableAttachments, Description | Out-Null
    $attachments = [bool](Get-Opt $Def 'attachments' $false)
    $settingsOk = $list.EnableVersioning -and ($list.MajorVersionLimit -eq $MajorVersionLimit) -and
        ($list.EnableAttachments -eq $attachments) -and ($list.Description -eq $Def.description)
    if (-not $settingsOk -and $PSCmdlet.ShouldProcess($Def.title, 'Set versioning, attachments and description')) {
        Set-PnPList -Identity $list -EnableVersioning $true -MajorVersions $MajorVersionLimit `
            -EnableAttachments $attachments -Description $Def.description | Out-Null
        Write-Change "Versioning on ($MajorVersionLimit major versions), attachments $(if ($attachments) {'on'} else {'off'})"
    }
    return $list
}

function Confirm-Index {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($List, [string]$FieldName, $Existing, [bool]$Indexed, [bool]$Unique)
    if ($Indexed -and -not $Existing.Indexed -and $PSCmdlet.ShouldProcess($FieldName, 'Index column')) {
        Set-PnPField -List $List -Identity $FieldName -Values @{ Indexed = $true } | Out-Null
        Write-Change "Indexed $FieldName"
    }
    if ($Unique -and -not $Existing.EnforceUniqueValues -and $PSCmdlet.ShouldProcess($FieldName, 'Enforce unique values')) {
        Set-PnPField -List $List -Identity $FieldName -Values @{ EnforceUniqueValues = $true } | Out-Null
        Write-Change "Unique values enforced on $FieldName"
    }
}

function Confirm-TitleField {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($List, [hashtable]$TitleDef)
    $title = Get-PnPField -List $List -Identity 'Title'
    if ($title.Title -ne $TitleDef.displayName -and $PSCmdlet.ShouldProcess('Title', "Rename to '$($TitleDef.displayName)'")) {
        Set-PnPField -List $List -Identity 'Title' -Values @{ Title = $TitleDef.displayName } | Out-Null
        Write-Change "Title column renamed to '$($TitleDef.displayName)'"
    }
    Confirm-Index -List $List -FieldName 'Title' -Existing $title `
        -Indexed ([bool](Get-Opt $TitleDef 'indexed' $false)) -Unique ([bool](Get-Opt $TitleDef 'unique' $false))
}

function Update-ExistingField {
    <# Brings an existing column's display name, required flag, description, format,
       choices and default into line with the schema by editing its SchemaXml. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($Existing, [System.Xml.XmlDocument]$Desired)
    $current = [xml]$Existing.SchemaXml
    $cur = $current.DocumentElement
    $want = $Desired.DocumentElement
    $changed = [System.Collections.Generic.List[string]]::new()

    # SharePoint may omit boolean attributes that are FALSE; treat missing as FALSE.
    $booleanAttrs = 'Required', 'RichText', 'FillInChoice'
    foreach ($attr in 'DisplayName', 'Required', 'Description', 'Format', 'RichText', 'FillInChoice', 'UserSelectionMode') {
        if (-not $want.HasAttribute($attr)) { continue }
        $have = $cur.GetAttribute($attr)
        if ($attr -in $booleanAttrs -and $have -eq '') { $have = 'FALSE' }
        if ($have -cne $want.GetAttribute($attr)) {
            $cur.SetAttribute($attr, $want.GetAttribute($attr)); $changed.Add($attr)
        }
    }
    $wantChoices = @($want.SelectNodes('CHOICES/CHOICE') | ForEach-Object { $_.InnerText })
    $curChoices = @($cur.SelectNodes('CHOICES/CHOICE') | ForEach-Object { $_.InnerText })
    if ($want.SelectSingleNode('CHOICES') -and (($wantChoices -join "`n") -cne ($curChoices -join "`n"))) {
        $old = $cur.SelectSingleNode('CHOICES'); if ($old) { [void]$cur.RemoveChild($old) }
        [void]$cur.AppendChild($current.ImportNode($want.SelectSingleNode('CHOICES'), $true))
        $changed.Add('Choices')
    }
    $wantDefault = $want.SelectSingleNode('Default')
    $curDefault = $cur.SelectSingleNode('Default')
    $wd = if ($wantDefault) { $wantDefault.InnerText } else { $null }
    $cd = if ($curDefault) { $curDefault.InnerText } else { $null }
    if ($wantDefault -and $wd -cne $cd) {
        if ($curDefault) { [void]$cur.RemoveChild($curDefault) }
        [void]$cur.AppendChild($current.ImportNode($wantDefault, $true))
        $changed.Add('Default')
    }
    if ($changed.Count -gt 0 -and $PSCmdlet.ShouldProcess($want.GetAttribute('Name'), "Update $($changed -join ', ')")) {
        $Existing.SchemaXml = $cur.OuterXml
        $Existing.Update()
        Invoke-PnPQuery
        Write-Change "Updated $($want.GetAttribute('Name')): $($changed -join ', ')"
    }
}

function Confirm-Field {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([hashtable]$Schema, $List, [hashtable]$Field, [hashtable]$ListIds)
    $xml = New-FieldXml -Schema $Schema -Field $Field -ListIds $ListIds
    $existing = Get-PnPField -List $List -Identity $Field.name -ErrorAction SilentlyContinue
    if (-not $existing) {
        if ($PSCmdlet.ShouldProcess($Field.name, "Create $($Field.type) column")) {
            Add-PnPFieldFromXml -List $List -FieldXml $xml.OuterXml | Out-Null
            Write-Change "Created column $($Field.name) ($($Field.type))"
            # Indexed/EnforceUniqueValues in the XML are normally honoured; confirm explicitly.
            $created = Get-PnPField -List $List -Identity $Field.name
            Get-PnPProperty -ClientObject $created -Property Indexed, EnforceUniqueValues | Out-Null
            Confirm-Index -List $List -FieldName $Field.name -Existing $created `
                -Indexed ([bool]((Get-Opt $Field 'indexed' $false) -or $Field.type -eq 'Lookup')) -Unique ([bool](Get-Opt $Field 'unique' $false))
        }
        return
    }
    Get-PnPProperty -ClientObject $existing -Property SchemaXml, TypeAsString, Indexed, EnforceUniqueValues | Out-Null
    $expectedType = $script:TypeAsString[$Field.type]
    if ($existing.TypeAsString -ne $expectedType) {
        Write-Warning "Column '$($Field.name)' is $($existing.TypeAsString) but the schema says $expectedType. SharePoint cannot change a column's type in place; fix it manually (see docs/manual-provisioning.md, 'Changing a column type')."
        return
    }
    Update-ExistingField -Existing $existing -Desired $xml
    Confirm-Index -List $List -FieldName $Field.name -Existing $existing `
        -Indexed ([bool]((Get-Opt $Field 'indexed' $false) -or $Field.type -eq 'Lookup')) -Unique ([bool](Get-Opt $Field 'unique' $false))
}

function Confirm-View {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($List, [hashtable]$View)
    $fields = [string[]]@($View.fields)
    $rowLimit = [int](Get-Opt $View 'rowLimit' 100)
    $query = [string](Get-Opt $View 'query' '')
    $isDefault = [bool](Get-Opt $View 'default' $false)
    $existing = Get-PnPView -List $List -Identity $View.title -ErrorAction SilentlyContinue
    if (-not $existing) {
        if ($PSCmdlet.ShouldProcess($View.title, 'Create view')) {
            $params = @{ List = $List; Title = $View.title; Fields = $fields; RowLimit = $rowLimit; Paged = $true; SetAsDefault = $isDefault }
            if ($query) { $params.Query = $query }
            Add-PnPView @params | Out-Null
            Write-Change "Created view '$($View.title)'"
        }
        return
    }
    Get-PnPProperty -ClientObject $existing -Property ViewFields, ViewQuery, RowLimit, DefaultView | Out-Null
    $currentFields = [string[]]@($existing.ViewFields)
    $fieldsDiffer = ($currentFields -join ',') -ne ($fields -join ',')
    $queryDiffers = (Format-Caml $existing.ViewQuery) -ne (Format-Caml $query)
    $otherDiffers = ($existing.RowLimit -ne $rowLimit) -or ($isDefault -and -not $existing.DefaultView)
    if (($fieldsDiffer -or $queryDiffers -or $otherDiffers) -and $PSCmdlet.ShouldProcess($View.title, 'Update view')) {
        $values = @{ ViewQuery = $query; RowLimit = [uint32]$rowLimit; Paged = $true }
        if ($isDefault) { $values.DefaultView = $true }
        Set-PnPView -List $List -Identity $existing.Id -Fields $fields -Values $values | Out-Null
        Write-Change "Updated view '$($View.title)'"
    }
}

function Format-Caml([string]$Caml) {
    # SharePoint re-serialises CAML (quotes, whitespace); compare a normalised form.
    if (-not $Caml) { return '' }
    return ($Caml -replace '"', "'" -replace '\s+', ' ' -replace '\s*/>', '/>' -replace '>\s+<', '><').Trim()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$schema = Read-Schema $SchemaPath
Write-Step "Schema v$($schema.version): $($schema.lists.Count) lists"

if (-not $SkipConnect) {
    Write-Step "Connecting to $SiteUrl"
    $connect = @{ Url = $SiteUrl; ClientId = $ClientId }
    if ($DeviceLogin) {
        if (-not $Tenant) { throw '-Tenant is required with -DeviceLogin.' }
        $connect.DeviceLogin = $true; $connect.Tenant = $Tenant
    }
    else {
        $connect.Interactive = $true
        if ($Tenant) { $connect.Tenant = $Tenant }
    }
    Connect-PnPOnline @connect
}

if ($SetRegionalSettings) { Set-SiteRegionalSettings }

# Pass 1: lists (so lookup targets exist before any lookup column is created)
$listIds = @{}
$lists = @{}
foreach ($def in $schema.lists) {
    Write-Step "List: $($def.title)"
    $list = Confirm-List -Def $def
    if ($list) { $lists[$def.key] = $list; $listIds[$def.key] = $list.Id.ToString() }
}

# Pass 2: columns and indexes
foreach ($def in $schema.lists) {
    $list = $lists[$def.key]
    if (-not $list) { Write-Info "Skipping columns for '$($def.title)' (list not created in -WhatIf mode)."; continue }
    Write-Step "Columns: $($def.title)"
    Confirm-TitleField -List $list -TitleDef $def.titleField
    foreach ($field in $def.fields) {
        if ($field.type -eq 'Lookup' -and -not $listIds.ContainsKey($field.lookupList)) {
            Write-Info "Skipping lookup $($field.name) (target list not created in -WhatIf mode)."; continue
        }
        Confirm-Field -Schema $schema -List $list -Field $field -ListIds $listIds
    }
}

# Pass 3: views
foreach ($def in $schema.lists) {
    $list = $lists[$def.key]
    if (-not $list) { continue }
    Write-Step "Views: $($def.title)"
    foreach ($view in $def.views) { Confirm-View -List $list -View $view }
}

# Summary
Write-Step 'Summary'
$rows = foreach ($def in $schema.lists) {
    $list = $lists[$def.key]
    if (-not $list) { [pscustomobject]@{ List = $def.title; Items = '-'; Versioning = '-'; Indexed = '-'; Views = '-' }; continue }
    $list = Get-PnPList -Identity $def.url
    $indexed = @(Get-PnPField -List $list | Where-Object { $_.Indexed }).Count
    [pscustomobject]@{
        List       = $def.title
        Items      = $list.ItemCount
        Versioning = $(if ($list.EnableVersioning) { "On ($($list.MajorVersionLimit))" } else { 'OFF' })
        Indexed    = $indexed
        Views      = @(Get-PnPView -List $list).Count
    }
}
$rows | Format-Table -AutoSize | Out-String | Write-Host
if ($WhatIfPreference) { Write-Host 'WhatIf mode: nothing was changed.' -ForegroundColor Yellow }
elseif ($script:Changes -eq 0) { Write-Host 'No changes needed: the site already matches the schema.' -ForegroundColor Green }
else { Write-Host "$($script:Changes) change(s) applied." -ForegroundColor Green }
