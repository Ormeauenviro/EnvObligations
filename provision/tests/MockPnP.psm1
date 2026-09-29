# In-memory stand-in for the PnP.PowerShell cmdlets used by Provision-Lists.ps1.
# It mimics the behaviours the script relies on (SharePoint adds attributes such
# as ID/SourceID to SchemaXml, re-serialises CAML with double quotes, creates an
# "All Items" view and a Title column on every new list) so the script's
# create/update/idempotency logic can be tested without a tenant.

$script:State = @{ Lists = [ordered]@{}; Regional = @{ LocaleId = 1033; FirstDayOfWeek = 0; TimeZoneId = 13 }; Calls = [System.Collections.Generic.List[string]]::new() }

function ConvertTo-SpCaml([string]$Caml) { ($Caml -replace "'", '"') -replace '"/>', '" />' }

function Reset-MockPnP { $script:State.Lists = [ordered]@{}; $script:State.Calls.Clear() }
function Get-MockPnPState { $script:State }

function New-MockField([string]$SchemaXml) {
    $f = [pscustomobject]@{ InternalName = ''; Title = ''; TypeAsString = ''; Indexed = $false; EnforceUniqueValues = $false; SchemaXml = $SchemaXml }
    $f | Add-Member -MemberType ScriptMethod -Name Update -Value {
        $x = ([xml]$this.SchemaXml).DocumentElement
        $this.Title = $x.GetAttribute('DisplayName')
        $this.TypeAsString = $x.GetAttribute('Type')
        $this.Indexed = $x.GetAttribute('Indexed') -eq 'TRUE'
        $this.EnforceUniqueValues = $x.GetAttribute('EnforceUniqueValues') -eq 'TRUE'
    }
    $f.Update()
    $f.InternalName = ([xml]$SchemaXml).DocumentElement.GetAttribute('Name')
    return $f
}

function Resolve-MockList($Identity) {
    if ($Identity -is [pscustomobject]) { return $Identity }
    foreach ($l in $script:State.Lists.Values) {
        if ($l.Url -eq $Identity -or $l.Title -eq $Identity -or $l.Id.ToString() -eq "$Identity") { return $l }
    }
    return $null
}

function Connect-PnPOnline { [CmdletBinding()] param($Url, $ClientId, [switch]$Interactive, [switch]$DeviceLogin, $Tenant) }
function Invoke-PnPQuery { [CmdletBinding()] param() $script:State.Calls.Add('Invoke-PnPQuery') }
function Get-PnPProperty { [CmdletBinding()] param($ClientObject, [string[]]$Property) }

function Get-PnPWeb {
    [CmdletBinding()] param([string[]]$Includes)
    $rs = [pscustomobject]@{
        LocaleId = $script:State.Regional.LocaleId; FirstDayOfWeek = $script:State.Regional.FirstDayOfWeek
        TimeZone = [pscustomobject]@{ Id = $script:State.Regional.TimeZoneId }
        TimeZones = @([pscustomobject]@{ Id = 13; Description = '(UTC-08:00) Pacific Time' }, [pscustomobject]@{ Id = 18; Description = '(UTC+10:00) Brisbane' })
    }
    $web = [pscustomobject]@{ RegionalSettings = $rs }
    $web | Add-Member -MemberType ScriptMethod -Name Update -Value {
        $script:State.Regional.LocaleId = $this.RegionalSettings.LocaleId
        $script:State.Regional.FirstDayOfWeek = $this.RegionalSettings.FirstDayOfWeek
        $script:State.Regional.TimeZoneId = $this.RegionalSettings.TimeZone.Id
    }
    return $web
}

function Get-PnPList {
    [CmdletBinding()] param($Identity)
    if ($PSBoundParameters.ContainsKey('Identity')) { return Resolve-MockList $Identity }
    return @($script:State.Lists.Values)
}

function New-PnPList {
    [CmdletBinding()] param([string]$Title, [string]$Url, $Template, [switch]$OnQuickLaunch)
    $list = [pscustomobject]@{
        Id = [guid]::NewGuid(); Title = $Title; Url = $Url; ItemCount = 0; Description = ''
        EnableVersioning = $false; MajorVersionLimit = 50; EnableAttachments = $true
        Fields = [ordered]@{}; Views = [System.Collections.Generic.List[object]]::new()
    }
    $list.Fields['Title'] = New-MockField "<Field ID=`"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}`" Type=`"Text`" Name=`"Title`" DisplayName=`"Title`" Required=`"TRUE`" />"
    $list.Views.Add([pscustomobject]@{ Id = [guid]::NewGuid(); Title = 'All Items'; ViewFields = @('LinkTitle'); ViewQuery = ''; RowLimit = 30; Paged = $true; DefaultView = $true })
    $script:State.Lists[$Url] = $list
    $script:State.Calls.Add("New-PnPList $Url")
    return $list
}

function Set-PnPList {
    [CmdletBinding()] param($Identity, $EnableVersioning, $MajorVersions, $EnableAttachments, $Description)
    $l = Resolve-MockList $Identity
    if ($PSBoundParameters.ContainsKey('EnableVersioning')) { $l.EnableVersioning = [bool]$EnableVersioning }
    if ($PSBoundParameters.ContainsKey('MajorVersions')) { $l.MajorVersionLimit = [int]$MajorVersions }
    if ($PSBoundParameters.ContainsKey('EnableAttachments')) { $l.EnableAttachments = [bool]$EnableAttachments }
    if ($PSBoundParameters.ContainsKey('Description')) { $l.Description = $Description }
    $script:State.Calls.Add("Set-PnPList $($l.Url)")
}

function Get-PnPField {
    [CmdletBinding()] param($List, $Identity)
    $l = Resolve-MockList $List
    if ($PSBoundParameters.ContainsKey('Identity')) { if ($l.Fields.Contains($Identity)) { return $l.Fields[$Identity] } else { return $null } }
    return @($l.Fields.Values)
}

function Add-PnPFieldFromXml {
    [CmdletBinding()] param($List, [string]$FieldXml)
    $l = Resolve-MockList $List
    $doc = [xml]$FieldXml
    $el = $doc.DocumentElement
    $name = $el.GetAttribute('Name')
    if ($l.Fields.Contains($name)) { throw "A field with the name '$name' already exists" }
    if ($el.GetAttribute('EnforceUniqueValues') -eq 'TRUE' -and $el.GetAttribute('Indexed') -ne 'TRUE') { throw 'Unique fields must be indexed' }
    if ($el.GetAttribute('Type') -eq 'Lookup') {
        $target = $el.GetAttribute('List').Trim('{', '}')
        if (-not ($script:State.Lists.Values | Where-Object { $_.Id.ToString() -eq $target })) { throw "Lookup target list $target not found" }
    }
    # SharePoint adds its own attributes on save.
    $el.SetAttribute('ID', "{$([guid]::NewGuid())}")
    $el.SetAttribute('SourceID', "{$($l.Id)}")
    $el.SetAttribute('ColName', 'nvarchar1')
    if ($el.GetAttribute('Required') -eq 'FALSE') { $el.RemoveAttribute('Required') }
    $l.Fields[$name] = New-MockField $doc.OuterXml
    $script:State.Calls.Add("Add-PnPFieldFromXml $name")
}

function Set-PnPField {
    [CmdletBinding()] param($List, $Identity, [hashtable]$Values)
    $l = Resolve-MockList $List
    $f = $l.Fields[$Identity]
    $doc = [xml]$f.SchemaXml
    foreach ($k in $Values.Keys) {
        switch ($k) {
            'Title' { $doc.DocumentElement.SetAttribute('DisplayName', $Values[$k]) }
            'Indexed' { $doc.DocumentElement.SetAttribute('Indexed', $(if ($Values[$k]) { 'TRUE' } else { 'FALSE' })) }
            'EnforceUniqueValues' {
                if ($Values[$k] -and $doc.DocumentElement.GetAttribute('Indexed') -ne 'TRUE') { throw 'Unique fields must be indexed' }
                $doc.DocumentElement.SetAttribute('EnforceUniqueValues', $(if ($Values[$k]) { 'TRUE' } else { 'FALSE' }))
            }
            default { throw "Mock Set-PnPField does not support $k" }
        }
    }
    $f.SchemaXml = $doc.OuterXml
    $f.Update()
    $script:State.Calls.Add("Set-PnPField $Identity $($Values.Keys -join ',')")
}

function Get-PnPView {
    [CmdletBinding()] param($List, $Identity)
    $l = Resolve-MockList $List
    if ($PSBoundParameters.ContainsKey('Identity')) { return $l.Views | Where-Object { $_.Title -eq $Identity -or $_.Id -eq $Identity } | Select-Object -First 1 }
    return @($l.Views)
}

function Add-PnPView {
    [CmdletBinding()] param($List, [string]$Title, [string[]]$Fields, [string]$Query, [int]$RowLimit, [switch]$Paged, [switch]$SetAsDefault)
    $l = Resolve-MockList $List
    foreach ($fld in $Fields) { if ($fld -notin @('LinkTitle', 'ID', 'Created', 'Modified') -and -not $l.Fields.Contains($fld)) { throw "View field $fld not found on $($l.Title)" } }
    if ($SetAsDefault) { foreach ($v in $l.Views) { $v.DefaultView = $false } }
    $l.Views.Add([pscustomobject]@{ Id = [guid]::NewGuid(); Title = $Title; ViewFields = $Fields; ViewQuery = (ConvertTo-SpCaml $Query); RowLimit = $RowLimit; Paged = [bool]$Paged; DefaultView = [bool]$SetAsDefault })
    $script:State.Calls.Add("Add-PnPView $Title")
}

function Set-PnPView {
    [CmdletBinding()] param($List, $Identity, [string[]]$Fields, [hashtable]$Values)
    $l = Resolve-MockList $List
    $v = $l.Views | Where-Object { $_.Id -eq $Identity -or $_.Title -eq $Identity } | Select-Object -First 1
    if ($Fields) { $v.ViewFields = $Fields }
    foreach ($k in $Values.Keys) {
        if ($k -eq 'DefaultView' -and $Values[$k]) { foreach ($o in $l.Views) { $o.DefaultView = $false } }
        $v.$k = if ($k -eq 'ViewQuery') { ConvertTo-SpCaml $Values[$k] } else { $Values[$k] }
    }
    $script:State.Calls.Add("Set-PnPView $($v.Title)")
}

Export-ModuleMember -Function *
