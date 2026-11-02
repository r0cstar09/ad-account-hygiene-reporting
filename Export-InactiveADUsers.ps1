#requires -Version 5.1
<#
.SYNOPSIS
Read-only AD report with the layout of the supplied MaxPowerSoft XLSX.
.DESCRIPTION
Queries enabled users, scans every discovered DC by default, and exports an XLSX
using built-in .NET (no Excel, ImportExcel, Python, or original workbook required).
Does NOT disable, delete, unlock, or change any directory object.

Membership defaults to lastLogonTimestamp <= one captured 90-day cutoff; displayed
days independently use max(lastLogon) across DCs. This separation fits the sample. See README
for the unavoidable source-export ambiguities. Stops on failed/incomplete DC scans.
.EXAMPLE
.\Export-InactiveADUsers.ps1 -Days 90 -OutputPath C:\Reports\Inactive90.xlsx
.EXAMPLE
.\Export-InactiveADUsers.ps1 -Server dc01.example.com -SearchBase 'OU=People,DC=example,DC=com' -Days 90
.EXAMPLE
.\Export-InactiveADUsers.ps1 -InactivityBasis LastLogonTimestamp -DaysSinceBasis LastLogonTimestamp -LogonCountMode SourceDC
.EXAMPLE
.\Export-InactiveADUsers.ps1 -InactivityBasis LastLogon -UseTimestampSafetyNet -DayCalculation ElapsedDays
#>
[CmdletBinding()]
param(
    [ValidateRange(1,36500)][int]$Days = 90,
    [string]$OutputPath,
    [string]$Server,
    [string]$SearchBase,
    [System.Management.Automation.PSCredential]$Credential,
    [ValidateSet('LastLogon','LastLogonTimestamp')][string]$InactivityBasis = 'LastLogonTimestamp',
    [ValidateSet('LastLogon','LastLogonTimestamp')][string]$DaysSinceBasis = 'LastLogon',
    [ValidateSet('Sum','SourceDC')][string]$LogonCountMode = 'Sum',
    [ValidateSet('NearestDays','CalendarDays','ElapsedDays')][string]$DayCalculation = 'NearestDays',
    [string]$TimeZoneId = [TimeZoneInfo]::Local.Id,
    [DateTimeOffset]$AsOf = [DateTimeOffset]::Now,
    [switch]$UseTimestampSafetyNet,
    [switch]$SortBySamAccountName,
    [switch]$Force
)

function Get-ReportValue {
    param($Object, [string]$Name, $Default = $null)
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

function Convert-ReportInt64 {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrEmpty([string]$Value)) { return [long]0 }
    $result = [Convert]::ToInt64($Value, [Globalization.CultureInfo]::InvariantCulture)
    if ($result -lt 0) { throw 'Unexpected negative AD integer; refusing to silently turn it into Never/zero.' }
    return $result
}

function Format-ReportFileTime {
    param([long]$Value, [TimeZoneInfo]$Zone, [string]$Empty = 'Never')
    if ($Value -eq 0) { return $Empty }
    $utc = [DateTime]::FromFileTimeUtc($Value)
    return [TimeZoneInfo]::ConvertTimeFromUtc($utc, $Zone).ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ReportDays {
    param([long]$FileTime, [DateTimeOffset]$AsOf, [TimeZoneInfo]$Zone, [string]$Calculation)
    if ($Calculation -eq 'NearestDays') {
        return [int][Math]::Round(($AsOf.UtcDateTime - [DateTime]::FromFileTimeUtc($FileTime)).TotalDays, 0, [MidpointRounding]::ToEven)
    }
    if ($Calculation -eq 'ElapsedDays') {
        return [int][Math]::Floor(($AsOf.UtcDateTime - [DateTime]::FromFileTimeUtc($FileTime)).TotalDays)
    }
    $today = [TimeZoneInfo]::ConvertTime($AsOf, $Zone).Date
    $then = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::FromFileTimeUtc($FileTime), $Zone).Date
    return [int]($today - $then).TotalDays
}

function ConvertTo-InactiveReportRow {
    param($User, [long]$EffectiveLastLogon, [long]$FilterLastLogon, [long]$LogonCount, [int]$Days,
          [DateTimeOffset]$AsOf, [TimeZoneInfo]$Zone, [string]$DayCalculation)
    $uac = Convert-ReportInt64 (Get-ReportValue $User 'userAccountControl')
    if (($uac -band 2) -ne 0) { return }
    # Membership and displayed days are deliberately independent. Compare raw
    # Int64 FILETIMEs, inclusively, with no rounding at the inactivity boundary.
    $cutoff = $AsOf.UtcDateTime.AddDays(-$Days).ToFileTimeUtc()
    if ($FilterLastLogon -gt 0 -and $FilterLastLogon -gt $cutoff) { return }
    $daysValue = 'Never Logged On'
    if ($EffectiveLastLogon -gt 0) {
        $age = Get-ReportDays $EffectiveLastLogon $AsOf $Zone $DayCalculation
        $daysValue = $age.ToString([Globalization.CultureInfo]::InvariantCulture)
    }
    $expiry = Convert-ReportInt64 (Get-ReportValue $User 'accountExpires')
    $expiryText = 'Never'
    if ($expiry -ne 0 -and $expiry -ne [long]::MaxValue) { $expiryText = Format-ReportFileTime $expiry $Zone }
    $createdText = ''
    $created = Get-ReportValue $User 'whenCreated'
    if ($null -ne $created) {
        $createdText = [TimeZoneInfo]::ConvertTimeFromUtc(([DateTime]$created).ToUniversalTime(), $Zone).ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    }
    # All data cells intentionally remain text, as in the supplied export.
    [pscustomobject][ordered]@{
        'Disabled' = ''
        'Display Name' = [string](Get-ReportValue $User 'displayName' '')
        'SAM Account Name' = [string](Get-ReportValue $User 'sAMAccountName' '')
        'Distinguished Name' = [string](Get-ReportValue $User 'distinguishedName' '')
        'Canonical Name' = [string](Get-ReportValue $User 'canonicalName' '')
        'Account Expires On' = $expiryText
        'Password Last Set' = Format-ReportFileTime (Convert-ReportInt64 (Get-ReportValue $User 'pwdLastSet')) $Zone ''
        'When Created' = $createdText
        'Last Logon Timestamp' = Format-ReportFileTime (Convert-ReportInt64 (Get-ReportValue $User 'lastLogonTimestamp')) $Zone
        'Days Since Last Logon' = $daysValue
        'Logon Count' = $LogonCount.ToString([Globalization.CultureInfo]::InvariantCulture)
    }
}

function Add-ReportZipText {
    param($Zip, [string]$Name, [string]$Text)
    $entry = $Zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Write-ReportCell {
    param([Xml.XmlWriter]$Writer, [string]$Address, [int]$Style, [AllowEmptyString()][string]$Value, [switch]$Number)
    if ($Value.Length -gt 32767) { throw "Cell $Address exceeds the Excel 32,767-character limit." }
    $Writer.WriteStartElement('c')
    $Writer.WriteAttributeString('r', $Address)
    $Writer.WriteAttributeString('s', [string]$Style)
    if ($Number) {
        $Writer.WriteElementString('v', $Value)
    } else {
        # Explicit string type prevents formula/DDE execution for names starting =,+,-,@.
        $Writer.WriteAttributeString('t', 'inlineStr')
        $Writer.WriteStartElement('is')
        $Writer.WriteStartElement('t')
        $Writer.WriteAttributeString('xml','space','http://www.w3.org/XML/1998/namespace','preserve')
        $Writer.WriteString($Value)
        $Writer.WriteEndElement(); $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Write-InactiveReportWorkbook {
    param([AllowEmptyCollection()][object[]]$Rows, [string]$Path, [string]$DomainName,
          [int]$Days, [DateTime]$ReportTime, [switch]$Force)
    if ($Rows.Count -gt 1048573) { throw 'Too many rows for one Excel worksheet.' }
    $full = [IO.Path]::GetFullPath($Path)
    if ([IO.Path]::GetExtension($full) -ine '.xlsx') { throw 'OutputPath must end in .xlsx.' }
    if ((Test-Path -LiteralPath $full) -and -not $Force) { throw "Output already exists: $full. Use -Force to replace it." }
    $directory = [IO.Path]::GetDirectoryName($full)
    [void][IO.Directory]::CreateDirectory($directory)
    Add-Type -AssemblyName System.IO.Compression
    $temporary = [IO.Path]::Combine($directory, ('.ad-report-' + [guid]::NewGuid().ToString('N') + '.tmp'))
    $headers = @('Disabled','Display Name','SAM Account Name','Distinguished Name','Canonical Name','Account Expires On','Password Last Set','When Created','Last Logon Timestamp','Days Since Last Logon','Logon Count')
    # Exact style sheet and column dimensions extracted from the supplied workbook;
    # no account data, domain names, creator names or workstation paths are embedded.
    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy\ hh:mm:ss"/></numFmts><fonts count="3" x14ac:knownFonts="1"><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="14"/><color rgb="FF191970"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD3D3D3"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4682B4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF0F8FF"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>'
    $columns = '<cols><col min="1" max="1" width="11.85546875" style="3" customWidth="1"/><col min="2" max="2" width="34.140625" style="4" customWidth="1"/><col min="3" max="3" width="24.140625" style="4" customWidth="1"/><col min="4" max="4" width="131.85546875" style="4" customWidth="1"/><col min="5" max="5" width="103.28515625" style="4" customWidth="1"/><col min="6" max="6" width="19.85546875" style="4" customWidth="1"/><col min="7" max="8" width="18.5703125" style="4" customWidth="1"/><col min="9" max="9" width="23" style="4" customWidth="1"/><col min="10" max="10" width="22.42578125" style="4" customWidth="1"/><col min="11" max="11" width="19.140625" style="4" customWidth="1"/></cols>'
    try {
        $file = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew)
        $zip = [IO.Compression.ZipArchive]::new($file, [IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Add-ReportZipText $zip '[Content_Types].xml' '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
            Add-ReportZipText $zip '_rels/.rels' '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            Add-ReportZipText $zip 'xl/workbook.xml' '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>'
            Add-ReportZipText $zip 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
            Add-ReportZipText $zip 'xl/styles.xml' $styles
            $entry = $zip.CreateEntry('xl/worksheets/sheet1.xml', [IO.Compression.CompressionLevel]::Optimal)
            $settings = [Xml.XmlWriterSettings]::new(); $settings.Encoding = [Text.UTF8Encoding]::new($false); $settings.CloseOutput = $true
            $w = [Xml.XmlWriter]::Create($entry.Open(), $settings)
            try {
                $w.WriteStartDocument(); $w.WriteStartElement('worksheet','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                $last = $Rows.Count + 3
                $w.WriteRaw('<dimension ref="A1:K' + $last + '"/>')
                $w.WriteRaw('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozenSplit"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>')
                $w.WriteRaw($columns); $w.WriteStartElement('sheetData')
                $w.WriteStartElement('row'); $w.WriteAttributeString('r','1')
                Write-ReportCell $w 'A1' 3 $DomainName
                Write-ReportCell $w 'K1' 5 ($ReportTime.ToOADate().ToString('R',[Globalization.CultureInfo]::InvariantCulture)) -Number
                $w.WriteEndElement()
                $w.WriteStartElement('row'); $w.WriteAttributeString('r','2'); $w.WriteAttributeString('ht','18.75'); $w.WriteAttributeString('customHeight','1')
                Write-ReportCell $w 'A2' 2 "Enabled Users who Never Logged on or have Not Logged on for at least $Days days"
                foreach ($column in 1..10) { Write-ReportCell $w (([string][char](65+$column))+'2') 1 '' }
                $w.WriteEndElement()
                $w.WriteStartElement('row'); $w.WriteAttributeString('r','3'); $w.WriteAttributeString('ht','15.75'); $w.WriteAttributeString('customHeight','1')
                foreach ($column in 0..10) { Write-ReportCell $w (([string][char](65+$column))+'3') 6 $headers[$column] }
                $w.WriteEndElement()
                $rowIndex = 4
                foreach ($row in $Rows) {
                    $w.WriteStartElement('row'); $w.WriteAttributeString('r',[string]$rowIndex)
                    foreach ($column in 0..10) {
                        $style = 8; if ($column -eq 0) { $style = 7 }
                        $value = [string](Get-ReportValue $row $headers[$column] '')
                        Write-ReportCell $w (([string][char](65+$column))+$rowIndex) $style $value
                    }
                    $w.WriteEndElement(); $rowIndex++
                }
                $w.WriteEndElement()
                $w.WriteRaw('<mergeCells count="1"><mergeCell ref="A2:K2"/></mergeCells><pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75"/>')
                $w.WriteEndElement(); $w.WriteEndDocument()
            } finally { $w.Dispose() }
        } finally { $zip.Dispose(); $file.Dispose() }
        # Commit only a fully closed ZIP; preserve the old workbook on generation errors.
        if (Test-Path -LiteralPath $full) {
            if (-not $Force) { throw "Output appeared during generation: $full" }
            [IO.File]::Replace($temporary, $full, [System.Management.Automation.Language.NullString]::Value)
        } else { [IO.File]::Move($temporary, $full) }
    } finally {
        if (Test-Path -LiteralPath $temporary) { [IO.File]::Delete($temporary) }
    }
    return $full
}

function Invoke-InactiveADUserReport {
    [CmdletBinding()]
    param(
        [ValidateRange(1,36500)][int]$Days = 90, [string]$OutputPath,
        [string]$Server, [string]$SearchBase, [System.Management.Automation.PSCredential]$Credential,
        [ValidateSet('LastLogon','LastLogonTimestamp')][string]$InactivityBasis = 'LastLogonTimestamp',
    [ValidateSet('LastLogon','LastLogonTimestamp')][string]$DaysSinceBasis = 'LastLogon',
        [ValidateSet('Sum','SourceDC')][string]$LogonCountMode = 'Sum',
        [ValidateSet('NearestDays','CalendarDays','ElapsedDays')][string]$DayCalculation = 'NearestDays',
        [string]$TimeZoneId = [TimeZoneInfo]::Local.Id, [DateTimeOffset]$AsOf = [DateTimeOffset]::Now,
        [switch]$UseTimestampSafetyNet, [switch]$SortBySamAccountName, [switch]$Force
    )
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    $localAsOf = [TimeZoneInfo]::ConvertTime($AsOf,$zone).DateTime
    if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location).Path ("User-Inactive Enabled Users for {0} days_{1}.xlsx" -f $Days,$localAsOf.ToString('yyyyMMdd_HHmmss')) }
    $OutputPath = [IO.Path]::GetFullPath($OutputPath)
    $auditPath = $OutputPath + '.audit.json'
    if (-not $Force -and ((Test-Path -LiteralPath $OutputPath) -or (Test-Path -LiteralPath $auditPath))) { throw 'Report or audit file already exists. Choose a new path or pass -Force.' }
    Import-Module ActiveDirectory -ErrorAction Stop
    $connection = @{ErrorAction = 'Stop'}
    if ($Credential) { $connection.Credential = $Credential }
    $rootArgs = @{} + $connection
    if ($Server) { $rootArgs.Server = $Server }
    $root = Get-ADRootDSE @rootArgs
    $sourceServer = [string]$root.dnsHostName
    if (-not $sourceServer) { throw 'Could not resolve a concrete source domain controller.' }
    $domain = Get-ADDomain -Server $sourceServer @connection
    if (-not $SearchBase) { $SearchBase = [string]$root.defaultNamingContext }
    $domainDN = [string]$root.defaultNamingContext
    if (-not ($SearchBase.Equals($domainDN,[StringComparison]::OrdinalIgnoreCase) -or $SearchBase.EndsWith(','+$domainDN,[StringComparison]::OrdinalIgnoreCase))) { throw 'SearchBase must be within the selected domain naming context.' }
    $properties = @('displayName','sAMAccountName','distinguishedName','canonicalName','userAccountControl','accountExpires','pwdLastSet','whenCreated','lastLogon','lastLogonTimestamp','logonCount','objectGUID')
    $query = @{SearchBase=$SearchBase; SearchScope='Subtree'; ResultPageSize=1000; ResultSetSize=$null; ErrorAction='Stop'}
    if ($Credential) { $query.Credential = $Credential }
    $enabledFilter = '(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
    Write-Verbose "Reading enabled users from $sourceServer"
    $users = @(Get-ADUser -LDAPFilter $enabledFilter -Properties $properties -Server $sourceServer @query)
    $states = @{}
    foreach ($user in $users) {
        $guid = ([guid]$user.ObjectGUID).ToString('D')
        if ($states.ContainsKey($guid)) { throw 'Duplicate objectGUID in baseline query.' }
        $states[$guid] = @{MaxLastLogon=[long]0; SumCount=[long]0; ControllersSeen=0}
    }
    $scanAll = ($InactivityBasis -eq 'LastLogon' -or $DaysSinceBasis -eq 'LastLogon' -or $LogonCountMode -eq 'Sum')
    $controllers = @($sourceServer)
    if ($scanAll) {
        $controllers = @(Get-ADDomainController -Filter * -Server $sourceServer @connection | ForEach-Object { [string]$_.HostName } | Sort-Object -Unique)
        if ($controllers.Count -eq 0 -or $controllers -notcontains $sourceServer) { throw 'Domain-controller discovery was empty or omitted the source DC; report aborted.' }
    }
    $dcStatus = [Collections.Generic.List[object]]::new()
    foreach ($dc in $controllers) {
        Write-Verbose "Reading logon attributes from $dc"
        # Do NOT reapply enabled/inactive predicates on each replica: their state can differ.
        # Query every user in scope once per DC, not N users times D individual requests.
        if ($dc -ieq $sourceServer) { $scan = $users }
        else {
            try { $scan = @(Get-ADUser -LDAPFilter '(&(objectCategory=person)(objectClass=user))' -Properties @('lastLogon','logonCount','objectGUID') -Server $dc @query) }
            catch { throw "Required DC $dc could not be read. No new report written. $($_.Exception.Message)" }
        }
        $seen = @{}
        foreach ($item in $scan) {
            $key = ([guid]$item.ObjectGUID).ToString('D')
            if (-not $states.ContainsKey($key)) { continue }
            if ($seen.ContainsKey($key)) { throw "Duplicate objectGUID returned by $dc" }
            $seen[$key] = $true
            $state = $states[$key]
            $ll = Convert-ReportInt64 (Get-ReportValue $item 'lastLogon')
            $state.MaxLastLogon = [Math]::Max([long]$state.MaxLastLogon,$ll)
            $count = Convert-ReportInt64 (Get-ReportValue $item 'logonCount')
            $state.SumCount = [long]$state.SumCount + $count
            $state.ControllersSeen++
        }
        $missing = $users.Count - $seen.Count
        if ($missing -ne 0) { throw "$dc did not return $missing baseline user(s). Possible replication, scope or permission inconsistency. No new report written; resolve and rerun." }
        $dcStatus.Add([pscustomobject]@{Server=$dc;MatchedUsers=$seen.Count;Complete=$true})
    }
    $rows = [Collections.Generic.List[object]]::new()
    $timestampNewer = 0; $zeroLastWithTimestamp = 0
    foreach ($user in $users) {
        $state = $states[([guid]$user.ObjectGUID).ToString('D')]
        $timestamp = Convert-ReportInt64 (Get-ReportValue $user 'lastLogonTimestamp')
        $effective = [long]$state.MaxLastLogon
        if ($DaysSinceBasis -eq 'LastLogonTimestamp') { $effective = $timestamp }
        $filterLogon = $timestamp
        if ($InactivityBasis -eq 'LastLogon') { $filterLogon = [long]$state.MaxLastLogon }
        if ($timestamp -gt [long]$state.MaxLastLogon) {
            $timestampNewer++
            if ([long]$state.MaxLastLogon -eq 0) { $zeroLastWithTimestamp++ }
            if ($UseTimestampSafetyNet -and $InactivityBasis -eq 'LastLogon') { $filterLogon = $timestamp }
        }
        $count = [long]$state.SumCount
        if ($LogonCountMode -eq 'SourceDC') { $count = Convert-ReportInt64 (Get-ReportValue $user 'logonCount') }
        $row = ConvertTo-InactiveReportRow $user $effective $filterLogon $count $Days $AsOf $zone $DayCalculation
        if ($null -ne $row) { $rows.Add($row) }
    }
    if ($timestampNewer -gt 0) { Write-Warning "$timestampNewer user(s) have a replicated timestamp newer than observed max(lastLogon); $zeroLastWithTimestamp have zero max(lastLogon). Review these inconsistencies. -UseTimestampSafetyNet uses the newer timestamp for classification." }
    $outputRows = @($rows.ToArray())
    if ($SortBySamAccountName) { $outputRows = @($outputRows | Sort-Object 'SAM Account Name') }
    $finishedPath = Write-InactiveReportWorkbook -Rows $outputRows -Path $OutputPath -DomainName $domain.DNSRoot -Days $Days -ReportTime $localAsOf -Force:$Force
    $audit = [ordered]@{
        GeneratedUtc=[DateTimeOffset]::UtcNow.ToString('o'); AsOf=$AsOf.ToString('o'); TimeZone=$zone.Id
        Domain=$domain.DNSRoot; SearchBase=$SearchBase; SourceDC=$sourceServer; Controllers=$dcStatus.ToArray()
        Days=$Days; InactivityBasis=$InactivityBasis; DaysSinceBasis=$DaysSinceBasis; DayCalculation=$DayCalculation; LogonCountMode=$LogonCountMode
        CutoffUtc=$AsOf.UtcDateTime.AddDays(-$Days).ToString('o'); MembershipComparison='Absent/zero or <= cutoff FILETIME, independent of displayed rounded days'
        TimestampSafetyNet=[bool]$UseTimestampSafetyNet; EnabledUsersExamined=$users.Count; ReportRows=$outputRows.Count
        TimestampNewerThanObservedLastLogon=$timestampNewer; ZeroLastLogonWithNonzeroTimestamp=$zeroLastWithTimestamp
        DirectoryChangesMade=$false; NeverLabelMeans='No recorded time in selected attribute(s); not proof of no historical authentication.'
        ScopeLimit='One current domain; no Entra sign-in data, cached offline sign-ins, retired DC history or historical snapshot reconstruction.'
        WorkbookSha256=(Get-FileHash -LiteralPath $finishedPath -Algorithm SHA256).Hash
    }
    $auditJson = ConvertTo-Json -InputObject $audit -Depth 6
    [IO.File]::WriteAllText($auditPath,$auditJson,[Text.UTF8Encoding]::new($false))
    [pscustomobject]@{OutputPath=$finishedPath; AuditPath=$auditPath; Rows=$outputRows.Count; ControllersRead=$controllers.Count; DirectoryChangesMade=$false}
}

# Dot-sourcing defines functions only for inspection/tests. Normal execution queries AD.
if ($MyInvocation.InvocationName -ne '.') { Invoke-InactiveADUserReport @PSBoundParameters }
