# Offline tests only: all directory objects below are synthetic, NOT live AD data.
[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$TestRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $TestRoot 'test-output' }
[void][IO.Directory]::CreateDirectory($OutputDirectory)
. (Join-Path $TestRoot 'Export-InactiveADUsers.ps1')
$script:Passed=[Collections.Generic.List[string]]::new()
function Assert-Case([bool]$Condition,[string]$Name) { if(-not $Condition){throw "FAIL: $Name"};$script:Passed.Add($Name) }
function Assert-Throws([scriptblock]$Code,[string]$Name) { $caught=$false;try{& $Code | Out-Null}catch{$caught=$true};Assert-Case $caught $Name }
$tokens=$null;$parseErrors=$null
[void][Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Export-InactiveADUsers.ps1'),[ref]$tokens,[ref]$parseErrors)
Assert-Case ($parseErrors.Count -eq 0) 'PowerShell parser: zero errors'
$script:Clock=[DateTimeOffset]'2026-11-01T12:00:00Z'
$script:Queries=[Collections.Generic.List[object]]::new()
$script:FailDC=$false;$script:MissingUser=$false;$script:MissingSourceDC=$false;$script:LargeCount=$false
function New-FixtureUser([int]$Id,[int]$Age,[string]$Name) {
 $ll=[long]0;if($Age-ge 0){$ll=$script:Clock.UtcDateTime.AddDays(-$Age).ToFileTimeUtc()}
 [pscustomobject]@{ObjectGUID=[guid]('00000000-0000-0000-0000-'+$Id.ToString('000000000000'));userAccountControl=512;displayName=$Name;sAMAccountName="user$Id";distinguishedName="CN=$Name,OU=People,DC=example,DC=test";canonicalName="example.test/People/$Name";accountExpires=[long]0;pwdLastSet=$ll;whenCreated=$script:Clock.UtcDateTime.AddDays(-5);lastLogon=$ll;lastLogonTimestamp=$ll;logonCount=[long]7}
}
$script:Users=@(
 (New-FixtureUser 1 95 'Old account'),(New-FixtureUser 2 120 'Active on other DC'),
 (New-FixtureUser 3 -1 'New never used'),(New-FixtureUser 4 200 'Disabled'),
 (New-FixtureUser 5 -1 'Replicated time but zero local history'),(New-FixtureUser 6 90 'Boundary'),
 (New-FixtureUser 7 89 'Recent'),(New-FixtureUser 8 200 '=HYPERLINK("https://invalid.example","Do not execute") & <text>')
)
$script:Users[0].lastLogonTimestamp=$script:Clock.UtcDateTime.AddDays(-100).ToFileTimeUtc()
$script:Users[2].logonCount=0;$script:Users[2].pwdLastSet=0
$script:Users[3].userAccountControl=514
$script:Users[4].lastLogonTimestamp=$script:Clock.UtcDateTime.AddDays(-10).ToFileTimeUtc();$script:Users[4].logonCount=0
$script:Users[5].lastLogon=$script:Clock.UtcDateTime.Date.AddDays(-90).AddHours(23).ToFileTimeUtc();$script:Users[5].lastLogonTimestamp=$script:Clock.UtcDateTime.AddDays(-90).ToFileTimeUtc()
$script:Users[7].accountExpires=[long]::MaxValue
# Deliberate function mocks. The production script has no fixture switch and no mock data.
function Import-Module { [CmdletBinding()]param([string]$Name);if($Name-ne'ActiveDirectory'){throw 'Unexpected import'} }
function Get-ADRootDSE { [CmdletBinding()]param($Server,$Credential);[pscustomobject]@{dnsHostName='dc1.example.test';defaultNamingContext='DC=example,DC=test'} }
function Get-ADDomain { [CmdletBinding()]param($Server,$Credential);[pscustomobject]@{DNSRoot='example.test'} }
function Get-ADDomainController { [CmdletBinding()]param($Filter,$Server,$Credential);if(-not $script:MissingSourceDC){[pscustomobject]@{HostName='dc1.example.test'}};[pscustomobject]@{HostName='dc2.example.test'} }
function Get-ADUser {
 [CmdletBinding()]param($LDAPFilter,$Properties,$Server,$SearchBase,$SearchScope,$ResultPageSize,$ResultSetSize,$Credential)
 $script:Queries.Add([pscustomobject]@{Server=$Server;Filter=$LDAPFilter;Base=$SearchBase;Page=$ResultPageSize;Unlimited=($null-eq$ResultSetSize)})
 if($Server-eq'dc2.example.test'-and$script:FailDC){throw 'Synthetic DC failure'}
 foreach($u in $script:Users){
  if($LDAPFilter-like'*1.2.840.113556.1.4.803*'-and($u.userAccountControl-band 2)){continue}
  if($Server-eq'dc2.example.test'){
   if($script:MissingUser-and$u.sAMAccountName-eq'user1'){continue}
   $ll=$u.lastLogon;if($u.sAMAccountName-eq'user2'){$ll=$script:Clock.UtcDateTime.AddDays(-5).ToFileTimeUtc()}
   $dcCount=[long]11; if($u.sAMAccountName -in @('user3','user5')){$dcCount=0}; if($u.sAMAccountName-eq'user1'-and$script:LargeCount){$dcCount=50000}; [pscustomobject]@{ObjectGUID=$u.ObjectGUID;lastLogon=$ll;logonCount=$dcCount}
  }else{$u}
 }
}
function Read-FixtureSheet([string]$Path){
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $z=[IO.Compression.ZipFile]::OpenRead($Path)
 try{$r=[IO.StreamReader]::new($z.GetEntry('xl/worksheets/sheet1.xml').Open());try{[xml]$r.ReadToEnd()}finally{$r.Dispose()}}finally{$z.Dispose()}
}
function Cell-Text($Sheet,[string]$Address){$n=[Xml.XmlNamespaceManager]::new($Sheet.NameTable);$n.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main');$node=$Sheet.SelectSingleNode("//s:c[@r='$Address']/s:is/s:t",$n);if($node){return $node.InnerText};return ''}
$params=@{AsOf=$script:Clock;TimeZoneId='UTC';Days=90;Force=$true;WarningAction='SilentlyContinue'}
$result=Invoke-InactiveADUserReport @params -OutputPath (Join-Path $OutputDirectory 'synthetic-default.xlsx')
Assert-Case ($result.Rows-eq5) 'Default: five expected inactive/never rows'
Assert-Case ($result.ControllersRead-eq2) 'All discovered controllers queried'
Assert-Case ($script:Queries.Count-eq2) 'One paged user query per DC, not per user'
Assert-Case (@($script:Queries|Where-Object{-not $_.Unlimited -or $_.Page-ne1000}).Count-eq0) 'Paged query has no result-size cap'
$sheet=Read-FixtureSheet $result.OutputPath
Assert-Case ((Cell-Text $sheet 'J4')-eq'95') 'Days based on latest lastLogon, not displayed timestamp'
Assert-Case ((Cell-Text $sheet 'K4')-eq'18') 'Logon count summed across both DCs'
Assert-Case ((Cell-Text $sheet 'C6')-eq'user3') 'Recent never-used user is included regardless of account age'
Assert-Case ((Cell-Text $sheet 'I6')-eq'Never'-and(Cell-Text $sheet 'J6')-eq'Never Logged On') 'Never labels preserved'
Assert-Case ((Cell-Text $sheet 'G6')-eq'') 'Zero pwdLastSet is blank'
Assert-Case ((Cell-Text $sheet 'F8')-eq'Never') 'MaxInt64 accountExpires is Never'
Assert-Case ((Cell-Text $sheet 'A4')-eq'') 'Disabled column is blank for enabled users'
Assert-Case ((Cell-Text $sheet 'B8')-eq$script:Users[7].displayName) 'Formula-looking and XML-special name round-trips as text'
Assert-Case ($sheet.OuterXml-notmatch'<f[ >]') 'No formula elements created from untrusted directory data'
$audit=Get-Content -LiteralPath $result.AuditPath -Raw|ConvertFrom-Json
Assert-Case ($audit.ZeroLastLogonWithNonzeroTimestamp-eq1) 'Contradictory lastLogon and timestamp counted in audit'
Assert-Case ($audit.DirectoryChangesMade-eq$false) 'Read-only result declared'
$safe=Invoke-InactiveADUserReport @params -OutputPath (Join-Path $OutputDirectory 'synthetic-safe.xlsx') -InactivityBasis LastLogon -UseTimestampSafetyNet
Assert-Case ($safe.Rows-eq3) 'LastLogon safety mode excludes recent activity and contradictory zero history'
Assert-Case ((Cell-Text $sheet 'J5')-eq'5') 'Default timestamp filter retains a row with displayed days below threshold'
$elapsed=Invoke-InactiveADUserReport @params -OutputPath (Join-Path $OutputDirectory 'synthetic-elapsed.xlsx') -DayCalculation ElapsedDays
Assert-Case ($elapsed.Rows-eq5 -and (Cell-Text (Read-FixtureSheet $elapsed.OutputPath) 'J7')-eq'89') 'Day rounding changes display only, not timestamp membership'
$before=$script:Queries.Count
$fast=Invoke-InactiveADUserReport @params -OutputPath (Join-Path $OutputDirectory 'synthetic-fast.xlsx') -InactivityBasis LastLogonTimestamp -DaysSinceBasis LastLogonTimestamp -LogonCountMode SourceDC
Assert-Case (($script:Queries.Count-$before)-eq1) 'Timestamp/source-count mode queries one DC only'
Assert-Case ($fast.Rows-eq5) 'Timestamp-based membership includes active-on-other-DC user'
$fs=Read-FixtureSheet $fast.OutputPath
Assert-Case ((Cell-Text $fs 'K4')-eq'7') 'SourceDC mode preserves source count'
$script:FailDC=$true;$failedPath=Join-Path $OutputDirectory 'must-not-exist.xlsx'
if(Test-Path -LiteralPath $failedPath){[IO.File]::Delete($failedPath)}
Assert-Throws {Invoke-InactiveADUserReport @params -OutputPath $failedPath} 'Unreachable DC terminates report'
Assert-Case (-not(Test-Path -LiteralPath $failedPath)) 'DC failure leaves no misleading XLSX'
$script:FailDC=$false;$script:MissingUser=$true
Assert-Throws {Invoke-InactiveADUserReport @params -OutputPath $failedPath} 'Missing baseline object on a DC fails closed'
$script:MissingUser=$false;$script:MissingSourceDC=$true
Assert-Throws {Invoke-InactiveADUserReport @params -OutputPath $failedPath} 'Discovery omitting source DC fails closed'
$script:MissingSourceDC=$false
Assert-Throws {Invoke-InactiveADUserReport @params -OutputPath $failedPath -SearchBase 'DC=outside,DC=test'} 'Cross-domain scope rejected'
Assert-Throws {Invoke-InactiveADUserReport -AsOf $script:Clock -TimeZoneId UTC -OutputPath $result.OutputPath} 'Overwrite requires Force'
$replacement = Invoke-InactiveADUserReport @params -OutputPath $result.OutputPath
Assert-Case ($replacement.Rows -eq 5 -and (Cell-Text (Read-FixtureSheet $replacement.OutputPath) 'C4') -eq 'user1') 'Force replaces an existing workbook successfully'
$empty=Write-InactiveReportWorkbook -Rows @() -Path (Join-Path $OutputDirectory 'synthetic-empty.xlsx') -DomainName 'example.test' -Days 90 -ReportTime $script:Clock.UtcDateTime -Force
$emptyDoc=Read-FixtureSheet $empty
Assert-Case ((Cell-Text $emptyDoc 'K3')-eq'Logon Count') 'Zero results still produce complete headers'
try { $zone=[TimeZoneInfo]::FindSystemTimeZoneById('America/New_York') } catch { $zone=[TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time') }
$t=[DateTimeOffset]'2026-03-09T12:00:00Z';$last=([DateTimeOffset]'2026-03-08T12:30:00Z').UtcDateTime.ToFileTimeUtc()
Assert-Case ((Get-ReportDays $last $t $zone CalendarDays)-eq1-and(Get-ReportDays $last $t $zone ElapsedDays)-eq0) 'Calendar vs elapsed day semantics across DST'
Assert-Throws {Convert-ReportInt64 -1} 'Invalid negative FILETIME not silently rendered Never'
$oldHash=(Get-FileHash $result.OutputPath).Hash
$bad=[pscustomobject]@{'Display Name'=('x'*32768)}
Assert-Throws {Write-InactiveReportWorkbook -Rows @($bad) -Path $result.OutputPath -DomainName example.test -Days 90 -ReportTime $script:Clock.UtcDateTime -Force} 'Overlong cell rejects staging output'
Assert-Case ((Get-FileHash $result.OutputPath).Hash-eq$oldHash) 'Generation failure preserves existing workbook'
# Regression cases derived from the supplied export's split filter/display behavior.
$utcZone=[TimeZoneInfo]::Utc
$fixture=$script:Users[0]
$oldStamp=$fixture.lastLogonTimestamp
$fixture.lastLogonTimestamp=$script:Clock.UtcDateTime.AddDays(-91).ToFileTimeUtc()
$eightyThree=$script:Clock.UtcDateTime.AddDays(-83).ToFileTimeUtc()
$r=ConvertTo-InactiveReportRow $fixture $eightyThree $fixture.lastLogonTimestamp 7 90 $script:Clock $utcZone NearestDays
Assert-Case ($null-ne$r -and $r.'Days Since Last Logon'-eq'83') '91-day timestamp includes a row with 83-day lastLogon'
$r=ConvertTo-InactiveReportRow $fixture 0 $fixture.lastLogonTimestamp 0 90 $script:Clock $utcZone NearestDays
Assert-Case ($null-ne$r -and $r.'Last Logon Timestamp'-ne'Never' -and $r.'Days Since Last Logon'-eq'Never Logged On' -and $r.'Logon Count'-eq'0') 'Old timestamp plus zero logon history preserves contradictory source labels'
$cut=$script:Clock.UtcDateTime.AddDays(-90).ToFileTimeUtc()
Assert-Case ($null-ne(ConvertTo-InactiveReportRow $fixture $cut $cut 7 90 $script:Clock $utcZone NearestDays)) 'Exactly 90 elapsed days is included'
Assert-Case ($null-eq(ConvertTo-InactiveReportRow $fixture $cut ($cut+1) 7 90 $script:Clock $utcZone NearestDays)) 'One FILETIME tick newer than cutoff is excluded'
Assert-Case ($null-ne(ConvertTo-InactiveReportRow $fixture $cut ($cut-1) 7 90 $script:Clock $utcZone NearestDays)) 'One FILETIME tick older than cutoff is included'
foreach($age in @(89.49,89.5,89.51,90.5)) {
 $ft=$script:Clock.UtcDateTime.AddDays(-$age).ToFileTimeUtc()
 $expected=[int][Math]::Round($age,0,[MidpointRounding]::ToEven)
 Assert-Case ((Get-ReportDays $ft $script:Clock $utcZone NearestDays)-eq$expected) "Explicit nearest-day ToEven rounding at $age"
}
$fixture.lastLogonTimestamp=$oldStamp
$script:Users[0].logonCount=40000;$script:LargeCount=$true
$large=Invoke-InactiveADUserReport @params -OutputPath (Join-Path $OutputDirectory 'synthetic-large-count.xlsx')
Assert-Case ((Cell-Text (Read-FixtureSheet $large.OutputPath) 'K4')-eq'90000') '40000 + 50000 DC counters are summed, not capped or replaced by max'
$script:Users[0].logonCount=7;$script:LargeCount=$false
$tests=[ordered]@{Scope='OFFLINE SYNTHETIC FIXTURES; no real directory contacted';PowerShell=$PSVersionTable.PSVersion.ToString();Passed=$script:Passed.Count;Failed=0;Tests=$script:Passed.ToArray()}
$tests|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'tests.json') -Encoding UTF8
$tests|ConvertTo-Json -Depth 5
