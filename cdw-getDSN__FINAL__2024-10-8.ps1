<#
.Synopsis
   CDW Script to capture a Window systems DSN Information
.DESCRIPTION
   Script to capture all Data Source Name (DSN) Information - Attribute, PropertyValue, 
   DriverName, DSNType, KeyValuePair, and Platform values as-well-as captures all the
   ODBC Driver Information - Name, Attribute, KeyValuePair, and Platform values. This
   information can be used to re-create a missing/deleted DSN from the said Windows 
   system and is used as a contingency planning.
.EXAMPLE
   .\cdw-getDSN.ps1
.EXAMPLE
   . .\cdw-getDSN.ps1
   getDSN
.EXAMPLE
   icm -cn ServerName -filename .\cdw-getDSN.ps1   #executed from the directory where the script resides, otherwise enter complete path
.INPUTS
   None
.OUTPUTS
   Outputs report to 'C:\cdw_coop\ETL\ODBC_DSN\Report' if it exists and if not reports are sent
   to current path of the user executing the report.
.NOTES
   Script to help track unplanned changes to the ETL DSN information. Best to schedule task to run
   daily.
.COMPONENT
   The component this cmdlet belongs to Jeff Giese, member of the Corporate Data Warehouse System 
   Administrator Team (CDW SA Team)
.ROLE
   The role this cmdlet belongs to CDW Teams\ETL Team
.FUNCTIONALITY
   Disaster Recover (DR) purposes in the event a change was made and will compare the 
   last two reports.
#>
#Referenced the below site within the pscustomobject section:
#https://powershellisfun.com/2023/03/13/using-pscustomobject-in-powershell/
function cdw-getDSN{
    [CmdletBinding()]
    [Alias("getDSN")]

$cn = ($env:COMPUTERNAME)
[datetime]$date = get-date
$dir_b4 = $pwd.Path
$ea_b4 = $ErrorActionPreference
$HourMinDATE = $date.ToString("HH" + "mm" + "__yyyy_MM_dd")
# save to whatever location you want and if none selected it will save to the user's default folder location - $env:USERPROFILE\Documents
$report_dir = "C:\cdw_coop\ETL\ODBC_DSN\Report"
$scriptEXE_NTacct = ($env:USERDOMAIN + "\" + $env:USERNAME)

$report_HT = [ordered]@{}
$report_HT += @{
    date_timestamp = $($date)
    default_directory = $($pwd.path)
    executed_SERVER = $($cn)
    script = "cdw-getDSN"
    script_description = "Data Source Name (DSN) information"
    script_executed_NTacct = $($scriptEXE_NTacct)
}

$ea = "silentlycontinue"
$ErrorActionPreference = $ea

#Clear-Host

$verify_reportDIR = test-path $report_dir

if($verify_reportDIR -eq $false){
    ($report_dir) = $($pwd.Path)
}
$report_HT += @{
    scriptREPORT_location = $($report_dir)
    scriptREPORT_type = "TXT extension"
}

sl $report_dir

#DSN Name,Platform,DriverName,Type,Attribute
$ODBCdsn_info = Get-OdbcDsn | select * | select Name,Platform,DriverName,DSNType,Attribute | sort Name

#2024-10-7 the below DSN Attribute info logic was gathered if the attribute count was less than '12' to gather the info, but i may have to
#redo so it's more like the Driver Attribute for gathering the info - i'll have to verify all the data with the old/possible new code and making
#a note of this so it will remind me to review this later on...
$ODBCdsn_ATTRIBUTE = foreach($o in $ODBCdsn_info){
    if($o.Attribute.count -LE 12){
    #Attribute with less than 12 params (seen this only in Intersystems DSN entries)
        $LE12 = $o | ?{$_.attribute.count -le 12} | select platform,name,attribute | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json

        [pscustomobject]@{
            DSN_Name = $o.Name
            DSN_Platform = $o.Platform
            DSN_DriverName = $o.DriverName
            DSN_Type = $o.DsnType
    #Attribute
            ATTRIBUTE_AB = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ab))){
                    $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ab) -join ",")
                }
            ATTRIBUTE_Authentication_Method = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Authentication Method"))){ #only the Healthshare Prod have this param
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Authentication Method") -join ",")
                }
            ATTRIBUTE_BI = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty BI))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty BI) -join ",")
                }
            ATTRIBUTE_BoolsAsChar = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty BoolsAsChar))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty BoolsAsChar) -join ",")
                }
            ATTRIBUTE_ByteaAsLongVarBinary = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ByteaAsLongVarBinary))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ByteaAsLongVarBinary) -join ",")
                }
            ATTRIBUTE_CommLog = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty CommLog))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty CommLog) -join ",")
                }
            ATTRIBUTE_ConnSettings = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ConnSettings))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ConnSettings) -join ",")
                }
            ATTRIBUTE_D6 = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty D6))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty D6))
                }                      
            ATTRIBUTE_Database = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Database))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Database) -join ",")
                }
            ATTRIBUTE_Debug = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Debug))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Debug))
                }
            ATTRIBUTE_Description = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Description))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Description) -join ",")
                }
            ATTRIBUTE_ExtraSysTablePrefixes = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ExtraSysTablePrefixes))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ExtraSysTablePrefixes) -join ",")
                }
            ATTRIBUTE_FakeOidIndex = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty FakeOidIndex))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty FakeOidIndex) -join ",")
                }
            ATTRIBUTE_Fetch = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Fetch))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Fetch) -join ",")
                }
            ATTRIBUTE_Host = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Host))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Host) -join ",")
                }
            ATTRIBUTE_KeepaliveInterval = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty KeepaliveInterval))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty KeepaliveInterval) -join ",")
                }
            ATTRIBUTE_KeepaliveTime = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty KeepaliveTime))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty KeepaliveTime) -join ",")
                }
            ATTRIBUTE_LFConversion = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty LFConversion))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty LFConversion))
                }
            ATTRIBUTE_LowerCaseIdentifier = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty LowerCaseIdentifier))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty LowerCaseIdentifier) -join ",")
                }
            ATTRIBUTE_MaxLongVarcharSize = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty MaxLongVarcharSize))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty MaxLongVarcharSize) -join ",")
                }
            ATTRIBUTE_MaxVarcharSize = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty MaxVarcharSize))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty MaxVarcharSize) -join ",")
                }
            ATTRIBUTE_Namespace = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Namespace))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Namespace) -join ",")
                }
            ATTRIBUTE_Parse = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Parse))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Parse) -join ",")
                }
            ATTRIBUTE_Port = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Port))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Port) -join ",")
                }
            ATTRIBUTE_pqopt = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty pqopt))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty pqopt))
                }
            ATTRIBUTE_Protocol = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Protocol))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Protocol))
                }
            ATTRIBUTE_Query_Timeout = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Query Timeout"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Query Timeout") -join ",")
                }
            ATTRIBUTE_ReadOnly = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ReadOnly))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ReadOnly) -join ",")
                }
            ATTRIBUTE_RowVersioning = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty RowVersioning))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty RowVersioning) -join ",")
                }
            ATTRIBUTE_Security_Level = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Security Level"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Security Level") -join ",")
                }
            ATTRIBUTE_Servername = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Servername))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Servername) -join ",")
                }
            ATTRIBUTE_Service_Principal_Name = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Service Principal Name"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Service Principal Name") -join ",")
                }
            ATTRIBUTE_ShowOidColumn = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ShowOidColumn))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ShowOidColumn) -join ",")
                }
            ATTRIBUTE_ShowSystemTables = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty ShowSystemTables))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty ShowSystemTables) -join ",")
                }
            ATTRIBUTE_SSL_Server_Name = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "SSL Server Name"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "SSL Server Name") -join ",")
                }
            ATTRIBUTE_SSLmode = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty SSLmode))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty SSLmode) -join ",")
                }
            ATTRIBUTE_Static_Cursors = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Static Cursors"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Static Cursors") -join ",")
                }
            ATTRIBUTE_TextAsLongVarchar = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty TextAsLongVarchar))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty TextAsLongVarchar) -join ",")
                }
            ATTRIBUTE_Timeout = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Timeout))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Timeout) -join ",")
                }
            ATTRIBUTE_TrueIsMinus1 = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty TrueIsMinus1))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty TrueIsMinus1) -join ",")
                }
            ATTRIBUTE_Unicode_SQLTypes = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty "Unicode SQLTypes"))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty "Unicode SQLTypes") -join ",")
                }
            ATTRIBUTE_UniqueIndex = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UniqueIndex))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UniqueIndex) -join ",")
                }
            ATTRIBUTE_UnknownsAsLongVarchar = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UnknownsAsLongVarchar))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UnknownsAsLongVarchar) -join ",")
                }
            ATTRIBUTE_UnknownSizes = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UnknownSizes))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UnknownSizes) -join ",")
                }
            ATTRIBUTE_UpdatableCursors = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UpdatableCursors))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UpdatableCursors) -join ",")
                }
            ATTRIBUTE_UseDeclareFetch = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UseDeclareFetch))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UseDeclareFetch) -join ",")
                }
            ATTRIBUTE_Username = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty Username))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty Username) -join ",")
                }
            ATTRIBUTE_UseServerSidePrepare = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty UseServerSidePrepare))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty UseServerSidePrepare) -join ",")
                }
            ATTRIBUTE_XaOpt = if([string]::IsNullOrWhiteSpace(@($LE12 | select -ExpandProperty XaOpt))){
                $null
                }
                else{
                    (@($LE12 | select -ExpandProperty XaOpt) -join ",")
                }
            }
        }
        else{
    [pscustomobject]@{
            DSN_Name = Get-OdbcDsn -Name $O.Name | select -ExpandProperty Name
            DSN_Platform = $o.platform
            DSN_DriverName = $o.DriverName
            DSN_Type = $o.DsnType

    #Attribute
            ATTRIBUTE_AB = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select -ExpandProperty attribute | convertto-json -Compress | ConvertFrom-Json | select -ExpandProperty ab))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select -ExpandProperty attribute | convertto-json -Compress | ConvertFrom-Json | select -ExpandProperty ab) -join ",")
            }
        ATTRIBUTE_Authentication_Method = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty syncroot | select -ExpandProperty "Authentication Method"))){ #only the Healthshare Prod have this param
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty syncroot | select -ExpandProperty "Authentication Method"))
            }
        ATTRIBUTE_BI = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty BI))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty BI))
            }
        ATTRIBUTE_BoolsAsChar = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty BoolsAsChar))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty BoolsAsChar))
            }
        ATTRIBUTE_ByteaAsLongVarBinary = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ByteaAsLongVarBinary))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ByteaAsLongVarBinary))
            }
        ATTRIBUTE_CommLog = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty CommLog))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty CommLog))
            }
        ATTRIBUTE_ConnSettings = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ConnSettings))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ConnSettings))
            }
        ATTRIBUTE_D6 = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty D6))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty D6))
            }                      
        ATTRIBUTE_Database = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Database))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Database))
            }
        ATTRIBUTE_Debug = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Debug))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Debug))
            }
        ATTRIBUTE_Description = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Description))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Description))
            }
        ATTRIBUTE_ExtraSysTablePrefixes = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ExtraSysTablePrefixes))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ExtraSysTablePrefixes))
            }
        ATTRIBUTE_FakeOidIndex = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty FakeOidIndex))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty FakeOidIndex))
            }
        ATTRIBUTE_Fetch = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Fetch))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Fetch))
            }
        ATTRIBUTE_Host = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Host))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Host))
            }
        ATTRIBUTE_KeepaliveInterval = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty KeepaliveInterval))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty KeepaliveInterval))
            }
        ATTRIBUTE_KeepaliveTime = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty KeepaliveTime))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty KeepaliveTime))
            }
        ATTRIBUTE_LFConversion = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty LFConversion))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty LFConversion))
            }
        ATTRIBUTE_LowerCaseIdentifier = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty LowerCaseIdentifier))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty LowerCaseIdentifier))
            }
        ATTRIBUTE_MaxLongVarcharSize = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty MaxLongVarcharSize))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty MaxLongVarcharSize))
            }
        ATTRIBUTE_MaxVarcharSize = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty MaxVarcharSize))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty MaxVarcharSize))
            }
        ATTRIBUTE_Namespace = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Namespace))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Namespace))
            }
        ATTRIBUTE_Parse = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Parse))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Parse))
            }
        ATTRIBUTE_Port = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Port))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Port))
            }
        ATTRIBUTE_pqopt = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty pqopt))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty pqopt))
            }
        ATTRIBUTE_Protocol = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Protocol))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Protocol))
            }
        ATTRIBUTE_Query_Timeout = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Query Timeout"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Query Timeout"))
            }
        ATTRIBUTE_ReadOnly = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ReadOnly))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ReadOnly))
            }
        ATTRIBUTE_RowVersioning = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty RowVersioning))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty RowVersioning))
            }
        ATTRIBUTE_Security_Level = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Security Level"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Security Level"))
            }
        ATTRIBUTE_Servername = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Servername))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Servername))
            }
        ATTRIBUTE_Service_Principal_Name = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Service Principal Name"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Service Principal Name"))
            }
        ATTRIBUTE_ShowOidColumn = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ShowOidColumn))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ShowOidColumn))
            }
        ATTRIBUTE_ShowSystemTables = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ShowSystemTables))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty ShowSystemTables))
            }
        ATTRIBUTE_SSL_Server_Name = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "SSL Server Name"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "SSL Server Name"))
            }
        ATTRIBUTE_SSLmode = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty SSLmode))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty SSLmode))
            }
        ATTRIBUTE_Static_Cursors = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty snycroot | select -ExpandProperty "Static Cursors"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty snycroot | select -ExpandProperty "Static Cursors"))
            }
        ATTRIBUTE_TextAsLongVarchar = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty TextAsLongVarchar))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty TextAsLongVarchar))
            }
        ATTRIBUTE_Timeout = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Timeout))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Timeout))
            }
        ATTRIBUTE_TrueIsMinus1 = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty TrueIsMinus1))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty TrueIsMinus1))
            }
        ATTRIBUTE_Unicode_SQLTypes = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Unicode SQLTypes"))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty "Unicode SQLTypes"))
            }
        ATTRIBUTE_UniqueIndex = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UniqueIndex))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UniqueIndex))
            }
        ATTRIBUTE_UnknownsAsLongVarchar = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UnknownsAsLongVarchar))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UnknownsAsLongVarchar))
            }
        ATTRIBUTE_UnknownSizes = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UnknownSizes))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UnknownSizes))
            }
        ATTRIBUTE_UpdatableCursors = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UpdatableCursors))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UpdatableCursors))
            }
        ATTRIBUTE_UseDeclareFetch = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UseDeclareFetch))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UseDeclareFetch))
            }
        ATTRIBUTE_Username = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Username))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty Username))
            }
        ATTRIBUTE_UseServerSidePrepare = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UseServerSidePrepare))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty UseServerSidePrepare))
            }
        ATTRIBUTE_XaOpt = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty XaOpt))){
            $null
            }
            else{
                (@(Get-OdbcDsn $o.Name | select * |select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty XaOpt))
            }
        }
    }
}
#reporting options:
#$ODBCdsn_ATTRIBUTE | select * | ft -AutoSize -Wrap | Out-File $report_dir\DSN_Information.txt -Append
#$ODBCdsn_ATTRIBUTE | ConvertTo-Csv -NoTypeInformation | Out-File $report_dir\DSN_Information.txt -append
#$ODBCdsn_ATTRIBUTE | Out-File $report_dir\DSN_Information.txt -Append



#ODBC installed driver information
#Write-Output "ODBC Driver Information:" #| Out-File $report -Append
$o = ""
$ODBCdriver_info = Get-OdbcDriver | select * | select Name,Platform,Attribute | sort Name
#$ODBCdriver_NAME = $ODBCdriver_info.name
#$ODBCname_NOdup = $ODBCdriver_info | select -ExpandProperty name -Unique
#$dup_name = diff $ODBCdriver_NAME $ODBCname_NOdup | select -ExpandProperty inputobject

$ODBCdriver_value = foreach($o in $ODBCdriver_info){
    if($o.Platform -eq "64-bit"){
        $x64 = $o | select platform,name,attribute | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json

[pscustomobject]@{
        ODBC_driverNAME = $o.name
        ODBC_platform = $o.platform
    #ATTRIBUTE
        DRIVER_APILevel = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty apilevel))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty apilevel) -join ",")
            }
        DRIVER_ConnectFunctions = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty ConnectFunctions ))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty ConnectFunctions) -join ",")
            }
        DRIVER_CPTimeout = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty CPTimeout))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty CPTimeout) -join ",")
            }
        DRIVER_DEBUG = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty DEBUG))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty DEBUG) -join ",")
            }
        DRIVER_Driver = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty Driver))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty Driver) -join ",")
            }
        DRIVER_DriverODBCVer = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty DriverODBCVer))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty DriverODBCVer) -join ",")
            }
        DRIVER_FileExtns = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty FileExtns))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty FileExtns) -join ",")
            }
        DRIVER_FileUsage = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty FileUsage))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty FileUsage) -join ",")
            }
        DRIVER_Locale_Decimal = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty "Local Decimal"))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty "Local Decimal") -join ",")
            }
        DRIVER_Setup = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty Setup))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty Setup) -join ",")
            }
        DRIVER_SQLLevel = if([string]::IsNullOrWhiteSpace(@($x64 | select -ExpandProperty SQLLevel))){
                $null
            }
            else{
                (@($x64 | select -ExpandProperty SQLLevel) -join ",")
            }
    } #end of if
} #end of pscustom
    else{
        $x86 = Get-OdbcDriver $o.name | ?{$_.platform -eq "32-bit"} | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json

[pscustomobject]@{
        ODBC_driverNAME = $o.name
        ODBC_platform = $o.platform # -join ","
    #ATTRIBUTE
        DRIVER_APILevel = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty apilevel))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty apilevel) -join ",")
            }
        DRIVER_ConnectFunctions = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty ConnectFunctions))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty ConnectFunctions))
            }
        DRIVER_CPTimeout = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty CPTimeout))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty CPTimeout))
            }
        DRIVER_DEBUG = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty DEBUG))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty DEBUG))
            }
        DRIVER_Driver = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty Driver))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty Driver))
            }
        DRIVER_DriverODBCVer = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty DriverODBCVer))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty DriverODBCVer))
            }
        DRIVER_FileExtns = if([string]::IsNullOrWhiteSpace(@(Get-OdbcDriver $o.name | select -ExpandProperty attribute | ConvertTo-Json -Compress | ConvertFrom-Json | select -ExpandProperty syncroot | select -ExpandProperty FileExtns))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty FileExtns))
            }
        DRIVER_FileUsage = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty FileUsage))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty FileUsage))
            }
        DRIVER_Locale_Decimal = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty "Locale Decimal"))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty "Locale Decimal"))
            }
        DRIVER_Setup = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty Setup))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty Setup))
            }
        DRIVER_SQLLevel = if([string]::IsNullOrWhiteSpace(@($x86| select -ExpandProperty SQLLevel))){
            $null
            }
            else{
                (@($x86 | select -ExpandProperty SQLLevel))
            }
    } #end of else
    } #end of pscustom
} #end of pscustom


#BREAK




#see above $ODBCdsn_ATTRIBUTE section reporting options and replace/update variable/HT name if ever needed

#i may rename the previous report from the alpha/omaga to a/z for easier identifying deltas in the future...
#PREVIOUS REPORT SECTION
$1 = gc $report_dir\1_DSN_Information.txt
$1_ts = gci $report_dir\1_DSN_Information.txt | select lastwritetime
if($1 -ne $null){
    $report_HT += @{
        previous_reportTimeStamp = $($1_ts.LastWriteTime.ToString("MM/dd/yyyy HH:mm tt"))
    }
    $report_b4 = $null
    $report_b4 = $1[20..$($1.Length)]
    $report_b4 | Out-File $report_dir\_DSN_Alpha.txt -Force
}
else{
    $report_HT += @{
        previous_reportTimeStamp = "no indication of previous report or timestamp is null"
    }
    New-Item -ItemType File $report_dir\1_DSN_Information.txt -Force
}

New-Item -ItemType File $report_dir\1_DSN_Information.txt -Force
$cn | Out-File $report_dir\1_DSN_Information.txt -Append
Write-Output "'cdw-getDSN.ps1' script executed on: `t $($cn)" | Out-File $report_dir\1_DSN_Information.txt -Append
$date | Out-File $report_dir\1_DSN_Information.txt -Append
$report_HT = $report_HT.GetEnumerator() | sort name
$report_HT | select name,value | Out-File $report_dir\1_DSN_Information.txt -Append
sleep -Milliseconds 50

Write-Output "`nData Source Name (DSN) Information (Name, Platform, DriverName, Type, and Attribute):`n" | Out-File $report_dir\1_DSN_Information.txt -Append
$ODBCdsn_ATTRIBUTE.GetEnumerator() | Out-File $report_dir\1_DSN_Information.txt -Append
Write-Output "`nODBC Driver Information (Name, Platform, Attribute):" | Out-File $report_dir\1_DSN_Information.txt -Append
$ODBCdriver_value.GetEnumerator() | Out-File $report_dir\1_DSN_Information.txt -Append
Write-Output "`nend of report.." | Out-File $report_dir\1_DSN_Information.txt -Append

$2_ts = gci $report_dir\1_DSN_Information.txt
$2 = $null
$2 = gc $report_dir\1_DSN_Information.txt
$2[20..$($2.Length)] | Out-File $report_dir\_DSN_Omega.txt -Force
$1_previous = $null
$1_previous = gc $report_dir\_DSN_Alpha.txt
$2_new = $null
$2_new = gc $report_dir\_DSN_Omega.txt
$delta = $null
$delta = diff $1_previous $2_new
if($delta -ne $null){ #if not empty then there are deltas!
    Write-Warning "CHANGES DETECTED! Verify changes are planned and contact appropriate personnel as-needed."
    New-Item -ItemType File $report_dir\DSN_Delta_Detected.txt -Force
    Write-Output "$($date)" | Out-File $report_dir\DSN_Delta_Detected.txt -Append
    Write-Output "$($cn)" | Out-File $report_dir\DSN_Delta_Detected.txt -Append
    Write-Output "Report loc: `t$($report_dir)`n" | Out-File $report_dir\DSN_Delta_Detected.txt -Append
    Write-Output "CHANGE DETECTED BETWEEN DSN INFORMATION! Verify DSN settings on $($cn)!" | Out-File $report_dir\DSN_Delta_Detected.txt -Append
    $delta.GetEnumerator() | select @{n="DIFF!";e={$_.inputobject}},@{n="File";e={$_.sideindicator -replace ("=>","1_DSN_report @ $($1_ts.LastWriteTime.ToString("HH" + "mm" + "__yyyy_MM_dd"))") -replace ("<=","Today's DSN report @ $($date.ToString("HH" + "mm" + "__yyyy_MM_dd"))")}} | ft -AutoSize -Wrap | Out-File $report_dir\DSN_Delta_Detected.txt -Append
#moved to below so alert email can be sent:    mi $report_dir\DSN_Delta_Detected.txt ("$($cn)__" + "DsnChangeDETECTED__" + "$($1_ts.LastWriteTime.ToString("HH" + "mm" + "__yyyy_MM_dd"))" + ".txt") -Force
    sleep -Milliseconds 50
    mi $report_dir\_DSN_Alpha.txt ("_DSN_AlphaTimeStamp__" + $($1_ts.LastWriteTime.ToString("HH" + "mm" + "__yyyy_MM_dd")) + ".txt") -force -Verbose
    sleep -Milliseconds 50
    mi $report_dir\_DSN_Omega.txt ("_DSN_OmegaTimeStamp__" + "$($report_HT.GetEnumerator() | ?{$_.key -eq "date_timestamp"} | select value | select @{n="TS";e={$_.value.ToString("HH" + "mm" + "__yyyy_MM_dd")}} | select -ExpandProperty TS)" + ".txt") -force -Verbose
    sleep -Milliseconds 50
    $DISPLAY_RESULT = $true
    $report_HT += @{
#!ADD report location to HT!!!!
        scriptREPORT_RESULT = "CHANGES DETECTED - view report for complete details listed in the above verbose output"
    }
}
else{
    $DISPLAY_RESULT = $false
    $report_HT += @{
        scriptREPORT_RESULT = "No changes detected between today's DSN report and report stored on $($cn)!"
    }
    sleep -Milliseconds 50
    ri $report_dir\_DSN_Alpha.txt
    ri $report_dir\_DSN_Omega.txt
}
sleep -Milliseconds 200

if($DISPLAY_RESULT -eq $false){
    Write-output "No changes detected between today's DSN report and file stored on $($cn)!"
}
else{
    ii .
    $report_HT | ft -AutoSize
    Write-Output "`nCHANGES DETECTED - view report for complete details listed in the above verbose output"
    Send-MailMessage -To "Jeff.Giese@va.gov" -From "DoNotReply$($cn)@va.gov" -SMTPServer "smtp.va.gov" -Subject "DSN Change from: '$($cn)' @ $HourMinDATE" -Body "See attached file for details:" -Attachments $report_dir\DSN_Delta_Detected.txt -Priority High
    mi $report_dir\DSN_Delta_Detected.txt ("$($cn)__" + "DsnChangeDETECTED__" + "$($1_ts.LastWriteTime.ToString("HH" + "mm" + "__yyyy_MM_dd"))" + ".txt") -Force
}
sl $dir_b4
$ErrorActionPreference = $ea_b4
}
cdw-getDSN