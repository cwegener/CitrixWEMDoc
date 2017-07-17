<#

.SYNOPSIS
This is a simple Powershell script to document Citrix Workspace Environment Manager Configuration

.DESCRIPTION
The Script uses a selection of SQL queries to Query the WEM databases, it then outputs relevant information to a simple HTML file with limited formatting
The Script looks for either the legacy (SSMO) Style powershell objects, and ideally should be run from a server with SQL SSMS Installed. 
If not found, it will detect and offer to install the new SQL powershell modules via WMF 5 functionality (Get-Module)
If you Decline. Script will exit
If you have no internet. Script will exit
If there is no SQL PoSH Objects, Script will exit
If you choose to not use SQL Auth, your currently logged on account will require permissions on the SQL Server and WEM Database

.PARAMETER DBServer
The SQL Server hosting WEM. If on a dedicated instance, Specify the "Instance Name"

.PARAMETER Database
The WEM Database Name

.PARAMETER Outfile
The Output file for the report

.EXAMPLE 
The following Documents WEM with the SQL Database installed on the Server Kindo-DDC with the Name CitrixWEM and Outputs the HTML file to C:\temp\DocumentWEM.html
.\DocumentWEM.ps1 -DBServer "Kindo-DDC" -Database "CitrixWEM" -Outfile "c:\temp\DocumentWEM.html"

.EXAMPLE
The following Documents WEM with the SQL Database installed on the Server Kindo-DDC within an instance called SQLEXPRESS with the Name CitrixWEM and Outputs the HTML file to C:\temp\DocumentWEM.html
.\DocumentWEM.ps1 -DBServer "Kindo-DDC\SQLEXPRESS" -Database "CitrixWEM" -Outfile "c:\temp\DocumentWEM.html"

.NOTES
This is purely a configuration dump, output can be dealt with as you see fit

This has been confirmed against WEM 4.3
WEM 4.2 Works, however this was coded against 4.3 DB Schema
This is supported on Windows Server 2012 R2 and Windows Server 2016 Only (required Windows Management Framework 5 to remediate missing SQL objects) 
It will run on Windows 10 as long as the pre-reqs exist (SQL Server Management Objects), however it will not attempt to remediate 
This has been tested against PowerShell V5 Only

V0.2
Fixed hardcoded SQL DB Name and swapped to variable - (Thanks George Spiers)
Split functions out for OS Detection, Framework Detection and Remediation, SQL PosH Modules Install, Internet Test

V0.3
Added Integrated Windows Auth Capability (Thanks George Spiers)
Added input Validation check

.LINK

#>

Param(
    [Parameter(Mandatory=$true)][string]$DBServer,
    [Parameter(Mandatory=$true)][string]$Database,
    [Parameter(Mandatory=$true)][string]$Outfile
)
#region Variables and Startup

$ErrorActionPreference = "Stop"

# Force HTML Outfile 
if ($Outfile -notlike "*.html") {
    Write-Warning "Output file must be HTML. Existing input $Outfile has been renamed to $Outfile.html"
    $OutFile = $Outfile + ".html"
}

# Test Outfile existence
if (Test-Path $Outfile) {
    Write-Host "Removing existing $Outfile" -ForegroundColor Green
    Remove-item $Outfile -Force
}
elseif (!(Test-Path $Outfile)) {
    New-Item $Outfile -ItemType File -Force
    Write-Host "Created $Outfile" -ForegroundColor Green
}

# Create install logs Directory
$Install_Logs = "$env:SystemDrive\InstallLogs"
if(!(Test-Path $Install_Logs)) {
    New-Item -Path "$env:SystemDrive\InstallLogs\" -ItemType Directory -Force | Out-Null
    Write-Host "Created InstallLogs Directory" -ForegroundColor Green
}

$LogTime = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
$LogFile = "$Install_Logs\Install_Log_Applications_$env:ComputerName $LogTime.log"

#endregion

#region Sub Functions
# Define Authentication Type Function 
function SelectAuthType {

    $SQLAuth = Read-Host "Do you wish to use SQL Authentication? Y, N or Q (Quit)"
    while("Y","N","Q" -notcontains $SQLAuth) {
        $SQLAuth = Read-Host "Enter Y, N or Q (Quit)"
    }
    
    if ($SQLAuth -eq "Y") {
        Write-Host "Script will utilise SQL Authentication to connect to $DBServer" -ForegroundColor Green
        $Credential = Get-Credential
        $Username = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password

        TestLegacySQLPoSHModules
    }
    elseif ($SQLAuth -eq "N") { 
        Write-Host "Script will utilise Windows Authentication to connect to $DBServer" -ForegroundColor Green
        $ConfirmWindowsAuth = Read-Host "Please ensure your current account has relevant permissions to execute SQL Queries against $Database on $DBServer.`nProceed? Y/N"
        while("Y","N" -notcontains $ConfirmWindowsAuth) {
            $ConfirmWindowsAuth = Read-Host "Enter Y or N"
        }
        if ($ConfirmWindowsAuth -eq "Y") {
            Write-Host "Script will utilise Windows Authentication to connect to $DBServer" -ForegroundColor Green
            TestLegacySQLPoSHModules
        }
        elseif ($ConfirmWindowsAuth -ne "Y") {
            SelectAuthType
        }
    }
    elseif ($SQLAuth -eq "Q"){
        exit
    }
}

# Define basic Log Write Function 
function Log-Write {
    param ([string]$logstring)
    add-content $logfile -value $logstring
}

# Define Download and Install Function    
function DownloadAndInstall {
    Write-Host "Downloading and Installing $ComponentName" -ForegroundColor Yellow 
    Log-Write "$(Get-Date -f o) - Downloading and Installing $ComponentName"

    if (!(Test-path $env:temp\$(Split-Path $Source_Url -Leaf))) {
        $Download = Start-BitsTransfer -Source $Source_Url -Description $ComponentName -Destination $env:temp -DisplayName $ComponentName   
    }
    
    Start-Process -FilePath $env:temp\$(Split-Path $Source_Url -Leaf) -ArgumentList $ArgumentList -Wait -PassThru
    Write-Host "Finished Installing $ComponentName" -ForegroundColor Green
    Log-Write "$(Get-Date -f o) - Finished Installing $ComponentName"

    Remove-Item -Path $env:temp\$(Split-Path $Source_Url -Leaf) -Force

}

# Define Internet Test Function
function TestInternetConnection {
    param([string]$Server, [int]$Port, [int]$Timeout=100)
    $Connection = New-Object System.Net.Sockets.TCPClient
        try {
            $Connection.Connect($Server,$Port)
            return $true
            }
        catch {
            return $false
            }
    }

# Define Document Heading Write Function 
function WriteHeading {
    ConvertTo-HTML -Head $b -body "<H2><br/>$Heading</H2>" | Out-File -append $OutFile -Force 
}

# Define Info Gathering Function 
function GetInfo {
        Write-Host "Attempting to Export $Text" -ForegroundColor Green
        try { if ($SQLAuth -eq "Y") {
            (Invoke-Sqlcmd -ServerInstance $DBServer -Username $Username -Password $Password -Database $DataBase -query $Query) `
            | Select-Object $Selection | ConvertTo-HTML -head $a -body "<H3>$Text</H3>" | Out-File -append $OutFile -Force
        }
              elseif (!($SQLAuth -eq "Y")) {
                (Invoke-Sqlcmd -ServerInstance $DBServer -Database $DataBase -query $Query) `
                | Select-Object $Selection | ConvertTo-HTML -head $a -body "<H3>$Text</H3>" | Out-File -append $OutFile -Force
            }
    }
        catch {
            [string]$ErrorText = $Error[0].CategoryInfo.Reason
            Write-Warning $ErrorText
        }
}

# Define Pre SQL 2017 SSMS PowerShell Modules Check
function TestLegacySQLPoSHModules{
	$SQLPowerShellModulesExist = (Get-Module -ListAvailable -Name Sqlps)
        if ($SQLPowerShellModulesExist.Name -eq "SQLPS") {
            Write-Host "SQL PowerShell Modules are Installed" -ForegroundColor Green 
            RunWemDoc
              } 
        else {
            "Legacy SQL PowerShell Modules Not Installed - Testing for Modern Modules"
            TestModernSQLPoSHModules
        } 
}

# Define SQL 2017 SSMS Powershell Modules Check
function TestModernSQLPoSHModules  {
    $SQLModuleExists = (Get-Module SqlServer)
	    if ($SQLModuleExists) {
	       Write-Host "Modern SQL Module Installed" -ForegroundColor Green
         RunWemDoc
	    } 
      elseif (!$SQLModuleExists){
            Write-Warning "Modern SQL Module Not Installed"
            CheckOSVersion
      }
}

# Define Operating System Check Function
function CheckOSVersion {
    write-Host "Checking Operating System Version" -ForegroundColor Green
    $OSVersion = (gwmi win32_operatingsystem).Caption

    if ($OSVersion -like "Microsoft Windows Server 2012 R2*" -Or $OSVersion -like "Microsoft Windows Server 2016*") {
        Write-Host "$OSVersion Detected" -ForegroundColor Green
    }
    else {
        Write-Warning "$OSVersion Detected. This Script is only Supported on Windows Server 2012 R2 and Windows Server 2016"
        Exit
    }
    if ($OSVersion -notlike "Microsoft Windows Server 2016*") {
        Write-Host "$OSVersion Detected. Checking Windows Management Framework Version Installed on $env:ComputerName"  -ForegroundColor Green
        CheckAndInstallWMF5 #If OS not Server 2016, Check WMF Version 
    }
    if ($OSVersion -like "Microsoft Windows Server 2016*") {
        DownloadSQLPoshMod #if OS is 16, WMF 5.0 is installed, install SQL PoSH
    }
}

# Define Check and Installation of WMF 5 Function 
function CheckAndInstallWMF5 {
    $WMFVerExist = $PSVersionTable.PSVersion.Major
    $WMFVerReq = "Windows Management Framework 5"
    $ComponentName = $WMFVerReq
    $Source_Url =  "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu"
    $ArgumentList = "/quiet /norestart"

    if ($WMFVerExist -eq "5") {
        Write-Host "$WMFVerReq Detected." -ForegroundColor Green
        DownloadSQLPoshMod # If framework 5 detected, run SQL PoSH Download Function
    }
    elseif ($WMFVerExist -ne "5") {
        Write-Warning "$WMFVerReq is required, however Windows Managment Framework $WMFVerExist is currently Installed on $env:ComputerName."
            $DownloadFramework5 = Read-Host "Would you like to download and install $WMFVerReq on $env:ComputerName? (Y/N)"
                while("Y","N" -notcontains $DownloadFramework5) {
                    $DownloadFramework5 = Read-Host "Enter Y or N"
                }
                if ($DownloadFramework5 -eq "Y") {
                        $ConfirmRebootRequirement = Read-Host "A reboot is required to complete Installation of $WMFVerReq. Do you wish to proceed? (Y/N)"
                            while("Y","N" -notcontains $ConfirmRebootRequirement) {
                                $ConfirmRebootRequirement = Read-Host "Enter Y or N"
                            }
                            if ($ConfirmRebootRequirement -eq "Y") {
                                Write-Host "Testing Internet Connection on $env:ComputerName" -ForegroundColor Green
                                $InternetAccessible = TestInternetConnection -Server "www.google.com" -Port "80" 
                                    if ($InternetAccessible) {
                                        Write-Host "Internet Connection Successful!" -ForegroundColor Green
                                        $WUAUInitState = (Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='wuauserv'").StartMode
                                            if ($WUAUInitState -eq "Disabled") {
                                                Write-Warning "Windows Update Service Startup set to $WUAUInitState. WMF 5.0 Install Requires the use of Windows Update Service. Setting Service Startup to Manual"
                                                Set-Service -Name "wuauserv" -StartupType Manual
                                            } elseif ($WUAUInitState -eq "Automatic" -or $WUAUInitState -eq "Manual") {
                                                Write-Host "Windows Update Service Startup Type OK. Proceeding to Install $WMFVerReq" -ForegroundColor Green 
                                            }
                                        DownloadAndInstall
                                        Write-Host "$WMFVerReq has been installed on $Env:ComputerName. Please restart Server to apply changes" -ForegroundColor Green
                                        Exit
                                    }
                                    elseif (!$InternetAccessible) {
                                        Write-Warning "Internet Connection Test Failed on $env:ComputerName. Please resolve connection issues or manually install $WMFVerReq"
                                        Exit
                                    }
                            }
                            elseif ($ConfirmRebootRequirement -ne "Y") {
                                Write-Warning "$WMFVerReq is required. Please allow installation or install manually before running this Script"
                                Exit
                            }
                }
                elseif ($DownloadFramework5 -ne "Y") {
                    Write-Warning "$WMFVerReq is required to install the relevant SQL PoSH Modules"
                    exit
                    }
                }
    }

# Define Install SQL PowerShell Modules
function DownloadSQLPoshMod {
    $GetSQLPoSHModule = Read-Host "SQL Module Not Installed, Would you like to Install now? (Y/N)"
        while("Y","N" -notcontains $GetSQLPoSHModule) {
            $GetSQLPoSHModule = Read-Host "Enter Y or N"
        }
        if ($GetSQLPoSHModule -eq "Y") {
            Write-Host "Happy to Oblige - Installing Module SqlServer" -ForegroundColor Green
            try {
                Install-module -Name SqlServer -Scope AllUsers -Force -AllowClobber
                Write-Host "Installed SqlServer Module" -ForegroundColor Green
                Import-Module SqlServer
                Write-Host "Imported SqlServer Module" -ForegroundColor Green
                RunWemDoc #Run WEM Documentation
            }
            catch {
                [string]$ErrorText = $Error[0].CategoryInfo.Reason
                Write-Warning $ErrorText                
            }
        } 
        elseif ($GetSQLPoSHModule -ne "Y") {
            Write-Warning "You cannot run this Script without the SQL Powershell Modules Installed"
            Exit
        }
}

#endregion

#region Primary WEM Function

# Define Main Function of WEM Documentation
function RunWemDoc {

#region styles
## \\ Table HTML Style
$a=@'
<style>
    body{
         background-color:white;
    }
    table{
         border-width: 1px;
         border-style: solid;
         solid;border-color: black;
         border-collapse: collapse;
         background-color: white;
         color: black;
    }
    th{
    	border-width: 1px;
    	padding: 5px;
    	border-style: solid;
    	border-color: black;
        border-collapse: collapse;
    	background-color: LIGHTGRAY;
        font-family: Verdana;
        font-size: 8pt;
        color: black;
    }
    td{
    	border-width: 1px;
    	padding: 5px;
    	border-style: solid;
    	border-color: black;
        border-collapse: collapse;
    	background-color: white;
        font-family: Verdana;
        font-size: 8pt;
        color: black;
    }
</style>
'@

## \\ Heading HTML Style
$b=@'
<style>
    body{
         font-family: Verdana;
         font-size: 10pt;
         color: black;
</style>
'@

#endregion

#region Write HTML Header

#Write Title to HTML Output
$Date = Get-Date -Format "dd-MM-yyyy"
ConvertTo-HTML -Head $b -body "<H2>Citrix Workspace Environment Manager Configuration Report $Date</H2>" | Out-File -append $OutFile

#endregion

#region Global WEM Settings

# Add Heading to HTML Output
$Heading = "Global WEM Details"
WriteHeading

# WEM Site Details
$Text = "WEM Site (Configuration Set) Information"
$WEMSiteList_Query = "SELECT [IdSite]
      ,[Name]
      ,[Description]
      ,[State]
  FROM [$Database].[dbo].[VUEMSites]"

$Query = $WEMSiteList_Query
$Selection = "Name","Description","IdSite","State"
GetInfo

#endregion

#region WEM Actions

# Add Heading to HTML Output
$Heading = "WEM Actions"
WriteHeading

# Application Actions
$Text = "Defined Applications Actions List"
$WEMActionApps_Query = "SELECT [IdApplication]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[AppType]
      ,[ActionType]
      ,[DisplayName]
      ,[StartMenuTarget]
      ,[TargetPath]
      ,[Parameters]
      ,[WorkingDirectory]
      ,[IconLocation]
  FROM [$Database].[dbo].[VUEMApps]"

$Query = $WEMActionApps_Query
$Selection = "Name","Description","DisplayName","StartMenuTarget","TargetPath","Parameters","WorkingDirectory","State","IdSite"
GetInfo

# Printer Actions
$Text = "Defined Printers Actions List"
$WEMActionPrinters_Query = "SELECT [IdPrinter]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[DisplayName]
      ,[State]
      ,[ActionType]
      ,[TargetPath]
      ,[UseExtCredentials]
      ,[ExtLogin]
  FROM [$Database].[dbo].[VUEMPrinters]"

$Query = $WEMActionPrinters_Query
$Selection = "Name","Description","DisplayName","TargetPath","UseExtCredentials","ExtLogin","State","IdSite"
GetInfo

# Network Drive Actions
$Text = "Defined Network Drive Actions List"
$WEMActionNetworkDrives_Query = "SELECT [IdNetDrive]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[DisplayName]
      ,[State]
      ,[ActionType]
      ,[TargetPath]
      ,[UseExtCredentials]
      ,[ExtLogin]
  FROM [$Database].[dbo].[VUEMNetDrives]"

$Query = $WEMActionNetworkDrives_Query
$Selection = "Name","Description","DisplayName","TargetPath","UseExtCredentials","ExtLogin","State","IdSite"
GetInfo

# Virtual Drives Actions
$Text = "Defined Virtual Drive Actions List"
$WEMActionVirtualDrives_Query = "SELECT [IdVirtualDrive]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[TargetPath]
  FROM [$Database].[dbo].[VUEMVirtualDrives]"

$Query = $WEMActionVirtualDrives_Query
$Selection = "Name","Description","TargetPath","State","IdSite"
GetInfo

# Registry Actions
$Text = "Defined Registry Actions List"
$WEMActionRegistry_Query = "SELECT [IdRegValue]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[TargetRoot]
      ,[TargetName]
      ,[TargetPath]
      ,[TargetType]
      ,[TargetValue]
      ,[RunOnce]
  FROM [$Database].[dbo].[VUEMRegValues]"

$Query = $WEMActionRegistry_Query
$Selection = "Name","Description","TargetPath","TargetName","TargetType","TargetValue","RunOnce","State","IdSite"
GetInfo

# External Tasks Actions
$Text = "Defined External Tasks Actions List"
$WEMActionExtTasks_Query = "SELECT [IdExtTask]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[TargetPath]
      ,[TargetArgs]
      ,[RunHidden]
      ,[WaitforFinish]
      ,[TimeOut]
      ,[ExecOrder]
      ,[RunOnce]
  FROM [$Database].[dbo].[VUEMExtTasks]"

$Query = $WEMActionExtTasks_Query
$Selection = "Name","Description","TargetPath","TargetArgs","RunHidden","WaitforFinish","TimeOut","ExecOrder","RunOnce","State","IdSite"
GetInfo

# Environment Variable Actions
$Text = "Defined Environment Variables Actions List"
$WEMActionEnvVariables_Query = "SELECT [IdEnvVariable]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[VariableName]
      ,[VariableValue]
      ,[VariableType]  
  FROM [$Database].[dbo].[VUEMEnvVariables]"

$Query = $WEMActionEnvVariables_Query
$Selection = "Name","Description","VariableName","VariableValue","VariableType","State","IdSite"
GetInfo

# Port Actions
$Text = "Defined Port Actions List"
$WEMActionPorts_Query = "SELECT [IdPort]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[PortName]
      ,[TargetPath]      
  FROM [$Database].[dbo].[VUEMPorts]"

$Query = $WEMActionPorts_Query
$Selection = "Name","Description","PortName","TargetPath","State","IdSite"
GetInfo

# Ini File Actions
$Text = "Defined INI File Actions List"
$WEMActionINIFiles_Query = "SELECT [IdIniFileOp]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[TargetPath]
      ,[TargetSectionName]
      ,[TargetValueName]
      ,[TargetValue]
      ,[RunOnce]
  FROM [$Database].[dbo].[VUEMIniFilesOps]"

$Query = $WEMActionINIFiles_Query
$Selection = "Name","Description","TargetPath","TargetSectionName","TargetValueName","TargetValue","State","IdSite"
GetInfo

# File System Actions
$Text = "Defined File System Actions List"
$WEMActionFileSystemOps_Query = "SELECT [IdFileSystemOp]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[SourcePath]
      ,[TargetPath]
      ,[TargetOverwrite]
      ,[RunOnce]   
  FROM [$Database].[dbo].[VUEMFileSystemOps]"

$Query = $WEMActionFileSystemOps_Query
$Selection = "Name","Description","SourcePath","TargetPath","TargetOverWrite","RunOnce","State","IdSite"
GetInfo

# User DSN Actions
$Text = "Defined User DSN Actions List"
$WEMActionUserDSN_Query = "SELECT [IdUserDSN]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[TargetName]
      ,[TargetDriverName]
      ,[TargetServerName]
      ,[TargetDatabaseName]
      ,[UseExtCredentials]
      ,[ExtLogin]
      ,[ExtPassword]
      ,[RunOnce]
  FROM [$Database].[dbo].[VUEMUserDSNs]"

$Query = $WEMActionUserDSN_Query
$Selection = "Name","Description","TargetName","TargetDriverName","TargetServerName","TargetDatabaseName","UseExtCredentials","ExtLogin","RunOnce","State","IdSite"
GetInfo

# File Association Actions
$Text = "Defined File Type Association Actions List"
$WEMActionFTA_Query = "SELECT [IdFileAssoc]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[ActionType]
      ,[FileExt]
      ,[ProgId]
      ,[Action]
      ,[isDefault]
      ,[TargetPath]
      ,[TargetCommand]
      ,[TargetOverWrite]
      ,[RunOnce]
  FROM [$Database].[dbo].[VUEMFileAssocs]"

$Query = $WEMActionFTA_Query
$Selection = "Name","Description","FileExt","ProgId","Action","isDefault","TargetPath","TargetCommand","TargetOverwrite","RunOnce","State","IdSite"
GetInfo

#endregion

#region Filters and Conditions

# Add Heading to HTML Output
$Heading = "WEM Filters and Conditions"
WriteHeading

#Filters List
$Text = "Defined Filters Detail"
$WEMFilters_Query = "SELECT [IdFilterRule]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[Conditions]
  FROM [$Database].[dbo].[VUEMFiltersRules]"

$Query = $WEMFilters_Query
$Selection = "Name","Description","Conditions","State","IdSite"
GetInfo

# Conditions List
$Text = "Defined Conditions Detail"
$WEMConditions_Query = "SELECT [IdFilterCondition]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[Type]
      ,[TestValue]
      ,[TestResult]    
  FROM [$Database].[dbo].[VUEMFiltersConditions]"

$Query = $WEMConditions_Query
$Selection = "Name","Description","TestValue","TestResult","State","IdFilterCondition","IdSite"
GetInfo

#endregion

#region Active Directory Groups

# Add Heading to HTML Output
$Heading = "Active Directory Objects"
WriteHeading

# Active Directory Objects
$Text = "Defined Active Directory Object Detail"
$WEMADObjects_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Description]
      ,[State]
      ,[Type]
      ,[Priority]
  FROM [$Database].[dbo].[VUEMItems]"

$Query = $WEMADObjects_Query
$Selection = "Name","Description","Type","Priority","State","IdSite","IdItem"
GetInfo

#endregion

#region Action Assignments

# Add Heading to HTML Output
$Heading = "WEM Action Assignments"
WriteHeading

# Application Assignment
$Text = "Assigned Applications"
$WEMAssignmentApplications_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMApps.DisplayName as AppDisplayName
    , dbo.VUEMApps.TargetPath as AppTargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMApps, dbo.VUEMAssignedApps, dbo.VUEMFiltersRules
  WHERE dbo.VUEMApps.IdApplication = dbo.VUEMAssignedApps.IdAssignedApplication 
  AND dbo.VUEMAssignedApps.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedApps.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentApplications_Query
$Selection = "GroupSID","GroupDescription","AppDisplayName","AppTargetPath","RuleName","RuleDescription"
GetInfo

#Printer Assignment
$Text = "Assigned Printers"
$WEMAssignmentPrinters_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMPrinters.Name as PrinterName
    , dbo.VUEMPrinters.TargetPath as PrinterTargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMPrinters, dbo.VUEMAssignedPrinters, dbo.VUEMFiltersRules
  WHERE dbo.VUEMPrinters.IdPrinter = dbo.VUEMAssignedPrinters.IdPrinter
  AND dbo.VUEMAssignedPrinters.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedPrinters.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentPrinters_Query
$Selection = "GroupSID","GroupDescription","PrinterName","PrinterTargetPath","RuleName","RuleDescription"
GetInfo

# Network Drive Assignment
$Text = "Assigned Network Drives"
$WEMAssignmentNetworkDrives_Query  = "SELECT dbo.VUEMItems.Name as GroupSID
      ,dbo.VUEMItems.Description as GroupDescription
      ,dbo.VUEMNetDrives.DisplayName as DriveDisplayname
      ,dbo.VUEMNetDrives.TargetPath as DriveTargetPath
      ,dbo.VUEMAssignedNetDrives.DriveLetter
      ,dbo.VUEMFiltersRules.Name as RuleName
      ,dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMNetDrives, dbo.VUEMAssignedNetDrives, dbo.VUEMFiltersRules
  WHERE dbo.VUEMNetDrives.IdNetDrive = dbo.VUEMAssignedNetDrives.IdNetDrive 
  AND dbo.VUEMAssignedNetDrives.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedNetDrives.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentNetworkDrives_Query
$Selection = "GroupSID","GroupDescription","DriveDisplayName","DriveTargetPath","DriveLetter","RuleName","RuleDescription"
GetInfo

#Virtual Drive Assignment
$Text = "Assigned Virtual Drives"
$WEMAssignmentVirtualDrives_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMVirtualDrives.Name as VirtualDriveName
    , dbo.VUEMVirtualDrives.TargetPath as VirtualDriveTargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMVirtualDrives, dbo.VUEMAssignedVirtualDrives, dbo.VUEMFiltersRules
  WHERE dbo.VUEMVirtualDrives.IdVirtualDrive = dbo.VUEMAssignedVirtualDrives.IdVirtualDrive
  AND dbo.VUEMAssignedVirtualDrives.IdItem= dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedVirtualDrives.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentVirtualDrives_Query
$Selection = "GroupSID","GroupDescription","VirtualDriveName","VirtualDriveTargetPath","RuleName","RuleDescription"
GetInfo

# Registry Key Assignment
$Text = "Assigned Registry Values"
$WEMAssignmentRegistry_Query = "Select dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMRegValues.Name as RegValueName
    , dbo.VUEMRegValues.TargetPath as RegValueTargetPath
    , dbo.VUEMRegValues.TargetName
    , dbo.VUEMRegValues.TargetType
    , dbo.VUEMRegValues.TargetValue
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMRegValues, dbo.VUEMAssignedRegValues, dbo.VUEMFiltersRules
  WHERE dbo.VUEMRegValues.IdRegValue = dbo.VUEMAssignedRegValues.IdRegValue
  AND dbo.VUEMAssignedRegValues.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedRegValues.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentRegistry_Query
$Selection = "GroupSID","GroupDescription","RegValueName","RegValueTargetPath","TargetName","TargetType","TargetValue","RuleName","RuleDescription"
GetInfo

# Environment Variable Assignment
$Text = "Assigned Environment Variables"
$WEMAssignmentEnvVariables_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMEnvVariables.Name
    , dbo.VUEMEnvVariables.VariableName
    , dbo.VUEMEnvVariables.VariableValue
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMEnvVariables, dbo.VUEMAssignedEnvVariables, dbo.VUEMFiltersRules
  WHERE dbo.VUEMEnvVariables.IdEnvVariable = dbo.VUEMAssignedEnvVariables.IdEnvVariable
  AND dbo.VUEMAssignedEnvVariables.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedEnvVariables.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentEnvVariables_Query
$Selection = "GroupSID","GroupDescription","Name","VariableName","VariableValue","RuleName","RuleDescription"
GetInfo

# Port Assignment
$Text = "Assigned Ports"
$WEMAssignmentPort_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMPorts.Name
    , dbo.VUEMPorts.PortName
    , dbo.VUEMPorts.Description
    , dbo.VUEMPorts.TargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMPorts, dbo.VUEMAssignedPorts, dbo.VUEMFiltersRules
  WHERE dbo.VUEMPorts.IdPort = dbo.VUEMAssignedPorts.IdPort
  AND dbo.VUEMAssignedPorts.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedPorts.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentPort_Query
$Selection = "GroupSID","GroupDescription","Name","PortName","Description","TargetPath","RuleName","RuleDescription"
GetInfo

# INI Assignment
$Text = "Assigned Ini Files"
$WEMAssignmentIniOps_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMIniFilesOps.Name
    , dbo.VUEMIniFilesOps.Description
    , dbo.VUEMIniFilesOps.TargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMIniFilesOps, dbo.VUEMAssignedIniFilesOps, dbo.VUEMFiltersRules
  WHERE dbo.VUEMIniFilesOps.IdIniFileOp = dbo.VUEMAssignedIniFilesOps.IdIniFileOp
  AND dbo.VUEMAssignedIniFilesOps.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedIniFilesOps.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentIniOps_Query
$Selection = "GroupSID","GroupDescription","Name","Description","TargetPath","RuleName","RuleDescription"
GetInfo

# External Task Assignment
$Text = "Assigned External Tasks"
$WEMAssignmentExternalTask_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMExtTasks.Name
    , dbo.VUEMExtTasks.Description
    , dbo.VUEMExtTasks.TargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMExtTasks, dbo.VUEMAssignedExtTasks, dbo.VUEMFiltersRules
  WHERE dbo.VUEMExtTasks.IdExtTask = dbo.VUEMAssignedExtTasks.IdExtTask
  AND dbo.VUEMAssignedExtTasks.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedExtTasks.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentExternalTask_Query
$Selection = "GroupSID","GroupDescription","Name","Description","TargetPath","RuleName","RuleDescription"
GetInfo

# File System Assignments
$Text = "Assigned File System Operations"
$WEMAssignmentFileSystemOps_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMFileSystemOps.Name
    , dbo.VUEMFileSystemOps.Description
    , dbo.VUEMFileSystemOps.SourcePath
    , dbo.VUEMFileSystemOps.TargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMFileSystemOps, dbo.VUEMAssignedFileSystemOps, dbo.VUEMFiltersRules
  WHERE dbo.VUEMFileSystemOps.IdFileSystemOp = dbo.VUEMAssignedFileSystemOps.IdFileSystemOp
  AND dbo.VUEMAssignedFileSystemOps.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedFileSystemOps.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentFileSystemOps_Query
$Selection = "GroupSID","GroupDescription","Name","Description","SourcePath","TargetPath","RuleName","RuleDescription"
GetInfo

# User DSN Assignments
$Text = "Assigned User DSNs"
$WEMAssignmentUserDSNs_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMUserDSNs.Name
    , dbo.VUEMUserDSNs.Description
    , dbo.VUEMUserDSNs.TargetName
    , dbo.VUEMUserDSNs.TargetDriverName
    , dbo.VUEMUserDSNs.TargetServerName
    , dbo.VUEMUserDSNs.TargetDatabaseName
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMUserDSNs, dbo.VUEMAssignedUserDSNs, dbo.VUEMFiltersRules
  WHERE dbo.VUEMUserDSNs.IdUserDSN = dbo.VUEMAssignedUserDSNs.IdAssignedUserDSN
  AND dbo.VUEMAssignedUserDSNs.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedUserDSNs.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentUserDSNs_Query
$Selection = "GroupSID","GroupDescription","Name","Description","TargetName","TargetDriverName","TargetServerName","TargetDatabaseName","RuleName","RuleDescription"
GetInfo

# File Assoc Assignments
$Text = "Assigned File Type Associations"
$WEMAssignmentFTA_Query = "SELECT dbo.VUEMItems.Name as GroupSID
    , dbo.VUEMItems.Description as GroupDescription
    , dbo.VUEMFileAssocs.Name
    , dbo.VUEMFileAssocs.Description
    , dbo.VUEMFileAssocs.FileExt
    , dbo.VUEMFileAssocs.TargetPath
    , dbo.VUEMFiltersRules.Name as RuleName
    , dbo.VUEMFiltersRules.Description as RuleDescription
  FROM dbo.VUEMItems, dbo.VUEMFileAssocs, dbo.VUEMAssignedFileAssocs, dbo.VUEMFiltersRules
  WHERE dbo.VUEMFileAssocs.IdFileAssoc = dbo.VUEMAssignedFileAssocs.IdAssignedFileAssoc
  AND dbo.VUEMAssignedFileAssocs.IdItem = dbo.VUEMItems.IdItem
  AND dbo.VUEMAssignedFileAssocs.IdFilterRule = dbo.VUEMFiltersRules.IdFilterRule"

$Query = $WEMAssignmentFTA_Query
$Selection = "GroupSID","GroupDescription","Name","Description","FileExt","TargetPath","RuleName","RuleDescription"
GetInfo

#endregion

#region Environmental and Agent

# Add Heading to HTML Output
$Heading = "WEM Environmental and Agent Settings"
WriteHeading

# Agent Settings
$Text = "WEM Agent Settings Detail"
$WEMAgentSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Value]
      ,[State]    
  FROM [$Database].[dbo].[VUEMAgentSettings]"

$Query = $WEMAgentSettings_Query
$Selection = "Name","Value","State","IdSite"
GetInfo

# WEM Environment Parameters
$Text = "WEM Environment Parameters"
$WEMParameters_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Value]
      ,[State]
  FROM [$Database].[dbo].[VUEMParameters]"

$Query = $WEMParameters_Query
$Selection = "Name","Value","State","IdSite"
GetInfo

#endregion

#region Policies and Profile Management

# Add Heading to HTML Output
$Heading = "Policies and Profile Management"
WriteHeading

# WEM Environmental Settings
$Text = "WEM Environment Settings"
$WEMEnvironmentalSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Type]
      ,[Value]
      ,[State]
  FROM [$Database].[dbo].[VUEMEnvironmentalSettings]"

$Query = $WEMEnvironmentalSettings_Query
$Selection = "Name","Type","Value","State","IdSite"
GetInfo

# Microsoft USV Settings
$Text = "Microsoft User State Virtualisation Settings"
$WEMMicrosoftUSVSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Type]
      ,[Value]
      ,[State]
  FROM [$Database].[dbo].[VUEMUSVSettings]"

$Query = $WEMMicrosoftUSVSettings_Query
$Selection = "Name","Type","Value","State","IdSite"
GetInfo

# Citrix UPM Settings
$Text = "Citrix Profile Management Settings"
$WEMUPMSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Value]
      ,[State]    
  FROM [$Database].[dbo].[VUEMUPMSettings]"

$Query = $WEMUPMSettings_Query
$Selection = "Name","Value","State","IdSite"
GetInfo

# Vmware Persona Settings
$Text = "VMware Persona Settings"
$WEMVMwarePersonaSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Value]
      ,[State]    
  FROM [$Database].[dbo].[VUEMPersonaSettings]"

$Query = $WEMVMwarePersonaSettings_Query
$Selection = "Name","Value","State","IdItem"
GetInfo

#endregion

#region Transformer Settings

# Add Heading to HTML Output
$Heading = "Transformer Settings"
WriteHeading

#Kiosk Settings
$Text = "Transformer Kiosk Settings"
$WEMKioskSettings_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Type]
      ,[Value]
      ,[State]
  FROM [$Database].[dbo].[VUEMKioskSettings]"

$Query = $WEMKioskSettings_Query
$Selection = "Name","Type","Value","State","IdSite"
GetInfo

#endregion

#region System Optimisation Settings

# Add Heading to HTML Output
$Heading = "System Monitoring Settings"
WriteHeading

# WEM System Monitoring Settings
$Text = "System Monitoring Settings"
$WEMSystemMonitoring_Query = "SELECT [IdItem]
      ,[IdSite]
      ,[Name]
      ,[Value]
      ,[State] 
  FROM [$Database].[dbo].[VUEMSystemMonitoringSettings]"

$Query = $WEMSystemMonitoring_Query
$Selection = "Name","Value","State","IdSite"
GetInfo

# WEM System Utilities Settings
$Text = "System Utilities Settings"
$WEMSystemUtilies_Query = "SELECT [IdItem]
     ,[IdSite]
     ,[Name]
     ,[Type]
     ,[Value]
     ,[State]    
 FROM [$Database].[dbo].[VUEMSystemUtilities]"

$Query = $WEMSystemUtilies_Query
$selection = "Name","Type","Value","State","IdSite"
GetInfo
#endregion

Start $OutFile
powershell -command Import-Module sqlServer 

}

#endregion

SelectAuthType

