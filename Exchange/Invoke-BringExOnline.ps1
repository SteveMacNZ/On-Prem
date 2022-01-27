<#
.SYNOPSIS
  Brings selected exchange server Online out of maintenance
.DESCRIPTION
  Brings selected exchange server Online out of maintenance
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  Log file for transcription logging
  CSV files 
.NOTES
  Version:        0.1
  Author:         Steve McIntyre     
  Creation Date:  1/12/21
  Purpose/Change: Initial Script
.LINK
  None
.EXAMPLE
  .\Invoke-BringExOnline.ps1
  Brings selected exchange server Online out of maintenance  
#>

#requires -version 4
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  # Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
#$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins

# Initialize your variables
#Set-Variable 

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:File        = ''                                                            # File var for Get-FilePicker Function
$Script:ScriptName  = 'Invoke-BringExOnline'                                        # Script Name used in the Open Dialogue
$Script:dest        = $PSScriptRoot                                                 # Destination path - uncomment to use PS script root
#$Script:dest        = "$($env:ProgramData)\What\Path"                               # Destination Path - comment to use PS Script root
$Script:LogFile     = $Script:dest + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name

$Script:GUID        = 'bfb7ca46-a172-436f-bf7c-1de51dfe4a71'                        # Script GUID
    #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
$Script:Version     = '0.1'                                                         # Script Version Number
$Script:Client      = 'MOJ'                                                         # Set Client Name - Used in Registry Operations
$Script:Operation   = 'Install'                                                     # Operations Feild for registry
$Script:Source      = 'Script'                                                      # Source (Script / MSI / Scheduled Task etc)
$Script:PackageName = $Script:ScriptName                                            # Packaged Name - Used in Registry Operations (may be same as script name)
$Script:RegPath     = "HKLM:\Software\$Script:Client\$Script:Source\$Script:PackageName\$Script:Operation"   # Registry Hive Location for Registry Operations
$Script:Desc        = "Bring Exchange Server Online from maintenance for $Script:Client"

$script:Domain      = ".corp.justice.govt.nz"
$Script:ShowOnTop   = 4096                                                                      # Set Dialogue Box to always show on top
$script:value       = ''                                                                        # Init variable for popup return value

#-----------------------------------------------------------[Hash Tables]-----------------------------------------------------------
#* Hash table for Write-Host Errors to be used as spatted
$cerror = @{ForeGroundColor = "Red"; BackgroundColor = "white"}
#* Hash table for Write-Host Warnings to be used as spatted
$cwarning = @{ForeGroundColor = "Magenta"; BackgroundColor = "white"}
#* Hash table for Write-Host highlighted to be used as spatted
$chighlight = @{ForeGroundColor = "Blue"; BackgroundColor = "white"}

# hash table for dialogue buttons
$buttons = @{
  OK               = 0
  OkCancel         = 1  
  AbortRetryIgnore = 2
  YesNoCancel      = 3
  YesNo            = 4
  RetryCancel      = 5
}
$bstyle = @{
  O = "OK"
  OC = "OkCancel"           
  ARI = "AbortRetryIgnore"
  YNC = "YesNoCancel"      
  YN = "YesNo"            
  RC = "RetryCancel"      
}

# hash table for dialogue icons  
<#$icons = @{
  Stop        = 16
  Question    = 32
  Exclamation = 48
  Information = 64
}#>
$itype = @{
  S = "Stop"        
  Q = "Question"   
  E = "Exclamation"
  I = "Information"
}

# hash table for dialogue button clicked
$clickedButton = @{
  -1 = 'Timeout'
  1  = 'OK'
  2  = 'Cancel'
  3  = 'Abort'
  4  = 'Retry'
  5  = 'Ignore'
  6  = 'Yes'
  7  = 'No'
}
#-----------------------------------------------------------[Functions]------------------------------------------------------------
#& Start Transcriptions
Function Start-Logging{

  try {
      Stop-Transcript | Out-Null
  } catch [System.InvalidOperationException] { }                                     # jobs are running
  $ErrorActionPreference = "Continue"                                                # Set Error Action Handling
  Get-Now                                                                            # Get current date time
  Start-Transcript -path $Script:LogFile -IncludeInvocationHeader -Append            # Start Transcription append if log exists
  Write-Host ''                                                                      # write Line spacer into Transcription file
  Write-Host ''                                                                      # write Line spacer into Transcription file
  Write-Host "================================================================================" 
  Write-Host "================== $Script:Now Processing Started ====================" 
  Write-Host "================================================================================"  
  Write-Host ''
  
  Write-Host ''                                                                       # write Line spacer into Transcription file
}

#& Display ScriptInfo
Function Get-ScriptInfo{
  
  Write-Host "#==============================================================================="
  Write-Host "# Name:             $Script:ScriptName"
  Write-Host "# Version:          $Script:Version"
  Write-Host "# GUID:             $Script:GUID"
  Write-Host "# Description:"
  Write-Host "# $Script:Desc"
  Write-Host "#-------------------------------------------------------------------------------"
  Write-Host "# Log:              $Script:LogFile"
  Write-Host "# Exports:          $Script:dest"
  Write-Host "#==============================================================================="
  Write-Host ""
  Write-Host ""
  Write-Host ""
}

#& Date time formatting for timestamped updated
Function Get-Now{
  # PowerShell Method - uncomment below is .NET is unavailable
  #$Script:Now = (get-date).tostring("[dd/MM HH:mm:ss:ffff]")
  # .NET Call which is faster than PowerShell Method - comment out below if .NET is unavailable
  $Script:Now = ([DateTime]::Now).tostring("[dd/MM HH:mm:ss:ffff]")
}

#& Clean up log files in script root older than 15 days
Function Clear-TransLogs{
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Cleaning up transaction logs over 15 days old" @cwarning
  Get-ChildItem $PSScriptRoot -recurse "*$Script:ScriptName.log" -force | Where-Object {$_.lastwritetime -lt (get-date).adddays(-15)} | Remove-Item -force
}

#& TestPath function for testing and creating directories
Function Invoke-TestPath{
  [CmdletBinding()]
  param (
      #^ Path parameter for testing/creating destination paths
      [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
      [String]
      $ParamPath
  )
  Try{
      # Check to see if the report location exists, if not create it
      if ((Test-Path -Path $ParamPath -PathType Container) -eq $false){
          Get-Now
          Write-Host "$Script:Now [INFORMATION] Destination Path $($ParamPath) does not exist: creating...." @chighlight
          New-Item $ParamPath -ItemType Directory | Out-Null
          Get-Now
          Write-Verbose "$Script:Now [INFORMATION] Destination Path $($ParamPath) created"
      }
  }  
  Catch{
      #! Error handling for folder creation 
      Get-Now
      Write-Host "$Script:Now [Error] Error creating directories" @cerror
      Write-Host $PSItem.Exception.Message
      Stop-Transcript
      Break
  }
  Finally{
    $Error.Clear()                                                                     # Clear error log
  }
}

#& Test for exchange console loaded and connected
Function Test-Exchange{
    try {
      $IsExchangeShell = Get-ExchangeServer -ErrorAction SilentlyContinue
    } 
    Catch {}
    if ($null -eq $IsExchangeShell){
      Get-Now
      write-host "$Script:Now [INFORMATION] Exchange - is not connected - connecting to Exchange ...." -ForegroundColor Magenta
      Import-Module "$env:exchangeinstallpath\Bin\RemoteExchange.ps1"
      Connect-ExchangeServer -auto -ClientApplication:ManagementShell
    }
    else {
      Get-Now
      write-host "$Script:Now [INFORMATION] Already connected to Exchange - proceeding" -ForegroundColor Green 
    }
}

#& Invoke Popup message
Function Invoke-Popup {
  Param(
      $message,
      $title,
      $buttonstyle,
      $icontype,
      $timeout
  )
  
  $shell = New-Object -ComObject WScript.Shell
  $script:value = $shell.Popup($message, $timeout, $title, $buttons.$buttonstyle + $icon.$icontype + $Script:ShowOnTop)

}

Function Get-ExHealth{

  $exsvr = Get-ExchangeServer | Select-Object -ExpandProperty Name

  Get-ExchangeServer | Select-Object Name,AdminDisplayVersion,Site | Format-Table | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  ForEach ($ex in $exsvr){
    Write-Host "[INFORMATION] Getting Exchange Service Health for $ex" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Write-Host "$_" -ForegroundColor Yellow -NoNewline ; Test-ServiceHealth -Server $ex | Format-Table Services* -Auto | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Write-Host "" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

    Write-Host "[INFORMATION] Database Copy Status for $ex" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Get-MailboxDatabaseCopyStatus -Server $ex | Select-Object Name,Status,ContentIndexState,CopyQueueLength,ReplayQueueLength | Format-Table -AutoSize | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Write-Host "" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

    Write-Host "[INFORMATION] Testing Mapi for $ex" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Test-MAPIConnectivity -Server $ex | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Write-Host "" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

    Write-Host "[INFORMATION] Testing Replication Health for $ex" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Test-ReplicationHealth -Server $ex | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
    Write-Host "" | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  }



  <#
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Getting Exchange Service Health"
  $Svc_Health = Get-ExchangeServer | ForEach-Object {Write-Host "$_" -ForegroundColor Yellow -NoNewline ; Test-ServiceHealth | Format-Table Services* -Auto}
  $Svc_Health | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  Write-Host ''
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Getting Exchange Versions"
  $EX_Versions = Get-ExchangeServer | Select-Object Name,AdminDisplayVersion,Site | Format-Table
  $EX_Versions | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  Write-Host ''
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Getting Database Copy Status"
  ForEach ($ex in $exsvr){
    $DB_Stat = Get-MailboxDatabaseCopyStatus -Server $ex | Select-Object Name,Status,ContentIndexState,CopyQueueLength,ReplayQueueLength | Format-Table -AutoSize
    $DB_Stat | Out-File $Script:dest\ExchangeOrgHealth.txt -Append
  }
  $DB_Stat = Get-MailboxDatabaseCopyStatus -Server $ex | Select-Object Name,Status,ContentIndexState,CopyQueueLength,ReplayQueueLength | Format-Table -AutoSize
  $DB_Stat | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  Write-Host ''
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Testing Mapi"
  $Mapi = Test-MAPIConnectivity
  $Mapi | Out-File $Script:dest\ExchangeOrgHealth.txt -Append

  Write-Host ''
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Getting Exchange Versions"
  $DB_Avil = Get-DatabaseAvailabilityGroup | Select-Object -ExpandProperty Servers | Test-ReplicationHealth
  $DB_Avil| Out-File $Script:dest\ExchangeOrgHealth.txt -Append
#>
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
<#
? ---------------------------------------------------------- [NOTES:] -------------------------------------------------------------
& Best veiwed and edited with Microsoft Visual Studio Code with colorful comments extension
^ Transcription logging formatting use Get-Now before write-host to return current timestamp into $Scipt:Now variable
  Write-Host "$Script:Now [INFORMATION] Information Message"
  Write-Host "$Script:Now [INFORMATION] Highlighted Information Message" @chighlight
  Write-Host "$Script:Now [WARNING] Warning Message" @cwarning
  Write-Host "$Script:Now [ERROR] Error Message" @cerror

? ---------------------------------------------------------------------------------------------------------------------------------
#>

# Script Execution goes here

Start-Logging                                                                           # Start Transcription logging
Get-ScriptInfo                                                                          # Display Script Info
Clear-TransLogs                                                                         # Clear logs over 15 days old
Test-Exchange                                                                           # Test connection to Exchange

$script:server = Get-ExchangeServer | Select-Object Name | Out-GridView -Title "Select a Server to bring online" -PassThru | Select-Object -ExpandProperty Name
$script:server = $script:server + $script:Domain

# Check Services
Write-Host ""
Get-Now
Write-Host "$Script:Now [INFORMATION] Checking services on $script:server"
Write-Host "Please be aware that the IMAP Service may not be running while the host is in maintenance mode"
Get-ExchangeServer -Identity $script:server | ForEach-Object {Write-Host "$_" -ForegroundColor Yellow -NoNewline ; Test-ServiceHealth -Server $script:server | Format-Table Services* -Auto}
Invoke-Popup -message "Are the services on $script:server running as expected?" -title "Service Check" -buttonstyle $bstyle.YN -icontype $itype.Q -timeout 600

if ($clickedButton.$value -eq 'No')
{
  Get-Now
  Write-Host "$Script:Now [ERROR] Service Validation failed for $script:server exiting"
  exit
}
else {
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Service Vaildation successfuly proceding...."
  Write-Host ""
}

# Vaildate Exchange Search Host Controller Service is not disabled
Get-Now
Write-Host "$Script:Now [INFORMATION] Checking Exchange Search Host Controller Service Status...."
$ESHC_Disabled = Get-Service -Name "HostControllerService" -ComputerName $script:server | Select-Object -ExpandProperty StartType
if ($ESHC_Disabled -eq "Disabled"){
  Get-Now
  write-host "$Script:Now [WARNING] Exchange Search Host Controller Service is $ESHC_Disabled re-enabling service to Automatic"
  Set-Service -Name "HostControllerService" -ComputerName $script:server -StartupType Automatic
  Get-Now
  Write-host "$Script:Now [INFORMATION] Exchange Search Host Controller Service has been re-enabled starting service...."
  Get-Service -Name "HostControllerService" -ComputerName $script:server | Start-Service -PassThru
}
else {
  Get-Now
  write-host "$Script:Now [INFORMATION] Exchange Search Host Controller Service is $ESHC_Disabled contuining...."
}

Get-Now
Write-Host "$Script:Now [INFORMATION] Setting SeverWideOffline to Active on $script:server"
Set-ServerComponentState $script:server -Component ServerWideOffline -State Active -Requester Maintenance

Get-Now
Write-Host "$Script:Now [INFORMATION] Resuming cluster operations on $script:server"
Resume-ClusterNode -Name $script:server

Get-Now
Write-Host "$Script:Now [INFORMATION] Set DatabaseCopyAutoActivationPolicy to Unrestricted on $script:server"
Set-MailboxServer $script:server -DatabaseCopyAutoActivationPolicy Unrestricted

Get-Now
Write-Host "$Script:Now [INFORMATION] Setting DatabaseCopyActivationDisabledAndMoveNow to false on $script:server"
Set-MailboxServer $script:server -DatabaseCopyActivationDisabledAndMoveNow $false

Get-Now
Write-Host "$Script:Now [INFORMATION] Setting Hub Transport to Active on $script:server"
Set-ServerComponentState $script:server -Component HubTransport -State Active -Requester Maintenance

# Retest Exchange Services
Get-Now
Write-Host "$Script:Now [INFORMATION] Re-checking services on $script:server"
Write-Host "Note: All Services should be showing as running"
Get-ExchangeServer -Identity $script:server | ForEach-Object {Write-Host "$_" -ForegroundColor Yellow -NoNewline ; Test-ServiceHealth -Server $script:server | Format-Table Services* -Auto}
Invoke-Popup -message "Are the services on $script:server running as expected?" -title "Service Check" -buttonstyle $bstyle.YN -icontype $itype.Q -timeout 600

if ($clickedButton.$value -eq 'No')
{
  Get-Now
  Write-Host "$Script:Now [ERROR] Service Validation failed for $script:server exiting"
  exit
}
else {
  Get-Now
  Write-Host "$Script:Now [INFORMATION] Service Vaildation successfuly proceding..."
  Write-Host ""
}

Get-ExHealth