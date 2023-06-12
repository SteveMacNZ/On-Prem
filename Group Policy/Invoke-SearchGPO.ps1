<#
.SYNOPSIS
  Searches Active Directory group policy for policies with a specific search string specified, and exports matching policies to file
.DESCRIPTION
  Searches Active Directory group policy for policies with a specific search string specified, and exports matching policies to file 
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  18/05/2023
  Purpose/Change: Initial Release
.LINK
  None
.EXAMPLE
  ^ . Invoke-SearchGPO.ps1
  To search for specific GPO setting

#>

#requires -version 4
#region ------------------------------------------------------[Script Parameters]--------------------------------------------------

Param (
  #Script parameters go here
)

#endregion
#region ------------------------------------------------------[Initialisations]----------------------------------------------------

#& Global Error Action
#$ErrorActionPreference = 'SilentlyContinue'

#& Module Imports
#Import-Module ActiveDirectory

#& Includes - Scripts & Modules
#. .\Get-CommonFunctions.ps1                          # Include Common Functions

#endregion
#region -------------------------------------------------------[Declarations]------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:Now         = ''                                                            # script sourced veriable for Get-Now function
$Script:dest        = $PSScriptRoot                                                 # Destination path
$Script:LogDir      = $Script:dest                                                  # Logdir for Clear-TransLogs function for $PSScript Root
$Script:ScriptName  = 'Invoke-SearchGPO'                                            # Script Name used in the Open Dialogue
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = '20b1ff23-c782-40e0-bd05-b4204678eab7'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '1.0.0.0'                                               # Script Version Number
$Script:Client      = ''                                                            # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = "Active Directory Group Policy Search"                        # Description displayed in Get-ScriptInfo function
$Script:Desc2       = "Searches for GPO policies that match the serch string"       # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#^ Script specific variables
$Script:String      = "MinimumPasswordLength"                                       # Enter Group Policy Search String (e.g. Windows Defender Firewall)
$Script:Domain      = "corp.justice.govt.nz"                                        # Domain Name
$Script:NearestDC   = ""                                                            # Place holder for Nearest Domain Controller
#endregion
#region --------------------------------------------------------[Hash Tables]------------------------------------------------------
#* Hash table for Write-Host Errors to be used as spatted
$cerror = @{ForeGroundColor = "Red"; BackgroundColor = "white"}
#* Hash table for Write-Host Warnings to be used as spatted
$cwarning = @{ForeGroundColor = "Magenta"; BackgroundColor = "white"}
#* Hash table for Write-Host highlighted to be used as spatted
$chighlight = @{ForeGroundColor = "Blue"; BackgroundColor = "white"}

#* Hash table for Write-Host green text to be used as spatted
$tgreen = @{ForeGroundColor = "Green"}
#* Hash table for Write-Host red text to be used as spatted
$tred = @{ForeGroundColor = "Red"}

#^ Dummy Write-host for spatted formatting
Write-Host @chighlight
Write-Host @cwarning
Write-Host @cerror
Write-Host @tgreen
Write-Host @tred

#endregion
#region -------------------------------------------------------[Functions]---------------------------------------------------------

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
  Write-Host "$Script:Now - Cleaning up transaction logs over 15 days old" @cwarning
  Get-ChildItem $Script:LogDir -recurse "*$Script:ScriptName.log" -force | Where-Object {$_.lastwritetime -lt (get-date).adddays(-15)} | Remove-Item -force
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
}

#& Display ScriptInfo
Function Get-ScriptInfo{
    
  Write-Host "#==============================================================================="
  Write-Host "# Name:             $Script:ScriptName"
  Write-Host "# Version:          $Script:Version"
  Write-Host "# GUID:             $Script:GUID"
  Write-Host "# Running As:       $Script:WHO"
  Write-Host "# PS Architecture:  $Script:PSArchitecture"
  Write-Host "# Description:"
  Write-Host "# $Script:Desc"
  Write-Host "# $Script:Desc2"
  Write-Host "#-------------------------------------------------------------------------------"
  Write-Host "# Log:              $Script:LogFile"
  Write-Host "# Exports:          $Script:dest"
  Write-Host "#==============================================================================="
  Write-Host ""
  Write-Host ""
  Write-Host ""
}

Function Get-PSArch{
  # Determines if PowerShell is running as a x86 or x64 bit process
  $Arch = [intptr]::Size

  If ($Arch -eq 4){Write-Host "PowerShell is running the script as x86"; $Script:PSArchitecture = "x86 [32 bit]"}
  If ($Arch -eq 8){Write-Host "PowerShell is running the script as x64"; $Script:PSArchitecture = "x64 [64 bit]"}
}

#endregion
#region ------------------------------------------------------------[Classes]-------------------------------------------------------------

#endregion
#region -----------------------------------------------------------[Execution]------------------------------------------------------------
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

Start-Logging                                                                                       # Start Transcription logging
Get-PSArch                                                                                          # Get PS Architecture
Get-ScriptInfo                                                                                      # Display Script Info
Clear-TransLogs                                                                                     # Clear logs over 15 days old

Get-Now
Write-Host "$Script:Now [INFORMATION] Detecting nearest AD Domain Controller"

$Script:NearestDC = (Get-ADDomainController -Discover -NextClosestSite).Name
Get-Now
Write-Host "$Script:Now [INFORMATION] $Script:NearestDC selected as AD Domain Controller"

Get-Now
Write-Host "$Script:Now [INFORMATION] Enumerating Group Policy Objects"
#Get a list of GPOs from the domain
$GPOs = Get-GPO -All -Domain $Domain -Server $Script:NearestDC | Sort-Object DisplayName

$counter = 0
$maximum = $GPOs.Count  # number of items to be processed

Write-host "[INFORMATION $maximum Group Policy Object found]" @chighlight

#$tc = $GPOs.count
#$lc = 0

#Go through each Object and check its XML against $String
Foreach ($GPO in $GPOs)  {
    #$lc++
    #Write-Progress "Processing $tc Objects" -Status "Completed: $lc of $tc. Remaining: $($tc-$lc)" -PercentComplete ($lc/$tc*100)
    $counter++
    $percentCompleted = $counter * 100 / $maximum

    #$message = '{0:p1} completed, processing {1}. Remaining {3}' -f ( $percentCompleted/100), $GPO.DisplayName, ( $maximum-$counter)
    $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $GPO.DisplayName
    Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
    Write-Host $message
    
    #Get Current GPO Report (XML)
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Producing GPO Report for $($GPO.DisplayName)"
     
    $CurrentGPOReport = Get-GPOReport -Guid $GPO.Id -ReportType Xml -Domain $Script:Domain -Server $Script:NearestDC

    If ($CurrentGPOReport -match $Script:String)  {
        Get-Now
        Write-Host "$Script:Now [INFORMATION] $($GPO.DisplayName) Matched" -Foregroundcolor Green
        Write-Host "A Group Policy matching ""$($String)"" has been found:" -Foregroundcolor Green
        Write-Host "-  GPO Name: $($GPO.DisplayName)" -Foregroundcolor Green
        Write-Host "-  GPO Id: $($GPO.Id)" -Foregroundcolor Green
        Write-Host "-  GPO Status: $($GPO.GpoStatus)" -Foregroundcolor Green
        $CurrentGPOReport | Out-File -FilePath $PSScriptRoot\$($GPO.DisplayName).xml
    } 
    else{
        Get-Now
        Write-Host "$Script:Now [INFORMATION] $($GPO.DisplayName) no match found" -ForegroundColor Yellow
    }

    # Clear CurrentGPOReport variable
    $null = $CurrentGPOReport

    Write-Host ""
}

Write-Output ''                                                                                     # write Line spacer into Transcription file
Get-Now
Write-Output "$Script:Now [INFORMATION] Processing finished + any outputs"                          # Write Status Update to Transcription file

Get-Now
Write-Host "================================================================================"  
Write-Host "================= $Script:Now Processing Finished ====================" 
Write-Host "================================================================================" 

Stop-Transcript                                                                           # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------