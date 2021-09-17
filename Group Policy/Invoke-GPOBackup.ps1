<#
.SYNOPSIS
  Runs as a scheduled task weekly to take copies of GPOs and save to network share
.DESCRIPTION
  Runs as a scheduled task weekly to take copies of GPOs and save to network share. A new folder for each weekly export will be created
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  Log file for transcription logging
  Report of GPO Backups
  Backup of GPOs to file share
.NOTES
  Version:        1.0
  Author:         Steve McIntyre     
  Creation Date:  25/08/21
  Purpose/Change: Initial Script
.LINK
  None
.EXAMPLE
  .\Invoke-GPOBackup.ps1
  Completes a Weekly GPO backup to disk
 
#>


#Requires -version 4 -Modules ActiveDirectory
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins
Import-Module ActiveDirectory

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Invoke-GPOBackup'                                                        # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
#$Script:BDir                = '\\server\share\sub-folder'                                               # Path for GPO Backups
$Script:BDir                = 'D:\GPOBackups'                                                           # Path for GPO Backups
$Script:BFolder             = $Script:BDir + "\" + $Script:Date + "_Backups"                            # Weekly folder for Backup
$Script:ReportFile          = $Script:BFolder + "\" + $Script:Date + "_" + "GPOBackupResults.txt"       # Reportfile location and name
$Script:Zip                 = $Script:BDir + "\" + $Script:Date + "-GPO-Backup.zip"                     # Location for Zip file backup
$Script:GUID                = '42d80bf0-9320-4bfd-b6e5-964a10f51a3f'                                    # Script GUID
#^ Use New-Guid cmdlet to generate new script GUID for each version change of the script

#-----------------------------------------------------------[Hash Tables]-----------------------------------------------------------
#* Hash table for Write-Host Errors to be used as spatted
$cerror = @{ForeGroundColor = "Red"; BackgroundColor = "white"}
#* Hash table for Write-Host Warnings to be used as spatted
$cwarning = @{ForeGroundColor = "Magenta"; BackgroundColor = "white"}
#* Hash table for Write-Host highlighted to be used as spatted
$chighlight = @{ForeGroundColor = "Blue"; BackgroundColor = "white"}

#-----------------------------------------------------------[Functions]------------------------------------------------------------
#& Start Transcriptions
Function Start-Logging{

    try {
        Stop-Transcript | Out-Null
    } catch [System.InvalidOperationException] { }                                                      # jobs are running
    $ErrorActionPreference = "Continue"                                                                 # Set Error Action Handling
    Get-Now                                                                                             # Get current date time
    Start-Transcript -path $Script:LogFile -IncludeInvocationHeader -Append                             # Start Transcription append if log exists
    Write-Host ''                                                                                       # write Line spacer into Transcription file
    Write-Host ''                                                                                       # write Line spacer into Transcription file
    Write-Host "========================================================" 
    Write-Host "====== $Script:Now Processing Started ========" 
    Write-Host "========================================================" 
    Write-Host ''                                                                                       # write Line spacer into Transcription file
    Write-Host ''                                                                                       # write Line spacer into Transcription file
}
  
#& Date time formatting for timestamped updated
Function Get-Now{
    $Script:Now = (get-date).tostring("[dd/MM HH:mm:ss:ffff]")
}

#& Clean up log files in script root older than 15 days
Function Clear-TransLogs{
    Get-Now
    Write-Host "$Script:Now - Cleaning up transaction logs over 15 days old" @cwarning
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
}

#& Adds exported GPOs in Zip file
Function Invoke-AddtoZip{
    [CmdletBinding()]
    param (
        #^ Path parameter for source and destination paths
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][String]$ParamSPath,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][String]$ParamDPath
    )
    Try{
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Adding $($ParamSPath) to zip file" @chighlight
        Compress-Archive -Path $ParamSPath -DestinationPath $ParamDPath
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Zip file: $ParamDPath created"
    }  
    Catch{
        #! Error handling for folder creation 
        Get-Now
        Write-Host "$Script:Now [ERROR] Unable to add files to ZIP" @cerror
        Write-Host $PSItem.Exception.Message
    }
    Finally{
        $Error.Clear()                                                                              # Clear error log
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
<#
? ---------------------------------------------------------- [NOTES:] -------------------------------------------------------------
& Best veiwed and edited with Microsoft Visual Studio Code with colorful comments extension
* Requires ActiveDirectory
* Transcription logging formatting use Get-Now before write-host to return current timestamp into $Scipt:Now variable
  Write-Host "$Script:Now [INFORMATION] Information Message"
  Write-Host "$Script:Now [INFORMATION] Highlighted Information Message" @chighlight
  Write-Host "$Script:Now [WARNING] Warning Message" @cwarning
  Write-Host "$Script:Now [ERROR] Error Message" @cerror
? ---------------------------------------------------------------------------------------------------------------------------------
#>

# Script Execution goes here
Start-Logging                                                                                       # Start Transcription logging
Clear-TransLogs                                                                                     # Clean up old transcription logs

Invoke-TestPath -ParamPath $Script:BDir                                                             # Test and create folder structure
Invoke-TestPath -ParamPath $Script:BFolder                                                          # Test and create folder structure

$NearestDC  = (Get-ADDomainController -Discover -NextClosestSite).Name                              # Discover the closest DC for processing
$Domain     = (Get-ADDomainController -Discover -NextClosestSite).Domain                            # Discover Domain Name
$GPOs       = Get-GPO -All -Domain $Domain -Server $NearestDC | Sort-Object DisplayName             # Get all GPOs 

$results =@()                                                                                       # Set up result array
$results += "========================================================"                              # Report log Header
$results += "====== Group Policy Backup results for $Script:Date ======"                            # Report log Header
$results += "========================================================"                              # Report log Header
$tc = $GPOs.count                                                                                   # Total counts of GPO Objects
$lc = 0                                                                                             # Initial lastcount

# Process each GPO Object attempt a backup and write results to array
Foreach ($GPO in $GPOs) {
    $lc++
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Processing $tc GPO Objects: Competed: $lc of $tc. Remaining: $($tc-$lc)"
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Working on $($GPO.DisplayName)"

    Try{
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Backup of GPO $($GPO.DisplayName) starting"
        $CurrentGPO = Backup-GPO -Guid $GPO.Id -Domain $Domain -Server $NearestDC -Path $Script:BFolder -Comment "Group Policy backup taken on $Script:Date"
    }
    Catch{
        #! Error handling for folder creation 
        Get-Now
        Write-Host "$Script:Now [ERROR] There was and error backing up $($GPO.DisplayName)" @cerror
        Write-Host $PSItem.Exception.Message    
    }
    Finally{
        $Error.Clear()                                                                              # Clear error log
    }

    If ($null -ne $CurrentGPO){
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Writing results of $($GPO.DisplayName) backup to report"
        $results += $CurrentGPO                                                                     # write backup info to results
    }
    else {
        Get-Now
        Write-Host "$Script:Now [WARNING] Backup of $($GPO.DisplayName) did not process, no results to process" @cerror
    }

    $CurrentGPO = $null                                                                             # Reset CurrenGPO variable  

}

$results | Out-File -FilePath $Script:ReportFile                                                    # Write results to reportfile   

Get-Now
Write-Host "$Script:Now [INFORMATION] Group Policy objects backed up to $Script:BFolder" -ForeGroundColor Green

Invoke-AddtoZip -ParamSPath $Script:BFolder  -ParmDPath $Script:Zip                                 # Zip GPO backup folder

Get-Now
Write-Output  "========================================================" 
Write-Output  "======== $Script:Now Processing Finished =========" 
Write-Output  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------
