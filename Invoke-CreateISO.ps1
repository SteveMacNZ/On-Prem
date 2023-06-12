<#
.SYNOPSIS
  Creates ISO file based source folder
.DESCRIPTION
  Creates ISO file based source folder 
.PARAMETER None
  None
.INPUTS
  -NewIsoFilePath   ==> Path ISO image will be saved 
  -ImageName        ==> ISO Image Name
  -SourceFilePath   ==> Source folder all files and folder under this folder will be added into ISO  
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  13/09/2022
  Purpose/Change: Initial Release
.LINK
  None
.EXAMPLE
  ^ . Invoke-CreateISO.ps1
  Creates ISO Image
  Invoke-CreateISO.ps1 -NewIsoFilePath $env:temp\MyTest.iso -ImageName Holiday -SourceFilePath 'C:\HolidayPics'

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
$Script:LogDir      = $PSScriptRoot                                                 # Logdir for Clear-TransLogs function for $PSScript Root
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:ScriptName  = ''                                                            # Script Name used in the Open Dialogue
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = '00000000-0000-0000-0000-000000000000'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '0.0.0.0'                                               # Script Version Number
$Script:Client      = ''                                                            # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = ""                                                            # Description displayed in Get-ScriptInfo function
$Script:Desc2       = ""                                                            # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#^ File Picker / Folder Picker Setup
[System.IO.FileInfo]$Script:File  = ''                                              # File var for Get-FilePicker Function
[System.IO.FileInfo]$Script:ISODir  = ''                                            # ISO Source selected with Get-FolderPicker Function
[System.IO.FileInfo]$Script:ISODest  = ''                                           # ISO Destination selected with Get-FolderPicker Function
$Script:FPDir       = '$PSScriptRoot'                                               # File Picker Initial Directory
$Script:FileTypes   = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # File types to be listed in file picker
$Script:FileIndex   = "2"                                                           # What file type to set as default in file picker (based on above order)

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

#& FilePicker function for selecting input file via explorer window
Function Get-FilePicker {
  Param ()
  [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
  $ofd = New-Object System.Windows.Forms.OpenFileDialog
  $ofd.InitialDirectory = $Script:FPDir                                                         # Sets initial directory to script root
  $ofd.Title            = "Select file for $Script:ScriptName"                                  # Title for the Open Dialogue
  $ofd.Filter           = $Script:FileTypes                                                     # File Types filter
  $ofd.FilterIndex      = $Script:FileIndex                                                     # What file type to default to
  $ofd.RestoreDirectory = $true                                                                 # Reset the directory path
  #$ofd.ShowHelp         = $true                                                                 # Legacy UI              
  $ofd.ShowHelp         = $false                                                                # Modern UI
  if($ofd.ShowDialog() -eq "OK") { $ofd.FileName }
  $Script:File = $ofd.Filename
}

#& FolderPicker function for selecting a folder via explorer window
Function Get-FolderPicker{
  param (
    #^ Description to use in the dialogue
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $InitialPath,
    #^ Description to use in the dialogue
    [Parameter(ValueFromPipeline=$true)]
    [String]
    $Description
  )
  
  [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
  $fdir = New-Object System.Windows.Forms.FolderBrowserDialog
  $fdir.InitialDirectory    = $InitialPath
  $fdir.ShowHiddenFiles     = $true
  $fdir.ShowNewFolderButton = $true
  $fdir.ShowPinnedPlaces    = $true
  $fdir.Description         = $Description
  $fdir.rootfolder          = "MyComputer"

  if($fdir.ShowDialog() -eq "OK"){ $folder += $fdir.SelectedPath }
  return $folder
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

# Function creates a new ISO file based on input provided
Function New-IsoFile
{
  param
  (
    # path to local folder to store in 
    # new ISO file (must exist)
    [Parameter(Mandatory)]
    [String]
    $SourceFilePath,

    # name of new ISO image (arbitrary, 
    # turns later into drive label)
    [String]
    $ImageName = 'MyCDROM',

    # path to ISO file to be created
    [Parameter(Mandatory)]
    [String]
    $NewIsoFilePath,

    # if specified, the source base folder is
    # included into the image file
    [switch]
    $IncludeRoot
  )
    
  # use this COM object to create the ISO file:
  $fsi = New-Object -ComObject IMAPI2FS.MsftFileSystemImage 

  # use this helper object to write a COM stream to a file:
  # compile the helper code using these parameters:
  $cp = [CodeDom.Compiler.CompilerParameters]::new()
  $cp.CompilerOptions = '/unsafe'
  $cp.WarningLevel = 4
  $cp.TreatWarningsAsErrors = $true
  $code = '
    using System;
    using System.IO;
    using System.Runtime.InteropServices.ComTypes;

    namespace CustomConverter
    {
     public static class Helper 
     {
      // writes a stream that came from COM to a filesystem file
      public static void WriteStreamToFile(object stream, string filePath) 
      {
       // open output stream to new file
       IStream inputStream = stream as IStream;
       FileStream outputFileStream = File.OpenWrite(filePath);
       int bytesRead = 0;
       byte[] data;

       // read stream in chunks of 2048 bytes and write to filesystem stream:
       do 
       {
        data = Read(inputStream, 2048, out bytesRead);  
        outputFileStream.Write(data, 0, bytesRead);
       } while (bytesRead == 2048);

       outputFileStream.Flush();
       outputFileStream.Close();
      }

      // read bytes from stream:
      unsafe static private byte[] Read(IStream stream, int byteCount, out int readCount) 
      {
       // create a new buffer to hold the read bytes:
       byte[] buffer = new byte[byteCount];
       // provide a pointer to the location where the actually read bytes are reported:
       int bytesRead = 0;
       int* ptr = &bytesRead;
       // do the read:
       stream.Read(buffer, byteCount, (IntPtr)ptr);   
       // return the read bytes by reference to the caller:
       readCount = bytesRead;
       // return the read bytes to the caller:
       return buffer;
      } 
     }
  }'

  Add-Type -CompilerParameters $cp -TypeDefinition $code 

  # define the ISO file properties:

  # create CDROM, Joliet and UDF file systems
  $fsi.FileSystemsToCreate = 7 
  $fsi.VolumeName = $ImageName
  # allow larger-than-CRRom-Sizes
  $fsi.FreeMediaBlocks = -1    

  $msg = 'Creating ISO File - this can take a couple of minutes.'
  Write-Host $msg -ForegroundColor Green
    
  # define folder structure to be written to image:
  $fsi.Root.AddTreeWithNamedStreams($SourceFilePath,$IncludeRoot. IsPresent)
        
  # create image and provide a stream to read it:
  $resultimage = $fsi.CreateResultImage()
  $resultStream = $resultimage.ImageStream

  # write stream to file
  [CustomConverter.Helper]::WriteStreamToFile($resultStream, $NewIsoFilePath )

  Write-Host 'DONE.' -ForegroundColor Green

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

Start-Logging                                                                           # Start Transcription logging
Get-ScriptInfo                                                                          # Display Script Info
Clear-TransLogs                                                                         # Clear logs over 15 days old

Get-Now
Write-Host "$Script:Now [INFORMATION] ISO Creating Script"
Write-Host

Get-Now
Write-Host "$Script:Now [INFORMATION] Select folder to be included in ISO Image file"
$Script:ISODir = Get-FolderPicker -InitialPath $Script:FPDir -Description "Select folder for ISO"   
Get-Now                                                                                 # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] $Script:ISODir has been selected for processing" @chighlight
Write-Host ""

Get-Now
Write-Host "$Script:Now [INFORMATION] Select folder to be included in ISO Image file"
$Script:ISODest = Get-FolderPicker -InitialPath $Script:FPDir -Description "Select destination for ISO"
Write-Host ""

Get-Now
Write-Host "$Script:Now [INFORMATION] Enter information for ISO image"
$ISOName = Read-Host -Prompt "Enter ISO filename (e.g., MyHoliday.iso)"
$ImgName = Read-Host -Prompt "Enter image name (e.g., Holiday)"

New-IsoFile -NewIsoFilePath $Script:ISODest\$ISOName -ImageName $ImgName -SourceFilePath $Script:ISODir


Stop-Transcript
#endregion
#---------------------------------------------------------[Execution Completed]----------------------------------------------------------