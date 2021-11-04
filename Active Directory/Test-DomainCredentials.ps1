<#
.SYNOPSIS
  Testing of Active Directory Domain credentials 
.DESCRIPTION
  Testing of Active Directory Domain credentials - script taken and adapted from https://www.powershellbros.com/test-credentials-using-powershell-function/
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         <Name>     
  Creation Date:  4/11/21
  Purpose/Change: Initial Script
.LINK
  None
.EXAMPLE
  .\Test-DomainCredentials.ps1
  description of what the example does
 
#>

#requires -version 4
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#-----------------------------------------------------------[Hash Tables]-----------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Test-Cred {
           
    [CmdletBinding()]
    [OutputType([String])] 
       
    Param ( 
        [Parameter( 
            Mandatory = $false, 
            ValueFromPipeLine = $true, 
            ValueFromPipelineByPropertyName = $true
        )] 
        [Alias( 
            'PSCredential'
        )] 
        [ValidateNotNull()] 
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] 
        $Credentials
    )
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
      
    If($Credentials -eq $null)
    {
        Try
        {
            $Credentials = Get-Credential "$env:USERDNSDOMAIN\$env:username" -ErrorAction Stop
        }
        Catch
        {
            $ErrorMsg = $_.Exception.Message
            Write-Warning "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
    }
      
    # Checking module
    Try
    {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
  
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    }
    Catch
    {
        $_.Exception.Message
        Continue
    }
  
    If(!$domain)
    {
        Write-Warning "Something went wrong"
    }
    Else
    {
        If ($null -ne $domain.name)
        {
            return "Authenticated"
        }
        Else
        {
            return "Not authenticated"
        }
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Test-Cred
