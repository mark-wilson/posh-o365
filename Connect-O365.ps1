Param(
    [String]$Tenant
  )

################################################################################
# Connect-O365.ps1                                                             #
# Takes one parameter (Tenant name) and connects to all Office 365 services    #
#                                                                              #
# Requires management modules for Azure AD, SharePoint Online, Skype for       #
# Business Online and Exchange Online to be available on the system used to    #
# run the script                                                               #
#                                                                              #
# Also requires credentials to access the tenant                               #
#                                                                              #
# Based on https://technet.microsoft.com/en-gb/library/dn568015.aspx           #
################################################################################

function loadModule($Name)
# Checks to see if a PowerShell module exists and is loaded, loading it if necessary
# Based on http://blogs.technet.com/b/heyscriptingguy/archive/2010/07/11/hey-scripting-guy-weekend-scripter-checking-for-module-dependencies-in-windows-powershell.aspx
{
  Write-Host "  Checking to see if PowerShell module" $Name "is loaded."
  If(-not(Get-Module -name $Name)) 
  # Is module loaded already to load?
  { 
    If(Get-Module -ListAvailable | Where-Object { $_.name -eq $Name })
    #Is module available? 
    { 
      # Module available, so import it
      Write-Host "  -" $Name "module is available. Importing..."
      Import-Module -Name $Name 
      $true 
    }
    Else
    {
      # Module not available
      Write-Host "  -" $Name "module is not available."
      $false   
    }  
  } 
  Else 
  {
    # Module already loaded
    Write-Host "  -" $Name "module is already loaded."
    $true
  }  
} #End function loadModule 

# Initialise variables
$Syntax = "SYNTAX: Connect-O365 TenantName"
$ErrorCode = 0
$Url = ""
$SPOServiceUrl = "https://tenant-admin.sharepoint.com"
$O365Creds = ""
$Result = ""
$SfBOSession = ""
$ExchangeSession = ""
$CCSession = ""

Write-Host $Syntax "`n" -BackgroundColor "White" -ForegroundColor "DarkBlue"
Write-Host "Validating input:"

# Check that a tenant name was provided
Write-Host "  Checking new tenant name has been provided: $Tenant"
If($Tenant)
{
  Write-Host -NoNewLine "~" -BackgroundColor "Yellow" -ForegroundColor "Black"
  Write-Host " Office 365 tenant name provided ($Tenant). Note that no validation has taken place as to whether this tenant exists."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Tenant name is missing. For an Office 365 domain name @markwilson.onmicrosoft.com the tenant name would be markwilson."
  $ErrorCode = 1
}

If($ErrorCode -ne 0)
{
  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
}

# Get credentials
$O365Creds = Get-Credential -Message "Please supply Office 365 Global Administrator credentials for $Tenant.onmicrosoft.com"

# Load the Azure Active Directory module
If(LoadModule -name MSOnline)
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Azure Active Directory module is available and loaded."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Azure Active Directory module is not available."
  $ErrorCode = 2
}

If($ErrorCode -ne 0)
{
  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
}

# Load the SharePoint Online Management module
If(LoadModule -name Microsoft.Online.SharePoint.PowerShell)
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " SharePoint Online Management module is available and loaded."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " SharePoint Online Management module is not available."
  $ErrorCode = 3
}

If($ErrorCode -ne 0)
{
  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
}

# Load the Lync Online Management module
If(LoadModule -name LyncOnlineConnector)
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Lync Online Management module is available and loaded."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Lync Online Management module is not available."
  $ErrorCode = 4
}

If($ErrorCode -ne 0)
{
  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
}

# All parameters seem to be valid and the pre-requisite PowerShell modules exist
Write-Host "`nIntended action: Connecting to $Tenant.onmicrosoft.com:"

# Connect to Microsoft Online
Write-Host "  Attempting to connect to Microsoft Online..."
$Result = Connect-MsolService -Credential $O365Creds

#If($Result -ne $null)
#{ 
#  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
#  Write-Host " Error connecting to Microsoft Online."
#  $ErrorCode = 6
#  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
#}
#Else
#{ 
#  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
#  Write-Host " Connected to Microsoft Online."
#}

# Calculate the URLs for connection to SharePoint Online
$SPOServiceUrl = $SPOServiceUrl -replace "tenant",$Tenant
$Url = $SPOServiceUrl+$ODBBaseUrl+$Url

# Connect to SharePoint Online
Write-Host "  Attempting to connect to SharePoint Online service at $SPOServiceUrl..."
$Result = Connect-SPOService –url $SPOServiceURL –credential $O365Creds

If($Result -ne $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error connecting to $SPOServiceURL."
  $ErrorCode = 6
  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Connected to SharePoint Online."
}

# Connect to Skype for Business Online
Write-Host "  Attempting to connect to Skype for Business Online..."

$SfBOSession = New-CsOnlineSession -Credential $O365Creds
$Result = Import-PSSession $SfBOSession

#If($Result -ne $null)
#{ 
#  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
#  Write-Host " Error creating Skype for Business Online session."
#  $ErrorCode = 7
#  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
#}
#Else
#{ 
#  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
#  Write-Host " Connected to Skype for Business Online."
#}

# Connect to Exchange Online
Write-Host "  Attempting to connect to Exchange Online..."
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Creds -Authentication "Basic" -AllowRedirection
$Result = $ExchangeSession

#If($Result -ne $null)
#{ 
#  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
#  Write-Host " Error creating Exchange Online session."
#  $ErrorCode = 8
#  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
#}
#Else
#{ 
#  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
#  Write-Host " Connected to Exchange Online."
#}

# Connect to Compliance Center
Write-Host "  Attempting to connect to Compliance Center..."
$CCSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $O365Creds -Authentication Basic -AllowRedirection
$Result = Import-PSSession $CCSession -Prefix cc

#If($Result -ne $null)
#{ 
#  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
#  Write-Host " Error creating Compliance Center session."
#  $ErrorCode = 9
#  Throw "Connect-O365 There was a problem. Error code $ErrorCode."
#}
#Else
#{ 
#  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
#  Write-Host " Connected to Compliance Center."
#}

Write-Host "  PowerShell sessions created. Current sessions are:"
Get-PSSession
Write-Host "`n  Remember to close sessions gracefully with Remove-PSSession. Alternatively close all sessions with `n  Get-PSSession | Remove-PSSession."