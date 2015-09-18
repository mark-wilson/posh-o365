Param(
    [String]$Filename,
    [String]$Tenant
  )

################################################################################
# Set-O365Licences.ps1
# Takes two parameters (CSV file containing a list of UPNs and subscription, Tenant name) and applies the appropriate licence
# Requires the Microsoft Online Services module to be available on the system
# Also requires credentials to access the tenant
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
$Syntax = "SYNTAX: Set-O365Licences CSVFile TenantName"
$UserList = @()
$ErrorCode = 0
$Url = ""
$MSOLServiceUrl = "https://tenant.onmicrosoft.com"
$O365Creds = ""
$Result = ""
$AccountSku = ""
$Username = ""
$Licence = ""
$UserLicenceTest = ""

Write-Host $Syntax "`n" -BackgroundColor "White" -ForegroundColor "DarkBlue"
Write-Host "Validating input:"
Write-Host "  Opening $Filename and reading data"

$UserList = Import-CSV $Filename

If($UserList[0] -eq $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error reading $Filename."
  $ErrorCode = 1
  Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Imported" $UserList.Count "UPNs from $Filename."
}

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
  $ErrorCode = 4
}

# Load the Microsoft Online Services  module
If(LoadModule -name MSOnline)
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Microsoft Online Services module is available and loaded."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Microsoft Online Services module is not available."
  $ErrorCode = 5
}

If($ErrorCode -ne 0)
{
  Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
}

# All parameters seem to be valid and the pre-requisite PowerShell module exists
Write-Host "`nIntended action: Setting Office 365 licences on $Tenant.onmicrosoft.com:"

# Calculate the URL for connection to SharePoint Online
$MSOLServiceUrl = $MSOLServiceUrl -replace "tenant",$Tenant

Write-Host "  Attempting to connect to Office 365 at $MSOLServiceUrl..."

# Connect to Office 365
$O365Creds = Get-Credential -Message "Please supply Office 365 Global Administrator credentials for $Tenant.onmicrosoft.com"
$Result = Connect-MSOLService –credential $O365Creds

If($Result -ne $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error connecting to $MSOLServiceURL."
  $ErrorCode = 6
  Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Connected to $MSOLServiceURL."
}

$AccountSku = Get-MsolAccountSku

Write-Host "  This tenant has" $AccountSku.Count "subscription(s):"

#$AccountSku | fl

ForEach ($Subscription in $AccountSku)
{
  Write-Host "  -" $AccountSku.AccountSkuID "has used" $AccountSku.ActiveUnits "out of" $AccountSku.ConsumedUnits "licences."
}

ForEach ($User in $UserList)
{
  $Username = $User.UPN

  Write-Host "  Processing $Username."
  $UserLicenceTest = Get-MsolUser -UserPrincipalName $Username
  If($UserLicenceTest.IsLicensed)
  {
    Write-Host "  -" $UserLicenceTest.Licenses.Count "licences."
    ForEach ($Licence in $UserLicenceTest)
    {
      Write-Host "  -" $UserLicenceTest.Licenses[$Licence].ServiceStatus
    }
  }

  $Licence = $User.Licence

  # Calculate the Licence to be allocated
  Switch($Licence)
  {
    "None" {"  - $Username does not require a licence to be allocated."}
    "E1" {"  - $Username requires an $Licence licence."}
    "E3" {"  - $Username requires an $Licence licence."}
    default {"  - Could not determine the licence requirements for $Username."}
  }
}