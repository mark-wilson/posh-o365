Param(
    [String]$Filename,
    [Int]$QuotaMb,
    [String]$Tenant
  )

################################################################################
# Check-ODBQuotas.ps1
# Takes three parameters (CSV file containing a list of UPNs, Storage Quota, Tenant name) and checks the specified users' OneDrive for Business storage quotas
# Requires the SharePoint Online Management module to be available on the system
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
$Syntax = "SYNTAX: Check-ODBQuotas CSVFile StorageQuota TenantName"
$UpnList = @()
$ProductName = "OneDrive for Business"
$MinQuotaMb = 1024
$MinQuotaGb = $MinQuotaMb/1024
$MaxQuotaMb = 1048576
$MaxQuotaGb = $MaxQuotaMb/1024
$ErrorCode = 0
$Url = ""
$SPOServiceUrl = "https://tenant-my.sharepoint.com"
$ODBBaseUrl = "/personal/"
$O365Creds = ""
$SPOSite = ""
$Result = ""

Write-Host $Syntax "`n" -BackgroundColor "White" -ForegroundColor "DarkBlue"
Write-Host "Validating input:"
Write-Host "  Opening $Filename and reading UPNs"

$UpnList = Import-CSV $Filename

If($UpnList[0] -eq $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error reading $Filename."
  $ErrorCode = 1
  Throw "Check-ODBQuotas There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Imported" $UpnList.Count "UPNs from $Filename."
}

#$UpnList | ForEach-Object {$_}

# Check storage quota value (although the Set-SPOSite command will also do this it will silently fail to a minimum or maximum value)
Write-Host "  Checking new $ProductName storage quota value:" $QuotaMb "MB"
If($QuotaMb -lt $MinQuotaMb -Or $QuotaMb -gt $MaxQuotaMb)
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " $QuotaMb is not a valid size. $ProductName storage quotas must be between $MinQuotaMb MB ($MinQuotaGb GB) and $MaxQuotaMb MB ($MaxQuotaGb GB)"
  $ErrorCode = 3
}
Else
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " $QuotaMb MB appears to be within range."
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
  $ErrorCode = 5
}

If($ErrorCode -ne 0)
{
  Throw "Check-ODBQuotas There was a problem. Error code $ErrorCode."
}

# All parameters seem to be valid and the pre-requisite PowerShell module exists
Write-Host "`nIntended action: Checking OneDrive for Business storage quotas on $Tenant.onmicrosoft.com:"

# Calculate the URL for connection to SharePoint Online
$SPOServiceUrl = $SPOServiceUrl -replace "tenant",$Tenant
$SPOServiceURL = $SPOServiceUrl -replace "-my","-admin"

Write-Host "  Attempting to connect to SharePoint Online service at $SPOServiceUrl..."

# Connect to SharePoint Online
$O365Creds = Get-Credential -Message "Please supply Office 365 Global Administrator credentials for $Tenant.onmicrosoft.com"
$Result = Connect-SPOService –url $SPOServiceURL –credential $O365Creds

If($Result -ne $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error connecting to $SPOServiceURL."
  $ErrorCode = 6
  Throw "Check-ODBQuotas There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Connected to $SPOServiceURL."
}

ForEach ($Upn in $UpnList)
{
  # Calculate the URL for the OneDrive for Business site
  $SPOSite = $null
  $Url = $Upn.upn
  $Url = $Url -replace "@", "_"
  $Url = $Url -replace "\.", "_"
  $SPOServiceUrl = $SPOServiceUrl -replace "-admin","-my"
  $Url = $SPOServiceUrl+$ODBBaseUrl+$Url
  #Write-Host "  $ProductName URL calculated as $Url."

  # Read the existing site details. This will fail if the user has never logged on to OneDrive for Business (so their site is unprovisioned).
  $SPOSite = Get-SPOSite -Identity $Url

  If($SPOSite.StorageQuota -eq $null)
  { 
    Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
    Write-Host " Error reading site details for" $Upn.upn "(this user may not have a OneDrive for Business site provisioned)."
    $ErrorCode = 7
  }
  Else
  { 
    If ($SPOSite.StorageQuota -ne $QuotaMb)
    {
      Write-Host -NoNewLine "~" -BackgroundColor "Yellow" -ForegroundColor "Black"
      Write-Host " Storage quota for" $Upn.upn "requires a reset to $QuotaMb MB (it is currently set to" $SPOSite.StorageQuota "MB)."
    }
    Else
    {
      Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
      Write-Host " Storage quota for" $Upn.upn "is set as expected ($QuotaMb MB)."
    }
  }
}