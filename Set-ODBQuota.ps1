Param(
    [String]$Upn,
    [Int]$QuotaMb,
    [String]$Tenant
  )

################################################################################
# Set-ODBQuota.ps1                                                             #
#                                                                              #
# Takes three parameters (UPN, Storage Quota, Tenant name) and sets the        #
# specified user's OneDrive for Business storage quota                         #
#                                                                              #
# Requires the SharePoint Online Management module to be available on the      #
# system used to run the script                                                #
#                                                                              #
# Also requires credentials to access the tenant                               #
################################################################################

function isEmailAddress($Object)
# Checks to see if a string looks like an email address
# Based on http://powershell.com/cs/media/p/389.aspx
# Doesn't check for missing . though...
{   
    ($Object -as [System.Net.Mail.MailAddress]).Address -eq $object -and $object -ne $null  
} #End function isEmailAddress

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
$Syntax = "SYNTAX: Set-ODBQuota UPN StorageQuota TenantName"
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
Write-Host "  Checking UPN:" $Upn

# Check UPN looks like an email address
If (isEmailAddress $Upn)
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " $Upn looks like an email address."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " $Upn does not appear to be a valid user principal name (UPN).`n  - Does not look like an email address. UPNs should look like alias@domain.tld."
  $ErrorCode = 1
}

# Check UPN contains .
If($Upn -match "\.")
{
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " $Upn contains a dot (.)."
}
Else
{
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " $Upn does not appear to be a valid user principal name (UPN).`n  - Does not contain `".`". UPNs should look like alias@domain.tld."
  $ErrorCode = 2
}

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
  Throw "Set-ODBQuota There was a problem. Error code $ErrorCode."
}

# All parameters seem to be valid and the pre-requisite PowerShell module exists
$QuotaGB = $QuotaMb/1024
Write-Host "`nIntended action: $Upn will have the OneDrive for Business Storage Quota set to $QuotaMb MB ($QuotaGB GB) for $Tenant.onmicrosoft.com:"

# Calculate the URLs for connection to SharePoint Online and for the OneDrive for Business site
$Url = $Upn -replace "@", "_"
$Url = $Url -replace "\.", "_"
$SPOServiceUrl = $SPOServiceUrl -replace "tenant",$Tenant
$Url = $SPOServiceUrl+$ODBBaseUrl+$Url
$SPOServiceUrl = $SPOServiceUrl -replace "-my","-admin"

Write-Host "  $ProductName URL calculated as $Url."
Write-Host "  Attempting to connect to SharePoint Online service at $SPOServiceUrl..."

# Connect to SharePoint Online
$O365Creds = Get-Credential -Message "Please supply Office 365 Global Administrator credentials for $Tenant.onmicrosoft.com"
$Result = Connect-SPOService –url $SPOServiceURL –credential $O365Creds

If($Result -ne $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error connecting to $SPOServiceURL."
  $ErrorCode = 6
  Throw "Set-ODBQuota There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Connected to $SPOServiceURL."
}

# Read the existing site details. This will fail if the user has never logged on to OneDrive for Business (so their site is unprovisioned).
$SPOSite = Get-SPOSite -Identity $Url

If($SPOSite.StorageQuota -lt $MinQuotaMb -Or $SPOSite.StorageQuota -gt $MaxQuotaMb)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error reading site details."
  $ErrorCode = 7
  Throw "Set-ODBQuota There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host "  Current storage quota is" $SPOSite.StorageQuota "MB."
}

# Set the new quota
Write-Host "  Attempting to set new storage quota..."
$Result = Set-SPOSite -Identity $Url -StorageQuota $QuotaMb

If($Result -ne $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error setting new storage quota."
  $ErrorCode = 8
  Throw "Set-ODBQuota There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Storage quota set."
}

# Read the site details again
$SPOSite = Get-SPOSite -Identity $Url
Write-Host "  New storage quota is" $SPOSite.StorageQuota "MB."