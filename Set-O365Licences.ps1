Param(
    [String]$FileName,
    [String]$Tenant
  )

################################################################################
# Set-O365Licences.ps1                                                         #
#                                                                              #
# Takes two parameters (CSV file containing a list of UPNs/subscriptions/      #
# locations, Tenant name) and applies the appropriate licence                  #
# Requires the Microsoft Online Services module to be available on the system  #
#                                                                              #
# Also requires credentials to access the tenant                               #
#                                                                              #
# Known issues: Error-checking on licence allocation is not working correctly  #
# (will say that licence is added even if command fails)                       #
#                                                                              #
# Notes: Script draws heavily on advice at http://windowsitpro.com/office-365/ #
# office-365-licensing-windows-powershell                                      #
#                                                                              #
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
$Syntax = "SYNTAX: Set-O365Licences CSVFile TenantName" # Syntax for this script
$UserList = @()                                         # Array used when reading list of users from $FileName
$ErrorCode = 0                                          # Used for error handling
$MSOLServiceUrl = "https://tenant.onmicrosoft.com"      # URL format for Microsoft Online Services (MSOL) tenants
$O365Creds = ""                                         # Used to store credentials for connection to MSOL
$Result = ""                                            # Used for error handling
$AccountSku = @()                                       # MSOL Account
$AccountSkuID = ""                                      # MSOL Account SKU - e.g. <tenant>:ENTERPRISEPACK
$AccountName = ""                                       # Account name, as recorded in the MSOL subscription
$Subscription = ""                                      # Subscription, as recorded against the $MSOLAccount
$UserName = ""                                          # User name currently being acted on (read from UPN in $FileName)
$Licence = ""                                           # Licence for $UserName (read from $FileName)
$Location = ""                                          # Usage Location for $UserName (read from $FileName)
$LicenceDetails = ""                                    # Licence information read from MSOL
$ServicePlanList = @()                                  # Array for list of service plans in a subscription - not currently used
$Index = 0                                              # Index for looping
$E3 = ":ENTERPRISEPACK"                                 # Identifier used by MSOL for E3
$E1 = ":STANDARDPACK"                                   # Identifier used by MSOL for E1
$Sku = ""                                               # Used for comparison of requested licence type with current licence

Write-Host $Syntax "`n" -BackgroundColor "White" -ForegroundColor "DarkBlue"
Write-Host "Validating input:"

# Read information from file
Write-Host "  Opening $FileName and reading data"

$UserList = Import-CSV $FileName

If($UserList[0] -eq $null)
{ 
  Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
  Write-Host " Error reading $FileName."
  $ErrorCode = 1
  Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
}
Else
{ 
  Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
  Write-Host " Imported" $UserList.Count "UPNs from $FileName."
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

# Load the Microsoft Online Services module
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

# Read account details
$AccountSku = Get-MsolAccountSku
$AccountName = $AccountSku[0].AccountName # $AccountSku is an array, but this entry should be the same for all subscriptions.

# List subscription details
Write-Host "  This tenant has" $AccountSku.Count "subscription(s):"

$Index = 0
ForEach ($Subscription in $AccountSku)
{
  Write-Host "  -" $AccountSku[$Index].AccountSkuID "has used" $AccountSku[$Index].ConsumedUnits "out of" $AccountSku[$Index].ActiveUnits "licences."
  $Index++
}

# Process each user in turn
ForEach ($User in $UserList)
{
  $UserName = $User.UPN
  Write-Host "  Processing $UserName."

  # Check if they have a licence
  $LicenceDetails = (Get-MsolUser -UserPrincipalName $UserName).Licenses

  # If there's a license, show the details.
  If ($LicenceDetails.Count -gt 0)
  {
    $AccountSkuID = $LicenceDetails.AccountSkuID
    Write-Host "  - The following licences are currently allocated:"
    Write-Host "    -" $AccountSkuID # "with the following service plans:"
#    ForEach ($Item in $LicenceDetails)
#    {
#      $ServicePlanList = $Item.ServiceStatus #List of services
#      
#      For($Index=0; $Index -lt $ServicePlanList.Count; $Index++)
#      {
#         Write-Host "      - " -NoNewline
#         $ServicePlanList[$Index]
#      }
#    }
  }
  Else
  {
    Write-Host "  - No licences currently allocated."
    $AccountSkuID = "None"
  }

  $Licence = $User.Licence
  $Location = $User.Location

  # Calculate any licencing changes
  Switch($Licence) # None, E1 or E3
  {
    "None"
    {
      $Sku = "None"
      Write-Host "  - No licence is required."
#      Write-Host "  - Existing = $AccountSkuID ; New = $Sku"
      If ($AccountSkuID -eq $Sku)
      {
        Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
        Write-Host " - No changes required."
      }
      Else
      {
        Write-Host "  - Attempting to remove any existing licence."
        $Result = Set-MsolUserLicense -UserPrincipalName $UserName -RemoveLicenses $AccountSkuID
        If($Result -ne $null)
        { 
          Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
          Write-Host " Error removing existing licence."
          $ErrorCode = 8
          Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
        }
        Else
        { 
          Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
          Write-Host "   - Licence removed."
        }
      }
    }
    "E1"
    {
      $Sku = $AccountName + $E1
      Write-Host "  - $Licence licence for $Location has been requested."
#      Write-Host "  - Existing = $AccountSkuID ; New = $Sku"
      If ($AccountSkuID -eq $Sku)
      {
        Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
        Write-Host " - No changes required. No action taken for this user."
      }
      Else
      {
        Write-Host "  - Setting usage location to $Location."
        $Result = Set-MsolUser -UserPrincipalName $UserName -UsageLocation $Location
        If($Result -ne $null)
        { 
          Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
          Write-Host "   - Error setting usage location."
          $ErrorCode = 7
          Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
        }
        Else
        { 
          Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
          Write-Host "   - Usage location set."
        }
        Write-Host "  - Assigning $Sku licence."
        $Result = Set-MsolUserLicense -UserPrincipalName $UserName -AddLicenses $Sku
        If($Result -ne $null)
        { 
          Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
          Write-Host "   - Error adding licence."
          $ErrorCode = 8
          Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
        }
        Else
        { 
          Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
          Write-Host "   - Licence added."
        }
      }
    }
    "E3"
    {
      $Sku = $AccountName + $E3
      Write-Host "  - $Licence licence for $Location has been requested."
#      Write-Host "  - Existing = $AccountSkuID ; New = $Sku"
      If ($AccountSkuID -eq $Sku)
      {
        Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
        Write-Host " - No changes required. No action taken for this user."
      }
      Else
      {
        Write-Host "  - Setting Usage Location to $Location."
        $Result = Set-MsolUser -UserPrincipalName $UserName -UsageLocation $Location
        If($Result -ne $null)
        { 
          Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
          Write-Host "   - Error setting usage location."
          $ErrorCode = 7
          Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
        }
        Else
        { 
          Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
          Write-Host "   - Usage location set."
        }
        Write-Host "  - Assigning $Sku licence."
        $Result = Set-MsolUserLicense -UserPrincipalName $UserName -AddLicenses $Sku
        If($Result -ne $null)
        { 
          Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White"
          Write-Host "   - Error adding licence."
          $ErrorCode = 8
          Throw "Set-O365Licences There was a problem. Error code $ErrorCode."
        }
        Else
        { 
          Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
          Write-Host "   - Licence added."
        }
      }
    }
    default # Not recognised
    {
      Write-Host -NoNewLine "~" -BackgroundColor "Yellow" -ForegroundColor "Black"
      Write-Host " - Could not determine the licence requirements for $UserName. No action taken for this user."
    }
  }
}