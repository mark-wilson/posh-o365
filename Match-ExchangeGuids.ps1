Param(
    [String]$FileName,
    [String]$Tenant
  )

################################################################################
# Match-ExchangeGuids.ps1                                                      #
#                                                                              #
# Takes two parameters (CSV file containing a list of UPNs/Exchange Guids,     #
# Tenant name) and compares the current Exchange GUID on a MailUser with the   #
# supplied ExchangeGUID)                                                       #
#                                                                              #
# Requires the Microsoft Online Services module to be available on the system. #
# Also requires credentials to access the tenant.                              #
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
$Syntax = "SYNTAX: Match-ExchangeGuids CSVFile TenantName" # Syntax for this script
$UserList = @()                                            # Array used when reading list of users from $FileName
$ErrorCode = 0                                             # Used for error handling
$MSOLServiceUrl = "https://tenant.onmicrosoft.com"         # URL format for Microsoft Online Services (MSOL) tenants
$O365Creds = ""                                            # Used to store credentials for connection to MSOL
$Result = ""                                               # Used for error handling
$GetResult = ""                                            # Error handling on Get-MailUser
$SetResult = ""                                            # Error handling on Set-MailUser
$UserName = ""                                             # Username
$Action = ""                                               # Flag to record intended action
$CurrentExchangeGuid = ""                                  # ExchangeGuid from get-mailuser
$NewExchangeGuid = ""                                      # ExchangeGuid as recorded in CSV file
$ReadGuid = ""                                             # Temporary variable used to read ExchangeGuid
$Check = 0                                                 # Used in string comparison for GUIDs
$LogFile = ""                                              # Filename for logging

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
  Throw "Match-ExchangeGuids There was a problem. Error code $ErrorCode."
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
  Throw "Match-ExchangeGuids There was a problem. Error code $ErrorCode."
}

# All parameters seem to be valid and the pre-requisite PowerShell module exists
Write-Host "`nIntended action: Setting ExchangeGuid values for mailuser objects on $Tenant.onmicrosoft.com:"

Write-Host "  Attempting to connect to Exchange Online..."
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Creds -Authentication "Basic" -AllowRedirection
$Result = Import-PSSession $ExchangeSession -DisableNameChecking

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

# List details
Write-Host "  Processing " $UserList.Count "Mail User objects(s):"
Write-Host -NoNewLine "  Analysis phase" -ForegroundColor "Green"
Write-Host " - no changes will be made."

# Process each user in turn
ForEach ($User in $UserList)
{
  $UserName = $User.Upn
  $NewExchangeGuid = $User.msExchMailboxGuid
  #Write-Host "1" $NewExchangeGuid
  $NewExchangeGuid = $NewExchangeGuid -replace '[{}]',''
  $NewExchangeGuid = $NewExchangeGuid.ToUpper()
  #Write-Host "2" $NewExchangeGuid
  $GetResult = (Get-MailUser -Identity $UserName)

  If(!$GetResult)
  {
    $Action = "Error"
    Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White" 
    Write-Host " - $UserName : $Action (Unable to read ExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)" 
  }
  Else
  {
    $ReadGuid = $GetResult.ExchangeGuid
    $CurrentExchangeGuid = $ReadGuid.Guid
    $CurrentExchangeGuid = $CurrentExchangeGuid.ToUpper()
    #Write-Host "3" $CurrentExchangeGuid
    $Check = $NewExchangeGuid.CompareTo($CurrentExchangeGuid)
    If($NewExchangeGuid.CompareTo($CurrentExchangeGuid))
    {
      $Action = "Change"
      Write-Host -NoNewLine "~" -BackgroundColor "Yellow" -ForegroundColor "Black"
      Write-Host " - $UserName : $Action (Online ExchangeGuid is $CurrentExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)"
    }
    Else
    {
      $Action = "Match"
      Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
      Write-Host " - $UserName : $Action (Online ExchangeGuid is $CurrentExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)"
    }
  }

  $Action = ""
  $CurrentExchangeGuid = ""
  $NewExchangeGuid = ""
  $ReadGuid = ""
  $GetResult = ""
  $Check = 0
}

# Go ahead? Based on https://technet.microsoft.com/en-us/library/ff730939.aspx
$MenuTitle = "Change GUIDs?"
$MenuMessage = "Do you want to go ahead and make the changes?"
$MenuYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Makes the changes to ExchangeGuid on each Azure AD user object, as identified above."
$MenuNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Exits this script."
$Menuoptions = [System.Management.Automation.Host.ChoiceDescription[]]($MenuYes, $MenuNo)
$MenuResult = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0) 

Switch ($MenuResult)
{
  0 {

# Do it for real now...

# Logging thanks to http://sharepointjack.com/2013/simple-powershell-script-logging/
$LogFile = "$env:USERPROFILE\Match-ExchangeGuids_$(Get-Date -Format `"yyyyMMdd_HHmmss`").txt" 

Write-Host -NoNewLine "`n  Action phase" -ForegroundColor "Yellow"
Write-Host " - updating GUIDs (logging to $LogFile)."

"Match-ExchangeGuids.ps1" | Out-File -Filepath $LogFile -Append
"Called with parameters $FileName $Tenant" | Out-File -Filepath $LogFile -Append
"Logging to $LogFile" | Out-File -Filepath $LogFile -Append

ForEach ($User in $UserList)
{
  $UserName = $User.Upn
  $NewExchangeGuid = $User.msExchMailboxGuid
  #Write-Host "1" $NewExchangeGuid
  $NewExchangeGuid = $NewExchangeGuid -replace '[{}]',''
  $NewExchangeGuid = $NewExchangeGuid.ToUpper()
  #Write-Host "2" $NewExchangeGuid
  
  "$(Get-Date -Format `"yyyyMMdd_HHmmss`"): Calling Get-MailUser -Identity $UserName" | Out-File -Filepath $LogFile -Append
  
  $GetResult = (Get-MailUser -Identity $UserName)

  If(!$GetResult)
  {
    $Action = "Error"
    "*** ERROR: Unable to read Exchange GUID ***" | Out-File -Filepath $LogFile -Append
    "($UserName may not exist, or may have a mailbox already - this script only looks for mail attributes on users without mailboxes)" | Out-File -Filepath $LogFile -Append
    Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White" 
    Write-Host " - $UserName : $Action (Unable to read ExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)" 
  }
  Else
  {
    $ReadGuid = $GetResult.ExchangeGuid
    $CurrentExchangeGuid = $ReadGuid.Guid
    $CurrentExchangeGuid = $CurrentExchangeGuid.ToUpper()
    #Write-Host "3" $CurrentExchangeGuid
    $Check = $NewExchangeGuid.CompareTo($CurrentExchangeGuid)
    If($NewExchangeGuid.CompareTo($CurrentExchangeGuid))
    {
      $Action = "Change"
      "Attempting to change ExchangeGUID from $CurrentExchangeGuid to $NewExchangeGuid" | Out-File -Filepath $LogFile -Append
      Write-Host -NoNewLine "~" -BackgroundColor "Yellow" -ForegroundColor "Black"
      Write-Host " - $UserName : $Action (Online ExchangeGuid is $CurrentExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)"
      "$(Get-Date -Format `"yyyyMMdd_HHmmss`"): Calling Set-MailUser -Identity $UserName -ExchangeGuid $NewExchangeGuid" | Out-File -Filepath $LogFile -Append
      $SetResult = (Set-MailUser -Identity $UserName -ExchangeGuid $NewExchangeGuid)

      If($SetResult)
      {
        $Action = "Error"
        "*** ERROR: Unable to set Exchange GUID ***" | Out-File -Filepath $LogFile -Append
        Write-Host -NoNewLine "x" -BackgroundColor "Red" -ForegroundColor "White" 
        Write-Host " - $UserName : $Action (Unable to set ExchangeGuid)" 
      }
      Else
      {
        "ExchangeGUID changed to $NewExchangeGuid" | Out-File -Filepath $LogFile -Append
        Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
        Write-Host "   Changed successfully"
      }
    }
    Else
    {
      $Action = "Match"
      "No need to change ExchangeGUID attribute - $CurrentExchangeGuid matches $NewExchangeGuid" | Out-File -Filepath $LogFile -Append
      Write-Host -NoNewLine " " -BackgroundColor "Green" -ForegroundColor "White"
      Write-Host " - $UserName : $Action (Online ExchangeGuid is $CurrentExchangeGuid; Intended ExchangeGuid is $NewExchangeGuid)"
    }
  }

  $Action = ""
  $CurrentExchangeGuid = ""
  $NewExchangeGuid = ""
  $ReadGuid = ""
  $GetResult = ""
  $SetResult = ""
  $Check = 0
}
}
 1 {"Exiting as requested."}
}