[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param (
    # These parameters are required when running in either interactive or non-interactive mode
    [Parameter(Mandatory = $true, ParameterSetName = "Interactive")]
    [Parameter(Mandatory = $true, ParameterSetName = "NonInteractive")]
    [string]$Path = "login.csv", # The path to the CSV file. It MUST have two fields in it, login and tid
    [string]$UFOURL, # The URL to your UniFLOW Online Instance
    [string]$Domain, # Your email domain (e.g. example.com) which is used to match users
    [string]$IdentityType = "CardNumber", # Which identity type you wish to update. Defaults to CardNumber
    [string]$LogFile = "results.log", # Where to write the results of the import

    # This parameter is required when in interactive mode (default)
    [Parameter(Mandatory = $true, ParameterSetName = "Interactive")]
    [switch]$Interactive, # If true, you'll be asked to log in to UniFLOW Online 

    # These two parameters are only required if using non-interactive mode. 
    [Parameter(Mandatory = $true, ParameterSetName = "NonInteractive")]
    [string]$UFOClientID = $Null, # The UniFLOW Online Client ID
    [string]$UFOClientSecret = $Null # The UniFLOW Online Secret
)

# Try and import the module. If that doesn't work (because it's not installed), exit with an error
try {
    Import-Module NTware.Ufo.PowerShell.ObjectManagement
}
catch {
    Write-Error -Message "UniFLOW Powershell module not available. Install using Install-Module NTware.Ufo.PowerShell.ObjectManagement" -Category NotInstalled
    Exit 1
}

# First, check if the CSV file exists
if (-not (Test-Path $Path)) {
    Write-Error -Message "The specified CSV file does not exist." -Category ResourceUnavailable
    Exit 1
}

# Then we read the CSV file. We need to cast it as an array, otherwise
# the .count property breaks if there's only one entry
[array]$data = Import-Csv -Path $Path

# These are the columns that are required in our CSV file
$requiredColumns = @("login", "tid")
# Then we get the header names out of the CSV file
$csvColumns = ($data | Get-Member -MemberType NoteProperty).Name

# If the columns in the CSV file don't exactly match our required columns, exit
if ((Compare-Object -ReferenceObject $requiredColumns -DifferenceObject $csvColumns).Count -gt 0) {
    Write-Error -Message "CSV file must contain 'login' and 'tid' columns (in that order)" -Category InvalidData
    Exit 1
}

try {
    # If -Interactive has been passed, ignore the Client ID and Client Secret and open a window to log in
    # If it's not been set, grab the credentials from the command line
    if ($Interactive) {
        Open-MomoConnection -TenantDomain $UFOURL -Interactive
    }
    else {
        # Make a secure string out of our client secret
        $secStringPassword = ConvertTo-SecureString $UFOClientSecret -AsPlainText -Force
        # And make a new credentials object that we'll pass to the UniFLOW Online powershell cmdlet
        $credObject = New-Object System.Management.Automation.PSCredential($UFOClientID, $secStringPassword)
        # Now actually log in 
        Open-MomoConnection -TenantDomain $UFOURL -NonInteractiveUserApplication $credObject
    }
   
}
catch {
    Write-Error -Message "Error Logging in to UniFLOW Online. Error returned was $_"
    Exit 1
}

# How many users are in this file
$total = $data.Count

# How many users we've processed so far
$count = 0

# How many failures there's been 
$failures = 0

# How many successes there's been
$successes = 0

# Make an empty table. 
[PSCustomObject]$resultsTable = @()

# Loop through every row 
foreach ($row in $data) {
    # Increment our count
    $count++
    # Get the username
    $login = $row.login
    # And get the card number
    $tid = $row.tid

    # The TID must be an 8 character hex string
    if (-not ($tid -Match '^[0-9a-fA-F]{8}$')) {
        $resultsTable += @{
            Login = $login;
            TID = $tid;
            Result = "Incorrect Card Number Format"
        }
        $failures++
        Continue
    }

    # The username must be between 3 and 7 characters long and must end with a digit
    if (-not ($login -Match '^[a-zA-Z]{3,7}\d$')) {
        $resultsTable += @{
            Login = $login;
            TID = $tid;
            Result = "Incorrect Username Format"
        }
        $failures++
        Continue
    }
    
    # Update progress bar
    Write-Progress -Activity "Processing CSV" -Status "Updating login for '$login@$Domain' with $tid" -PercentComplete (($count / $total) * 100)
    
    try {
        Add-MomoUserIdentity -Email "$login@$Domain" -IdentityType $IdentityType -IdentityValue "$tid" | Out-Null
        $resultsTable += @{
            Login = $login;
            TID = $tid;
            Result = "Success"
        }
        $successes++
    }
    catch {
        $resultsTable += @{
            Login = $login;
            TID = $tid;
            Result = "Error Processing: $_"
        }
        $failures++
        Continue
    }
}

# Spit out the results of the table we made. This is written to log.txt
$resultsTable | ForEach-Object {[PSCustomObject]$_} | Format-Table -AutoSize Login, TID, Result | Out-File -FilePath $LogFile

# Then we count up all the users, the successful imports and the failed imports and output the table to the screen.
@{ Total = $total; Successful = $successes; Failures = $failures } | ForEach-Object {[PSCustomObject]$_} | Format-Table -AutoSize Total, Successful, Failures
Write-Host "Details of import written to $LogFile"