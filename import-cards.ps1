[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(Mandatory = $true, ParameterSetName = "Interactive")]
    [Parameter(Mandatory = $true, ParameterSetName = "NonInteractive")]
    [string]$Path = "login.csv", # The path to the CSV file. It MUST have two fields in it, login and tid
    [string]$UFOURL, # The URL to your UniFLOW Online Instance
    [string]$Domain, # Your email domain (e.g. example.com) which is used to match users

    [Parameter(Mandatory = $true, ParameterSetName = "Interactive")]
    [switch]$Interactive, # If true, you'll be asked to log in to UniFLOW Online 

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

# Loop through every row 
foreach ($row in $data) {
    # Increment our count
    $count++
    # Get the username
    $login = $row.login
    # And get the card number
    $tid = $row.tid

    # The TID must be an 8 character hex string
    if (-not $tid -Match '^[0-9a-fA-F]{8}$') {
        Write-Error -Message "Card number for $login on row $count is not valid hexadecimal: $tid" -Category InvalidData
        Continue
    }

    # The username must be between 3 and 7 characters long and must end with a digit
    if (-not $login -Match '^[a-zA-Z]{3,7}\d$') {
        Write-Error -Message "Username $login on row $count does not match expected format!" -Category InvalidData
        Continue
    }
    
    # Update progress bar
    Write-Progress -Activity "Processing CSV" -Status "Updating login for '$login@$Domain' with $tid" -PercentComplete (($count / $total) * 100)
    
    try {
        Add-MomoUserIdentity -Email "$login@$Domain" -IdentityType 'CardNumber' -IdentityValue "$tid"
    }
    catch {
        Write-Error "Failed to process login: $login@$Domain with tid: $tid - $_"
        Continue
    }
}

Write-Host "Processing complete."
