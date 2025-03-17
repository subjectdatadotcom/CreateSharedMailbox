<#
.SYNOPSIS
This script automates the creation of shared mailboxes in Exchange Online using a CSV file.

.DESCRIPTION
The script first ensures that the `ExchangeOnlineManagement` module is installed and imported. It then connects to Exchange Online and reads mailbox details from a CSV file (`SharedMailboxes.csv`). 

For each entry in the CSV file, the script:
- Checks if the shared mailbox already exists.
- If it exists, logs it in a separate file (`SMB_Existing_SMBs.csv`).
- If it does not exist, creates a new shared mailbox with:
  - A modified display name prefixed with "TR-".
  - A primary SMTP address using the `targettenant.com` domain.
  - Optional settings such as archive and auto-expanding archive (based on the CSV data).

All failures encountered during mailbox creation are logged in `SMB_failures.csv`.

Once the process is complete, the script disconnects from Exchange Online.

.NOTES
- The script requires administrative privileges in Exchange Online.
- The `SharedMailboxes.csv` file must be in the same directory as the script.
- Output files (`SMB_Existing_SMBs.csv` and `SMB_failures.csv`) will be saved in the same directory.
- Ensure PowerShell execution policies allow the script to run.

.AUTHOR
SubjectData

.EXAMPLE
.\CreateSharedMailboxes.ps1
This will execute the script, processing the 'SharedMailboxes.csv' file, connecting to Exchange Online, and generating reports on created and existing mailboxes.
#>

 # Ensure the Exchange Management Shell module is loaded
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force
    Import-Module ExchangeOnlineManagement
} else {
    Import-Module ExchangeOnlineManagement
}

# Connect to Exchange Online, this might prompt for credentials
Connect-ExchangeOnline 

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$MyDir = "$myDir\"
$CSVPath = $MyDir + "SharedMailboxes.csv"

# Import mailboxes from a CSV file
$mailboxes = Import-Csv -Path $CSVPath

# Prepare a list to collect failed creations
$failures = @()
$alreadyExists = @()

# Create each mailbox
foreach ($mailbox in $mailboxes) {
    try 
    {
        $newDisplayName = "TR-" + $mailbox.DisplayName

        # Extract the local part before @ and append the new domain
        $firstPart = ($mailbox.UserPrincipalName -split "@")[0]

        $newPrimarySmtpAddress = "$firstPart@<targettenant>"

        # Check if mailbox already exists
        $existingMailbox = Get-Mailbox -Identity $newPrimarySmtpAddress -ErrorAction SilentlyContinue

        if ($existingMailbox) {
            # Log the existing mailbox
            $alreadyExists += [pscustomobject]@{
                UserPrincipalName = $mailbox.UserPrincipalName
                NewPrimarySmtpAddress = $newPrimarySmtpAddress
                DisplayName = $newDisplayName
                Status = "Already Exists"
            }

            continue
        }

        $params = @{
            Name = $newDisplayName
            DisplayName = $newDisplayName
            # Alias = $localPart
            PrimarySmtpAddress = $newPrimarySmtpAddress
            Shared = $true
        }

        # Set archive settings based on the CSV data
        if ($mailbox.ArchiveStatus -eq "Active") {
            $params['Archive'] = $true
        }

        if ($mailbox.AutoExpandingArchiveEnabled -eq "TRUE") {
            $params['AutoExpandingArchive'] = $true
        }

        # Attempt to create the mailbox with the specified parameters
        New-Mailbox @params
    } 
    catch {
        # Log the failure
        $failures += [pscustomobject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            NewPrimarySmtpAddress = $newPrimarySmtpAddress
            DisplayName = $newDisplayName
            Error = $_.Exception.Message
        }
    }
}

# Export the failures to a CSV file
$failures | Export-Csv -Path "$MyDir\SMB_failures.csv" -NoTypeInformation

# Export the failures to a CSV file
if ($failures.Count -gt 0) {
    $failures | Export-Csv -Path "$MyDir\SMB_failures.csv" -NoTypeInformation
}

# Export the already existing mailboxes to a CSV file
if ($alreadyExists.Count -gt 0) {
    $alreadyExists | Export-Csv -Path "$MyDir\SMB_Existing_SMBs.csv" -NoTypeInformation
}

# Disconnect from Exchange Online after operations are complete
Disconnect-ExchangeOnline -Confirm:$false
