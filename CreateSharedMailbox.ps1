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

        $newPrimarySmtpAddress = "$firstPart@autoscout24.com"

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