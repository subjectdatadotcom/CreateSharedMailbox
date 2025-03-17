# Shared Mailbox Creation Script for Microsoft 365

This PowerShell script automates the creation of **shared mailboxes** in a **Microsoft 365 (M365) tenant** by reading data from a CSV file. It ensures that mailboxes are only created if they do not already exist and logs any failures or already existing mailboxes.

## Features
- Reads mailbox details from a CSV file (`SharedMailboxes.csv`).
- Prefixes **"TR-"** (or configurable prefix) to the `DisplayName` while keeping the same email alias.
- Updates the `PrimarySmtpAddress` domain to match the target M365 tenant.
- Checks if a mailbox already exists before attempting creation.
- Logs failures and already existing mailboxes in separate CSV files.
- Ensures archive settings are applied based on CSV data.

## Prerequisites
Before running the script, ensure:
1. You have **Exchange Online Management Module** installed.
2. You have admin credentials for your **Microsoft 365 Exchange Online** tenant.
3. The CSV file (`SharedMailboxes.csv`) is formatted correctly with the following columns:
   - `UserPrincipalName` (original email address)
   - `DisplayName` (original display name)
   - `ArchiveStatus` (`Active` or `None`)
   - `AutoExpandingArchiveEnabled` (`TRUE` or `FALSE`)   
4. Replace `<targettenant>` with your tenant, for example: `contoso.microsoft.com`.

## Installation
If the **Exchange Online Management Module** is not installed, run:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force
```
Usage
Place your CSV file SharedMailboxes.csv in the same directory as the script.

Open PowerShell as an administrator and run the script:
./CreateSharedMailbox.ps1

If prompted, sign in with your Exchange Online Admin credentials.

## Output

- Success: Mailboxes are created successfully.
- Failures: Logged in SMB_failures.csv with error details.
- Already Existing Mailboxes: Logged in SMB_Existing_SMBs.csv.


## Notes
- Ensure you have sufficient admin permissions to create mailboxes.
- The script automatically checks for existing mailboxes before creation.
- The script will prompt for Exchange Online admin credentials.
