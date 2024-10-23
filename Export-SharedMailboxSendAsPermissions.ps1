# Define the output folder where the CSV file will be saved
$outputFolder = "C:\ExchangeMailboxReports"

# Specify the target Organizational Unit (OU)
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your actual OU

# Get all shared mailboxes in the specified OU and its sub-OUs
$sharedMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'SharedMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Initialize an empty array to store SendAs permissions data
$allSendAsPermissions = foreach ($mailbox in $sharedMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    try {
        # Retrieve SendAs permissions, ignore errors if the object does not exist
        $sendAsPermissions = Get-ADPermission -Identity $mailboxIdentity -ErrorAction SilentlyContinue | Where-Object {
            ($_.ExtendedRights -like "*send*") -and -not ($_.User -match "NT AUTHORITY")
        }
    } catch {
        continue  # Ignore errors if the object does not exist
    }

    foreach ($permission in $sendAsPermissions) {
        $user = $permission.User

        # Ignore entries starting with "S-1-5-21"
        if ($user -notmatch "S-1-5-21*") {
            try {
                # Try to get the recipient's details and ensure it's a valid user mailbox
                $recipient = Get-Recipient -Identity $user -ErrorAction SilentlyContinue
                if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                    $email = $recipient.PrimarySmtpAddress
                } else {
                    continue  # Skip if it's not a valid user mailbox
                }
            } catch {
                continue  # Ignore orphaned entries or entries that don't resolve as recipients
            }

            # Store the results for each mailbox and its SendAs permission
            [PSCustomObject]@{
                Mailbox        = $mailbox.DisplayName
                MailboxEmail   = $mailbox.PrimarySmtpAddress
                User           = $email
                AccessRights   = $permission.ExtendedRights -join ", "
            }
        }
    }
}

# Export the collected data to a CSV file
$allSendAsPermissions | Export-Csv -Path "$outputFolder\SharedMailboxSendAsPermissions.csv" -NoTypeInformation -Encoding UTF8

# Notify that the export was successful
Write-Host "SendAs permissions exported to $outputFolder\SharedMailboxSendAsPermissions.csv"
