# Define the output folder where the CSV file will be saved
$outputFolder = "C:\ExchangeMailboxReports"

# Specify the target Organizational Unit (OU)
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your actual OU

# Get all user mailboxes in the specified OU and its sub-OUs
$userMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'UserMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Initialize an empty array to store mailbox permissions data
$allMailboxPermissions = foreach ($mailbox in $userMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    # Retrieve explicit mailbox permissions, excluding inherited permissions and system accounts
    $mailboxPermissions = Get-MailboxPermission -Identity $mailboxIdentity | Where-Object {
        ($_.IsInherited -eq $False) -and -not ($_.User -match "NT AUTHORITY")
    }

    foreach ($permission in $mailboxPermissions) {
        $user = $permission.User

        # Ignore entries starting with "S-1-5-21"
        if ($user -notmatch "S-1-5-21*") {
            try {
                # Try to get the recipient's details and ensure it's a valid user mailbox
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                    $email = $recipient.PrimarySmtpAddress
                } else {
                    continue  # Skip if it's not a valid user mailbox
                }
            } catch {
                continue  # Ignore orphaned entries or entries that don't resolve as recipients
            }

            # Store the results for each mailbox and its permissions
            [PSCustomObject]@{
                Mailbox        = $mailbox.DisplayName
                MailboxEmail   = $mailbox.PrimarySmtpAddress
                User           = $email
                AccessRights   = $permission.AccessRights -join ", "
            }
        }
    }
}

# Export the collected data to a CSV file
$allMailboxPermissions | Export-Csv -Path "$outputFolder\UserMailboxPermissions.csv" -NoTypeInformation -Encoding UTF8

# Notify that the export was successful
Write-Host "Permissions exported to $outputFolder\UserMailboxPermissions.csv"
