# Define the output folder where the CSV file will be saved
$outputFolder = "C:\ExchangeMailboxReports"

# Specify the target Organizational Unit (OU)
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your actual OU

# Get all shared mailboxes in the specified OU and its sub-OUs
$sharedMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'SharedMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Initialize an empty array to store SendOnBehalf permissions data
$allSendOnBehalfPermissions = foreach ($mailbox in $sharedMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    # Retrieve the users who have SendOnBehalf permissions for this mailbox
    $sendOnBehalfUsers = $mailbox.GrantSendOnBehalfTo

    # Iterate through each user with the permission
    foreach ($user in $sendOnBehalfUsers) {
        try {
            # Try to get the recipient's details and ensure it's a valid user mailbox
            $recipient = Get-Recipient -Identity $user -ErrorAction SilentlyContinue
            if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                $email = $recipient.PrimarySmtpAddress
            } else {
                continue  # Skip if it's not a valid user mailbox
            }
        } catch {
            continue  # Skip orphaned entries or entries that don't resolve as recipients
        }

        # Store the results for each mailbox and its SendOnBehalf permission
        [PSCustomObject]@{
            Mailbox        = $mailbox.DisplayName
            MailboxEmail   = $mailbox.PrimarySmtpAddress
            User           = $email
        }
    }
}

# Export the collected data to a CSV file
$allSendOnBehalfPermissions | Export-Csv -Path "$outputFolder\SharedMailboxSendOnBehalfPermissions.csv" -NoTypeInformation -Encoding UTF8

# Notify that the export was successful
Write-Host "SendOnBehalf permissions exported to $outputFolder\SharedMailboxSendOnBehalfPermissions.csv"
