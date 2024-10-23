$outputFolder = "C:\ExchangeMailboxReports"
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your target OU

# Query all user mailboxes in the specified OU and its sub-OUs
$userMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'UserMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Query permissions for each user mailbox and collect email addresses of delegated users
$allSendOnBehalfPermissions = foreach ($mailbox in $userMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    $sendOnBehalfUsers = $mailbox.GrantSendOnBehalfTo

    foreach ($user in $sendOnBehalfUsers) {
        try {
            # Attempt to get the recipient and ensure it is a user mailbox
            $recipient = Get-Recipient -Identity $user -ErrorAction SilentlyContinue
            if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                $email = $recipient.PrimarySmtpAddress
            } else {
                continue  # Skip users who are not valid user mailboxes
            }
        } catch {
            continue  # Skip orphaned entries that are not valid recipients
        }

        # Result for each permission and mailbox
        [PSCustomObject]@{
            Mailbox        = $mailbox.DisplayName
            MailboxEmail   = $mailbox.PrimarySmtpAddress
            User           = $email
        }
    }
}

# Export the data to a CSV file
$allSendOnBehalfPermissions | Export-Csv -Path "$outputFolder\UserMailboxSendOnBehalfPermissions.csv" -NoTypeInformation -Encoding UTF8

Write-Host "SendOnBehalf permissions exported to $outputFolder\UserMailboxSendOnBehalfPermissions.csv"
