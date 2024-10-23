# Define the output folder where the CSV file will be saved
$outputFolder = "C:\ExchangeMailboxReports"

# Specify the target Organizational Unit (OU)
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your actual OU

# Get all shared mailboxes in the specified OU and its sub-OUs
$sharedMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'SharedMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Initialize an empty array to store mailbox permissions data
$allMailboxPermissions = foreach ($mailbox in $sharedMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    # Retrieve explicit mailbox permissions, excluding inherited permissions and system accounts
    $mailboxPermissions = Get-MailboxPermission -Identity $mailboxIdentity | Where-Object {
        ($_.IsInherited -eq $False) -and -not ($_.User -match "NT AUTHORITY")
    }

    foreach ($permission in $mailboxPermissions) {
        $user = $permission.User

        # Ignore entries starting with "S-1-5-21"
        if ($user -notmatch "S-1-5-21*") {
            # Try to resolve the user as a recipient
            try {
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                $email = if ($recipient -and $recipient.PrimarySmtpAddress) {
                    $recipient.PrimarySmtpAddress
                } else {
                    $user.ToString()  # If no email is found, use the username
                }
            } catch {
                $email = $user.ToString()  # If the user cannot be resolved, use the username
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
$allMailboxPermissions | Export-Csv -Path "$outputFolder\SharedMailboxPermissions.csv" -NoTypeInformation -Encoding UTF8

# Notify that the export was successful
Write-Host "Permissions exported to $outputFolder\SharedMailboxPermissions.csv"
