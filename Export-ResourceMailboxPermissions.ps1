# Define the output folder and the target OU for the resource mailboxes
$outputFolder = "C:\ExchangeMailboxReports"
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace this with your specific OU

# Retrieve all resource mailboxes (Room and Equipment) in the specified OU and its sub-OUs
$resourceMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Collect permissions for each resource mailbox and gather the email addresses of users with access
$allResourceMailboxPermissions = foreach ($mailbox in $resourceMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Use the alias instead of the primary email address

    # Retrieve all email addresses (primary and secondary) associated with the mailbox
    $emailAddresses = $mailbox.EmailAddresses | Where-Object { $_.PrefixString -eq 'smtp' } | ForEach-Object { $_.AddressString }

    # Retrieve permissions for the mailbox, excluding inherited permissions and system users (NT AUTHORITY)
    $mailboxPermissions = Get-MailboxPermission -Identity $mailboxIdentity | Where-Object {
        ($_.IsInherited -eq $False) -and -not ($_.User -match "NT AUTHORITY")
    }

    # Iterate through each permission entry and resolve the user's email
    foreach ($permission in $mailboxPermissions) {
        $user = $permission.User

        # Ignore any entries starting with "S-1-5-21" (representing system accounts or unresolved entries)
        if ($user -notmatch "S-1-5-21*") {
            try {
                # Attempt to resolve the user as a recipient and ensure they have a mailbox
                $recipient = Get-Recipient -Identity $user -ErrorAction SilentlyContinue
                if ($recipient -and $recipient.PrimarySmtpAddress) {
                    $email = $recipient.PrimarySmtpAddress
                } elseif ($recipient -and $recipient.UserPrincipalName) {
                    # If no primary SMTP address, use the UserPrincipalName (UPN) as a fallback
                    $email = $recipient.UserPrincipalName
                } else {
                    # If no email is found, use the domain\username format
                    $email = $user.ToString()
                }
            } catch {
                # If the user cannot be resolved, use the domain\username format as a fallback
                $email = $user.ToString()
            }

            # Output a custom object for each permission, including all email addresses and access rights
            [PSCustomObject]@{
                Mailbox        = $mailbox.DisplayName
                AllEmails      = $emailAddresses -join ", "  # Display all email addresses associated with the mailbox
                User           = $email                     # The resolved email or domain\user format
                AccessRights   = $permission.AccessRights -join ", "  # Join multiple access rights with commas
            }
        }
    }
}

# Export the collected data to a CSV file in the specified folder
$allResourceMailboxPermissions | Export-Csv -Path "$outputFolder\ResourceMailboxPermissionsWithEmails.csv" -NoTypeInformation -Encoding UTF8

# Notify that the script has successfully exported the data
Write-Host "Resource Mailbox permissions with all emails exported to $outputFolder\ResourceMailboxPermissionsWithEmails.csv"
