# Define the output folder where the CSV file will be saved
$outputFolder = "C:\ExchangeMailboxReports"

# Specify the target Organizational Unit (OU)
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Replace with your actual OU

# Retrieve all sub-OUs of the target OU
$allOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $targetOU -SearchScope Subtree

# Empty list for all distribution groups across all OUs
$allDistributionGroups = @()

# Collect all distribution groups for each OU
foreach ($ou in $allOUs) {
    $ouGroups = Get-DistributionGroup -OrganizationalUnit $ou.DistinguishedName -ResultSize Unlimited
    $allDistributionGroups += $ouGroups
}

# Gather information about distribution groups, their members, permissions, and extended configuration
$allDistributionGroupsReport = foreach ($group in $allDistributionGroups) {
    $groupIdentity = $group.Alias  # Use the alias instead of the primary email address

    # Retrieve members of the distribution group with their email addresses
    try {
        $members = Get-DistributionGroupMember -Identity $groupIdentity | Where-Object { $_.Name -ne "DefaultPermission" } | ForEach-Object {
            $_.PrimarySmtpAddress
        }
    } catch {
        Write-Host "Error retrieving members for the group: $groupIdentity" -ForegroundColor Red
        continue
    }

    # Retrieve owners of the distribution group with their email addresses
    $owner = Get-DistributionGroup -Identity $groupIdentity | Select-Object -ExpandProperty ManagedBy | ForEach-Object {
        $ownerRecipient = Get-Recipient -Identity $_ -ErrorAction SilentlyContinue
        $ownerRecipient.PrimarySmtpAddress
    }

    # Retrieve additional SMTP addresses of the group
    $additionalSmtpAddresses = $group.EmailAddresses | Where-Object { $_.PrefixString -eq 'smtp' } | ForEach-Object { $_.AddressString }

    # Retrieve group configurations like membership approval, delivery management, message approval, etc.
    $config = Get-DistributionGroup -Identity $groupIdentity
    $membershipApproval = $config.MembersCanInviteOthers
    $deliveryManagement = $config.RequireSenderAuthenticationEnabled
    $messageApproval = $config.ModerationEnabled
    $emailOptions = $config.AcceptMessagesOnlyFromSendersOrMembers -join ", "
    $mailTip = $config.MailTip

    # Retrieve SendAs and SendOnBehalf permissions, ignore errors
    try {
        $sendAsPermissions = Get-ADPermission -Identity $groupIdentity -ErrorAction SilentlyContinue | Where-Object {
            ($_.ExtendedRights -like "*send*") -and -not ($_.User -match "NT AUTHORITY")
        }
    } catch {
        $sendAsPermissions = @()  # Ignore block if errors occur
    }

    # Results for each group with all relevant information
    [PSCustomObject]@{
        GroupName           = $group.DisplayName
        GroupEmail          = $group.PrimarySmtpAddress
        AdditionalSmtp      = $additionalSmtpAddresses -join ", "
        Owners              = $owner -join ", "
        Members             = $members -join ", "
        MembershipApproval  = $membershipApproval
        DeliveryManagement  = $deliveryManagement
        MessageApproval     = $messageApproval
        EmailOptions        = $emailOptions
        MailTip             = $mailTip
        SendAsRights        = $sendAsPermissions.ExtendedRights -join ", "
    }
}

# Export the collected data to a CSV file
$allDistributionGroupsReport | Export-Csv -Path "$outputFolder\DistributionGroupsExtendedReport.csv" -NoTypeInformation -Encoding UTF8

# Notify that the export was successful
Write-Host "Distribution groups extended report exported to $outputFolder\DistributionGroupsExtendedReport.csv"
