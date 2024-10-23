$outputFolder = "C:\ExchangeMailboxReports"
$targetOU = "OU=ZielOU,DC=deineDomain,DC=com"  # Ersetze dies durch deine OU

# Alle User Mailboxen in der angegebenen OU und deren Sub-OUs abfragen
$userMailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq 'UserMailbox'} -OrganizationalUnit $targetOU -ResultSize Unlimited

# Berechtigungen für jede User Mailbox abfragen und E-Mail-Adressen der berechtigten Benutzer sammeln
$allSendOnBehalfPermissions = foreach ($mailbox in $userMailboxes) {
    $mailboxIdentity = $mailbox.Alias  # Verwende den Alias anstelle der Primären E-Mail-Adresse

    $sendOnBehalfUsers = $mailbox.GrantSendOnBehalfTo

    foreach ($user in $sendOnBehalfUsers) {
        try {
            # Versuche, den Benutzer als Empfänger zu erhalten und sicherstellen, dass es sich um eine User-Mailbox handelt
            $recipient = Get-Recipient -Identity $user -ErrorAction SilentlyContinue
            if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                $email = $recipient.PrimarySmtpAddress
            } else {
                continue  # Ignoriere Benutzer, die keine gültige User-Mailbox haben
            }
        } catch {
            continue  # Ignoriere verwaiste Einträge, die keine Empfänger sind
        }

        # Ergebnis für jede Berechtigung und Mailbox
        [PSCustomObject]@{
            Mailbox        = $mailbox.DisplayName
            MailboxEmail   = $mailbox.PrimarySmtpAddress
            User           = $email
        }
    }
}

# Exportiere die Daten in eine CSV-Datei
$allSendOnBehalfPermissions | Export-Csv -Path "$outputFolder\UserMailboxSendOnBehalfPermissions.csv" -NoTypeInformation -Encoding UTF8

Write-Host "SendOnBehalf permissions exported to $outputFolder\UserMailboxSendOnBehalfPermissions.csv"
