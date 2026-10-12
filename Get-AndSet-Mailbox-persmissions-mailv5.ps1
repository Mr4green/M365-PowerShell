# ===============================
# Mailbox Permissions v5
# Exchange Online focused
# ===============================
# Q (of 'exit') is op elke vraag te gebruiken:
#  - in het hoofdmenu en bij de startvragen: script afsluiten
#  - tijdens een actie: actie annuleren en terug naar het menu
# ===============================

# -------------------------------
# Logging Setup
# -------------------------------
$LogFolder = "C:\ps1\Logs"
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$script:LogFile    = "$($LogFolder)\MailboxPermissions-v5-$($timestamp).log"
$script:Actions    = New-Object System.Collections.Generic.List[string]
$script:LastAction = "Geen wijzigingen doorgevoerd."

# State (wordt gevuld door Select-Mailbox / Show-Permissions)
$script:mailbox             = $null
$script:calendarName        = $null
$script:contactsName        = $null
$script:calendarSharedWith  = @()

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $script:LogFile -Value $entry
}

Write-Log "=== Script Started (v5) ==="

# -------------------------------
# Helper functions
# -------------------------------

# Leest invoer. Q / exit / quit gooit een 'USER_EXIT' die door de aanroeper wordt afgevangen.
function Read-Input {
    param([string]$Prompt)
    $value = Read-Host "$Prompt (Q = stoppen)"
    if ($null -eq $value) { $value = "" }
    $value = $value.Trim()
    if ($value -match '^(q|exit|quit)$') {
        Write-Log "User typed exit at prompt: $Prompt"
        throw "USER_EXIT"
    }
    return $value
}

# Blijft vragen tot de invoer in de lijst met geldige opties staat.
function Read-Option {
    param(
        [string]$Prompt,
        [string[]]$Valid
    )
    while ($true) {
        $v = (Read-Input $Prompt).ToLower()
        if ($Valid -contains $v) { return $v }
        Write-Host "Ongeldige keuze. Kies uit: $($Valid -join ', ')" -ForegroundColor Red
    }
}

# Bevestiging voor elke wijziging.
function Confirm-Action {
    param([string]$Description)
    Write-Host "`nJe staat op het punt om: $Description" -ForegroundColor Yellow
    $answer = Read-Option "Doorvoeren? (J/N)" @('j','ja','y','yes','n','nee','no')
    $ok = $answer -in @('j','ja','y','yes')
    Write-Log "Confirmation for '$Description': $answer"
    if (-not $ok) { Write-Host "Geannuleerd, er is niets gewijzigd." -ForegroundColor Yellow }
    return $ok
}

# Registreert een uitgevoerde actie in log, samenvatting en console.
function Register-Action {
    param([string]$Text)
    $script:LastAction = $Text
    $script:Actions.Add($Text)
    Write-Log $Text
    Write-Host "`n$Text" -ForegroundColor Green
}

# Vraagt een gebruiker/groep en valideert die met Get-Recipient.
function Read-TargetUser {
    param([string]$Prompt = "Gebruiker (user@domain.com)")
    while ($true) {
        $u = Read-Input $Prompt
        if (-not $u) { continue }
        try {
            $found = @(Get-Recipient -Identity $u -ErrorAction Stop)
            if ($found.Count -ne 1) {
                Write-Host "'$u' levert $($found.Count) resultaten op. Wees specifieker (bijv. volledig e-mailadres)." -ForegroundColor Red
                continue
            }
            $resolved = $found[0].PrimarySmtpAddress.ToString()
            Write-Log "Target user validated: $u -> $resolved"
            Write-Host "Gevonden: $($found[0].DisplayName) <$resolved>" -ForegroundColor Green
            return $resolved
        }
        catch {
            Write-Host "Ontvanger '$u' niet gevonden. Probeer opnieuw." -ForegroundColor Red
            Write-Log "Target user not found: $u - $($_.Exception.Message)"
        }
    }
}

# Toont een genummerde lijst en geeft het gekozen item terug.
function Select-Entry {
    param(
        $Items,
        [string]$Prompt = "Kies een nummer"
    )
    $Items = @($Items)
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host ("{0} = {1}" -f ($i + 1), $Items[$i].Display)
    }
    while ($true) {
        $c = Read-Input $Prompt
        if ($c -match '^\d+$' -and [int]$c -ge 1 -and [int]$c -le $Items.Count) {
            return $Items[[int]$c - 1]
        }
        Write-Host "Ongeldige keuze." -ForegroundColor Red
    }
}

# Bestaande verbinding ophalen (EXO V3 gebruikt geen PSSessions meer).
function Get-ActiveExoConnection {
    try {
        return Get-ConnectionInformation -ErrorAction Stop |
            Where-Object { $_.State -eq "Connected" -and $_.TokenStatus -eq "Active" } |
            Select-Object -First 1
    }
    catch {
        return $null
    }
}

function Connect-Exo {
    $choice = Read-Input "Wil je verbinding maken met Exchange Online? (ja/nee)"
    Write-Log "Connect choice: $choice"

    $existing = Get-ActiveExoConnection

    if ($choice -match '^(ja|j|yes|y)$') {
        if ($existing) {
            Write-Host "Bestaande Exchange Online-sessie gevonden. Deze wordt hergebruikt." -ForegroundColor Green
            Write-Log "Existing EXO connection reused: $($existing.UserPrincipalName)"
            return
        }
    }
    else {
        if ($existing) {
            Write-Log "User chose not to connect. Existing connection found: $($existing.UserPrincipalName)"
            return
        }
        Write-Host "Geen actieve Exchange Online-sessie gevonden. Er wordt alsnog verbonden." -ForegroundColor Yellow
        Write-Log "User chose not to connect, but no active session exists. Connecting anyway."
    }

    try {
        Write-Host "Verbinden met Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online."
    }
    catch {
        Write-Host "Kon geen verbinding maken met Exchange Online." -ForegroundColor Red
        Write-Log "ERROR connecting to Exchange Online: $($_.Exception.Message)"
        throw
    }
}

# Vraagt mailbox, valideert, haalt mappen op en detecteert Calendar/Contacts.
function Select-Mailbox {
    $possibleCalendarNames = @("Calendar","Agenda","Calendrier","Calendario","Kalender","Kalendar","Terminar")
    $possibleContactsNames = @("Contacts","Contactpersonen")

    while ($true) {
        $input = Read-Input "Mailbox (bijv. user@domain.com)"
        Write-Log "Mailbox entered: $input"

        # Validatie
        try {
            $mailboxObject = Get-Mailbox -Identity $input -ErrorAction Stop
            Write-Log "Mailbox validated successfully: $input"
        }
        catch {
            Write-Host "Mailbox '$($input)' niet gevonden of niet toegankelijk. Probeer opnieuw." -ForegroundColor Red
            Write-Log "ERROR validating mailbox '$($input)': $($_.Exception.Message)"
            continue
        }

        # Mappen ophalen
        Write-Host "`n=== Maplijst ophalen voor $($input) ===" -ForegroundColor Cyan
        Write-Log "Fetching mailbox folder list..."
        try {
            $folders = Get-EXOMailboxFolderStatistics -Identity $input -ErrorAction Stop
            Write-Log "Folder list collected. Folder count: $($folders.Count)"
        }
        catch {
            Write-Host "Kon de mappen van '$($input)' niet ophalen. Probeer een andere mailbox." -ForegroundColor Red
            Write-Log "ERROR retrieving mailbox folders: $($_.Exception.Message)"
            continue
        }

        $script:mailbox       = $mailboxObject.PrimarySmtpAddress.ToString()
        $script:mailboxObject = $mailboxObject

        $folders |
            Select-Object Name, FolderPath, FolderType |
            Sort-Object FolderPath |
            Format-Table -AutoSize | Out-Host

        # Calendar (kortste pad = standaardagenda)
        $calendarFolder = $folders | Where-Object { $_.FolderType -eq "Calendar" } |
            Sort-Object { $_.FolderPath.Length } | Select-Object -First 1
        if (-not $calendarFolder) {
            $calendarFolder = $folders | Where-Object { $_.Name -in $possibleCalendarNames } | Select-Object -First 1
        }
        if ($calendarFolder) {
            $script:calendarName = $calendarFolder.Name
            Write-Host "Gedetecteerde Calendar/Agenda-map: $($script:calendarName)" -ForegroundColor Green
            Write-Log "Detected Calendar folder: $($script:calendarName)"
        }
        else {
            Write-Host "Calendar-map niet automatisch gevonden." -ForegroundColor Yellow
            Write-Log "Calendar autodetect failed. Asking user input."
            $script:calendarName = Read-Input "Voer de naam van de Calendar-map in"
            Write-Log "User entered Calendar folder: $($script:calendarName)"
        }

        # Contacts
        $contactsFolder = $folders | Where-Object { $_.FolderType -eq "Contacts" } |
            Sort-Object { $_.FolderPath.Length } | Select-Object -First 1
        if (-not $contactsFolder) {
            $contactsFolder = $folders | Where-Object { $_.Name -in $possibleContactsNames } | Select-Object -First 1
        }
        if ($contactsFolder) {
            $script:contactsName = $contactsFolder.Name
            Write-Host "Gedetecteerde Contacts-map: $($script:contactsName)" -ForegroundColor Green
            Write-Log "Detected Contacts folder: $($script:contactsName)"
        }
        else {
            Write-Host "Contacts-map niet automatisch gevonden." -ForegroundColor Yellow
            Write-Log "Contacts autodetect failed. Asking user input."
            $script:contactsName = Read-Input "Voer de naam van de Contacts-map in"
            Write-Log "User entered Contacts folder: $($script:contactsName)"
        }

        Write-Log "Calendar identity resolved as: $($script:mailbox):\$($script:calendarName)"
        Write-Log "Contacts identity resolved as: $($script:mailbox):\$($script:contactsName)"
        return
    }
}

function Show-MailboxDetails {
    Write-Host "`n=== Mailbox details voor $($script:mailbox) ===" -ForegroundColor Cyan
    Write-Log "Retrieving mailbox details for $($script:mailbox)"
    try {
        Get-Mailbox -Identity $script:mailbox -ErrorAction Stop |
            Format-List DisplayName,Name,Alias,PrimarySmtpAddress,WindowsEmailAddress,EmailAddresses,UserPrincipalName,ExchangeGuid |
            Out-Host
        Write-Log "Mailbox details retrieved successfully."
    }
    catch {
        Write-Warning "Kon mailbox details niet ophalen voor $($script:mailbox)."
        Write-Log "ERROR retrieving mailbox details: $($_.Exception.Message)"
    }
}

function Show-Permissions {
    $calendarIdentity = "$($script:mailbox):\$($script:calendarName)"
    $contactsIdentity = "$($script:mailbox):\$($script:contactsName)"

    # Calendar
    Write-Host "`n=== Alle rechten voor $($calendarIdentity) ===" -ForegroundColor Cyan
    Write-Log "Retrieving Calendar permissions for $($calendarIdentity)"
    $script:calendarSharedWith = @()
    try {
        $calendarPermissions = @(Get-EXOMailboxFolderPermission -Identity $calendarIdentity -ErrorAction Stop)
        $calendarPermissions | Format-Table -AutoSize | Out-Host
        Write-Log "Calendar permissions retrieved successfully."

        $script:calendarSharedWith = @($calendarPermissions |
            Where-Object { $_.User.ToString() -notin @("Default","Anonymous") })
    }
    catch {
        Write-Warning "Kon Calendar-rechten niet ophalen voor $($calendarIdentity)."
        Write-Log "ERROR retrieving Calendar permissions: $($_.Exception.Message)"
    }

    Write-Host "`n=== Gedeeld met (alleen Calendar, zonder Default/Anonymous) ===" -ForegroundColor Cyan
    if ($script:calendarSharedWith.Count -gt 0) {
        $script:calendarSharedWith | Select-Object User, AccessRights | Format-Table -AutoSize | Out-Host
        Write-Log "Calendar explicit shares found: $($script:calendarSharedWith.Count)"
    }
    else {
        Write-Host "Geen expliciete agenda-shares gevonden." -ForegroundColor Yellow
        Write-Log "No explicit calendar shares found."
    }

    # Contacts
    Write-Host "`n=== Rechten voor $($contactsIdentity) ===" -ForegroundColor Cyan
    Write-Log "Retrieving Contacts permissions for $($contactsIdentity)"
    try {
        Get-EXOMailboxFolderPermission -Identity $contactsIdentity -ErrorAction Stop | Format-Table -AutoSize | Out-Host
        Write-Log "Contacts permissions retrieved successfully."
    }
    catch {
        Write-Warning "Kon Contacts-rechten niet ophalen voor $($contactsIdentity)."
        Write-Log "ERROR retrieving Contacts permissions: $($_.Exception.Message)"
    }

    # Root mailbox
    Write-Host "`n=== Root mailbox rechten voor $($script:mailbox) ===" -ForegroundColor Cyan
    Write-Log "Retrieving root mailbox permissions for $($script:mailbox)"
    try {
        Get-EXOMailboxPermission -Identity $script:mailbox -ErrorAction Stop |
            Select-Object User, AccessRights, Deny, IsInherited |
            Format-Table -AutoSize | Out-Host
        Write-Log "Root mailbox permissions retrieved successfully."
    }
    catch {
        Write-Warning "Kon root mailbox-rechten niet ophalen voor $($script:mailbox)."
        Write-Log "ERROR retrieving root mailbox permissions: $($_.Exception.Message)"
    }
}

function Export-CalendarShares {
    if (-not $script:mailbox) { return }

    if (-not $script:calendarSharedWith -or $script:calendarSharedWith.Count -eq 0) {
        Write-Log "No calendar shares to export for $($script:mailbox)."
        return
    }

    $safeName = $script:mailbox -replace '[^\w@.\-]', '_'
    $csvPath  = "$($LogFolder)\CalendarSharedWith-$($safeName)-$($timestamp).csv"

    try {
        $mb = $script:mailbox
        $cf = $script:calendarName
        $script:calendarSharedWith |
            Select-Object @{ Name = "Mailbox"; Expression = { $mb } },
                          @{ Name = "Folder";  Expression = { $cf } },
                          @{ Name = "User";    Expression = { $_.User.ToString() } },
                          @{ Name = "AccessRights"; Expression = { $_.AccessRights -join ";" } } |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        Write-Host "`nCalendar shared-with export opgeslagen: $($csvPath)" -ForegroundColor Cyan
        Write-Log "Calendar shared-with export saved to: $csvPath"
    }
    catch {
        Write-Warning "Kon calendar shared-with export niet opslaan."
        Write-Log "ERROR exporting calendar shared-with results: $($_.Exception.Message)"
    }
}

function Show-Menu {
    Write-Host "`n==================== MENU ($($script:mailbox)) ====================" -ForegroundColor Yellow
    Write-Host "1 = Mailbox rechten toevoegen (FullAccess / ReadPermission)"
    Write-Host "2 = Mailbox rechten verwijderen"
    Write-Host "3 = Mailbox FullAccess van 1 persoon resetten (verwijderen + opnieuw toekennen)"
    Write-Host "4 = Calendar/Agenda rechten toevoegen of wijzigen"
    Write-Host "5 = Calendar/Agenda rechten verwijderen"
    Write-Host "6 = Contacts rechten toevoegen of wijzigen"
    Write-Host "7 = Contacts rechten verwijderen"
    Write-Host "8 = Overzicht verversen"
    Write-Host "9 = Andere mailbox kiezen"
    Write-Host "Q = Afsluiten"
}

# -------------------------------
# Main
# -------------------------------
try {
    Connect-Exo
    Select-Mailbox
    Show-MailboxDetails
    Show-Permissions

    while ($true) {
        Show-Menu
        $sel = (Read-Host "Keuze").Trim()
        Write-Log "Menu selection: $sel"

        if ($sel -match '^(q|exit|quit)$') {
            Write-Log "User chose to exit from main menu."
            break
        }

        $calendarIdentity = "$($script:mailbox):\$($script:calendarName)"
        $contactsIdentity = "$($script:mailbox):\$($script:contactsName)"
        $changed = $false

        try {
            switch ($sel) {

                # ---------- Mailbox: toevoegen ----------
                "1" {
                    $target = Read-TargetUser "Gebruiker die rechten krijgt (user@domain.com)"
                    $type   = Read-Option "Kies: 1 = FullAccess, 2 = ReadPermission" @('1','2')

                    if ($type -eq '1') {
                        $am = Read-Option "AutoMapping inschakelen (mailbox automatisch in Outlook tonen)? (J/N)" @('j','n')
                        $autoMap = ($am -eq 'j')

                        if (Confirm-Action "FullAccess (AutoMapping: $autoMap) toekennen aan $target op $($script:mailbox)") {
                            Add-MailboxPermission -Identity $script:mailbox -User $target -AccessRights FullAccess -InheritanceType All -AutoMapping $autoMap -ErrorAction Stop | Out-Null
                            Register-Action "Granted FullAccess (AutoMapping: $autoMap) to $target on $($script:mailbox)"
                            $changed = $true
                        }
                    }
                    else {
                        if (Confirm-Action "ReadPermission toekennen aan $target op $($script:mailbox)") {
                            Add-MailboxPermission -Identity $script:mailbox -User $target -AccessRights ReadPermission -InheritanceType All -ErrorAction Stop | Out-Null
                            Register-Action "Granted ReadPermission to $target on $($script:mailbox)"
                            $changed = $true
                        }
                    }
                }

                # ---------- Mailbox: verwijderen ----------
                "2" {
                    $entries = @(Get-EXOMailboxPermission -Identity $script:mailbox -ErrorAction Stop |
                        Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.User -ne "NT AUTHORITY\SELF" })

                    if ($entries.Count -eq 0) {
                        Write-Host "Geen verwijderbare (expliciete) mailbox-rechten gevonden." -ForegroundColor Yellow
                        Write-Log "No removable mailbox permissions found."
                        break
                    }

                    $items = $entries | ForEach-Object {
                        [pscustomobject]@{ Display = "$($_.User) [$($_.AccessRights -join ', ')]"; Data = $_ }
                    }
                    Write-Host "`nHuidige mailbox-rechten:" -ForegroundColor Cyan
                    $pick = Select-Entry $items "Welke wil je verwijderen? (nummer)"
                    $user   = $pick.Data.User
                    $rights = @($pick.Data.AccessRights)

                    if (Confirm-Action "rechten [$($rights -join ', ')] van $user verwijderen op $($script:mailbox)") {
                        foreach ($r in $rights) {
                            Remove-MailboxPermission -Identity $script:mailbox -User $user -AccessRights $r -InheritanceType All -Confirm:$false -ErrorAction Stop
                        }
                        Register-Action "Removed [$($rights -join ', ')] for $user on $($script:mailbox)"
                        $changed = $true
                    }
                }

                # ---------- Mailbox: reset FullAccess voor 1 persoon ----------
                "3" {
                    $target = Read-TargetUser "Gebruiker waarvan FullAccess opnieuw wordt ingesteld (user@domain.com)"

                    $existing = @(Get-EXOMailboxPermission -Identity $script:mailbox -User $target -ErrorAction SilentlyContinue |
                        Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited })

                    if ($existing.Count -gt 0) {
                        Write-Host "FullAccess is momenteel aanwezig voor $target." -ForegroundColor Cyan
                    }
                    else {
                        Write-Host "Let op: $target heeft nu geen expliciete FullAccess. Er wordt alleen opnieuw toegekend." -ForegroundColor Yellow
                    }

                    if (Confirm-Action "FullAccess van $target op $($script:mailbox) resetten (verwijderen en opnieuw toekennen met AutoMapping `$true)") {
                        if ($existing.Count -gt 0) {
                            Remove-MailboxPermission -Identity $script:mailbox -User $target -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                            Write-Log "Removed FullAccess for $target on $($script:mailbox) (reset step 1)"
                        }
                        Add-MailboxPermission -Identity $script:mailbox -User $target -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop | Out-Null
                        Register-Action "Reset FullAccess (AutoMapping: True) for $target on $($script:mailbox)"
                        $changed = $true
                    }
                }

                # ---------- Calendar: toevoegen / wijzigen ----------
                "4" {
                    $target = Read-TargetUser "Gebruiker die rechten krijgt (user@domain.com)"
                    $ca = Read-Option "Kies: 1 = Editor, 2 = Reviewer, 3 = Owner" @('1','2','3')
                    $accessRight = @{ '1' = 'Editor'; '2' = 'Reviewer'; '3' = 'Owner' }[$ca]

                    $existingPermission = Get-EXOMailboxFolderPermission -Identity $calendarIdentity -User $target -ErrorAction SilentlyContinue

                    if ($existingPermission) {
                        if (Confirm-Action "rechten van $target op $calendarIdentity wijzigen naar $accessRight (nu: $($existingPermission.AccessRights -join ', '))") {
                            Set-MailboxFolderPermission -Identity $calendarIdentity -User $target -AccessRights $accessRight -ErrorAction Stop | Out-Null
                            Register-Action "Updated rights to $accessRight for $target on $calendarIdentity"
                            $changed = $true
                        }
                    }
                    else {
                        if (Confirm-Action "$accessRight toekennen aan $target op $calendarIdentity") {
                            Add-MailboxFolderPermission -Identity $calendarIdentity -User $target -AccessRights $accessRight -ErrorAction Stop | Out-Null
                            Register-Action "Granted $accessRight rights to $target on $calendarIdentity"
                            $changed = $true
                        }
                    }
                }

                # ---------- Calendar: verwijderen ----------
                "5" {
                    $entries = @(Get-EXOMailboxFolderPermission -Identity $calendarIdentity -ErrorAction Stop |
                        Where-Object { $_.User.ToString() -notin @("Default","Anonymous") })

                    if ($entries.Count -eq 0) {
                        Write-Host "Geen expliciete agenda-rechten om te verwijderen." -ForegroundColor Yellow
                        Write-Log "No removable calendar permissions found."
                        break
                    }

                    $items = $entries | ForEach-Object {
                        [pscustomobject]@{ Display = "$($_.User) [$($_.AccessRights -join ', ')]"; Data = $_ }
                    }
                    Write-Host "`nHuidige agenda-rechten:" -ForegroundColor Cyan
                    $pick = Select-Entry $items "Welke wil je verwijderen? (nummer)"
                    $user = $pick.Data.User.ToString()

                    if (Confirm-Action "agenda-rechten van $user verwijderen op $calendarIdentity") {
                        Remove-MailboxFolderPermission -Identity $calendarIdentity -User $user -Confirm:$false -ErrorAction Stop
                        Register-Action "Removed calendar rights for $user on $calendarIdentity"
                        $changed = $true
                    }
                }

                # ---------- Contacts: toevoegen / wijzigen ----------
                "6" {
                    $target = Read-TargetUser "Gebruiker die rechten krijgt (user@domain.com)"
                    $co = Read-Option "Kies: 1 = Editor, 2 = Reviewer, 3 = Owner" @('1','2','3')
                    $accessRight = @{ '1' = 'Editor'; '2' = 'Reviewer'; '3' = 'Owner' }[$co]

                    $existingPermission = Get-EXOMailboxFolderPermission -Identity $contactsIdentity -User $target -ErrorAction SilentlyContinue

                    if ($existingPermission) {
                        if (Confirm-Action "rechten van $target op $contactsIdentity wijzigen naar $accessRight (nu: $($existingPermission.AccessRights -join ', '))") {
                            Set-MailboxFolderPermission -Identity $contactsIdentity -User $target -AccessRights $accessRight -ErrorAction Stop | Out-Null
                            Register-Action "Updated rights to $accessRight for $target on $contactsIdentity"
                            $changed = $true
                        }
                    }
                    else {
                        if (Confirm-Action "$accessRight toekennen aan $target op $contactsIdentity") {
                            Add-MailboxFolderPermission -Identity $contactsIdentity -User $target -AccessRights $accessRight -ErrorAction Stop | Out-Null
                            Register-Action "Granted $accessRight rights to $target on $contactsIdentity"
                            $changed = $true
                        }
                    }
                }

                # ---------- Contacts: verwijderen ----------
                "7" {
                    $entries = @(Get-EXOMailboxFolderPermission -Identity $contactsIdentity -ErrorAction Stop |
                        Where-Object { $_.User.ToString() -notin @("Default","Anonymous") })

                    if ($entries.Count -eq 0) {
                        Write-Host "Geen expliciete contacts-rechten om te verwijderen." -ForegroundColor Yellow
                        Write-Log "No removable contacts permissions found."
                        break
                    }

                    $items = $entries | ForEach-Object {
                        [pscustomobject]@{ Display = "$($_.User) [$($_.AccessRights -join ', ')]"; Data = $_ }
                    }
                    Write-Host "`nHuidige contacts-rechten:" -ForegroundColor Cyan
                    $pick = Select-Entry $items "Welke wil je verwijderen? (nummer)"
                    $user = $pick.Data.User.ToString()

                    if (Confirm-Action "contacts-rechten van $user verwijderen op $contactsIdentity") {
                        Remove-MailboxFolderPermission -Identity $contactsIdentity -User $user -Confirm:$false -ErrorAction Stop
                        Register-Action "Removed contacts rights for $user on $contactsIdentity"
                        $changed = $true
                    }
                }

                # ---------- Verversen ----------
                "8" {
                    Show-MailboxDetails
                    Show-Permissions
                }

                # ---------- Andere mailbox ----------
                "9" {
                    Export-CalendarShares
                    Select-Mailbox
                    Show-MailboxDetails
                    Show-Permissions
                }

                default {
                    Write-Host "Ongeldige keuze." -ForegroundColor Red
                    Write-Log "Invalid main menu selection: $sel"
                }
            }

            if ($changed) {
                Write-Host "`nOverzicht wordt ververst..." -ForegroundColor Cyan
                Show-Permissions
            }
        }
        catch {
            if ($_.Exception.Message -eq "USER_EXIT") {
                Write-Host "`nActie geannuleerd. Terug naar het menu." -ForegroundColor Yellow
                Write-Log "Action cancelled by user (menu selection $sel)."
            }
            else {
                Write-Warning "Actie mislukt: $($_.Exception.Message)"
                Write-Log "ERROR during menu action $($sel): $($_.Exception.Message)"
            }
        }
    }
}
catch {
    if ($_.Exception.Message -eq "USER_EXIT") {
        Write-Host "`nScript afgesloten door gebruiker." -ForegroundColor Yellow
        Write-Log "Script exited by user during startup."
    }
    else {
        Write-Host "Onverwachte fout: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "FATAL: $($_.Exception.Message)"
    }
}

# -------------------------------
# Export calendar explicit shares
# -------------------------------
Export-CalendarShares

# -------------------------------
# FINAL SUMMARY
# -------------------------------
Write-Host "`n=== FINAL SUMMARY ===" -ForegroundColor Cyan
if ($script:Actions.Count -gt 0) {
    foreach ($a in $script:Actions) { Write-Host " - $a" -ForegroundColor Yellow }
}
else {
    Write-Host $script:LastAction -ForegroundColor Yellow
}
Write-Host "`nLog opgeslagen in: $($script:LogFile)" -ForegroundColor Cyan
Write-Log "=== Script Completed. Actions performed: $($script:Actions.Count). Last action: $($script:LastAction) ==="