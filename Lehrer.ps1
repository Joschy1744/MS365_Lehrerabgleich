# Installiere die erforderlichen Module, wenn nicht bereits installiert
if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

if (-not (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module -Name AzureAD -Force -AllowClobber
}


if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module -Name Microsoft.Graph -Force -AllowClobber
}


# Attribute
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/configure-user-account-properties-with-microsoft-365-powershell?view=o365-worldwide
Import-Module -Name ImportExcel
Import-Module -Name Communary.PASM
Import-Module Microsoft.Graph.Users


##################################################################################################################################################################
## Löschen der Nutzer, die nicht mehr in eingelesener Datei sind ACHTUNG: immer erst Dry Run mit False machen und Exportdateien kontrollieren!
##
$deleteOldUsers = $false
##
##
## Anlegen der Nutzer, die in eingelesener Datei sind und noch nicht in AD ACHTUNG: immer erst Dry Run mit False machen und Exportdateien kontrollieren!
##
$generateNewUsers = $false
##
##
## erstelle Gruppen für Klassen
##
$gruppenerstellen = $true
##
##
## Für die Jahresgruppen Klassen: Zuerst Lehrer durchlaufen lassen, dort wird das Klassenteam erstellt
## Schuljahresstart
##
$jahr = "2024"
##
##
## Lehrer oder Schülerdomain
##
$TargetUsername = "@hm.de"
##
##
##
## Hinzufügen zu folgenden Gruppen
##
$GroupName1 = "Lehrer-Ersteller"
$GroupID1 = "123-fc67-4107-ab90-926d378f157d"
##
$GroupName2 = "#Lehrkräfte"
$GroupID2 = "123-fc69-4f55-b332-76fe1012b98b"
##
$GroupName3 = "#HMS"
$GroupID3 = "d123-1f85-4ecf-97a6-ea589da8e1f2"
##
##
## Array mit den Klassennamen, die übersprungen werden sollen
$klassenArray = @("E2*", "Q2*", "Q4*")
##
##
## Mail-Versand der Zugangsdaten an:
$MailIT = "it@hm.de"
##
## Mail des Abensender, sollte identisch mit dem Login sein
$MailAbsender = "office365@onmicrosoft.com"
##
##
$lizenzLuL = 'ENTERPRISEPACKPLUS_FACULTY'
##################################################################################################################################################################


# Verbindung mit Azure AD herstellen
Connect-AzureAD
Write-Host ("AzureAD verbunden") -ForegroundColor Green

# Verbindung mit Microsoft Teams herstellen
Connect-MicrosoftTeams
Write-Host ("MicrosoftTeam verbunden") -ForegroundColor Green

# Verbinden mit Microsoft Graph
Connect-MgGraph -Scopes  User.ReadWrite.All, Organization.Read.All, Directory.ReadWrite.All
Write-Host ("Graph verbunden") -ForegroundColor Green

# Pfad zur Excel-Datei
# Den Ordnerpfad, in dem das Skript ausgeführt wird, abrufen
$ordnerPfad = $PSScriptRoot
$logFilePath = "$PSScriptRoot\logfile.txt"

# Öffnen oder erstellen Sie die Logdatei und leiten Sie die Ausgabe dorthin um
Start-Transcript -Path $logFilePath

# Ausgabe des Ordnerpfads
Write-Host "Das Skript wird im Ordner ausgeführt: $ordnerPfad"


# Definieren Sie die Wildcard für den Dateinamen (z.B. #SPH*.txt)
$dateiWildcard = "#SPH-Lehrer*.xlsx"

# Suchen Sie nach Excel-Dateien im Ordner mit der Wildcard im Namen
$excelDatei = Get-ChildItem -Path $ordnerPfad -Filter $dateiWildcard | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Überprüfen, ob eine passende Datei gefunden wurde
if ($excelDatei -ne $null) {
    # Jetzt können Sie die gefundene Excel-Datei einlesen
    $excelData = [System.Collections.ArrayList](Import-Excel -Path $excelDatei.FullName)
}
else {
    Write-Host "Keine passende Datei gefunden."
    Exit
}


# Alle Benutzer abrufen
# ArrayList für performanteres entfernen
$users =  [System.Collections.ArrayList](Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -like "*$TargetUsername" } | Sort-Object -Property Surname)
 
$anzahlUserInAD = $users.Count
$anzahlUserInImport = $excelData.Count

# Array für fehlende Benutzer erstellen
$missingUsers = @()

$missingUsersinAD = @()

$usersAddToAD = @()

$usersDeletedFromAD = @()

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# check, ob eingelesene Nutzer in AD angelegt sind
$i = 1


foreach ($excelUser in $excelData) {

    Write-Host ("" + $i + " von " + $anzahlUserInImport + " Pruefen:  " + $excelUser.Nachname + ", " + $excelUser.Vorname + " Kürzel: " + $excelUser.Lehrer_Kuerzel )
              
    $userFound = $false
    foreach ($user in $users) {
        [string]$GivenName = $user.GivenName | Out-String
        [string]$Vorname = $excelUser.Vorname | Out-String

        if ( ($user.PhysicalDeliveryOfficeName -eq $excelUser.Lehrer_Kuerzel) -or ($GivenName -eq $Vorname -and $user.Surname -eq $excelUser.Nachname) -or ($excelUser.SAP_Personalnummer -and ($user.ExtensionProperty.employeeId -eq $excelUser.SAP_Personalnummer)) ) {
            $userFound = $true
            Write-Host ("" + $i + " von " + $anzahlUserInImport + " Gefunden: " + $excelUser.Nachname + ", " + $excelUser.Vorname + " Kürzel: " + $user.PhysicalDeliveryOfficeName)
            break
        }
    }

    # Wenn der Benutzer nicht in der AD gefunden wurde, zur Liste der fehlenden Benutzer hinzufügen
    if (!$userFound) {
        $missingUserInAD = [PSCustomObject]@{
            Lehrer_Kuerzel     = $excelUser.Lehrer_Kuerzel
            Nachname           = $excelUser.Nachname
            Vorname            = $excelUser.Vorname
            SAP_Personalnummer = $excelUser.SAP_Personalnummer
           
        }
        $missingUsersinAD += $missingUserInAD

        if ($generateNewUsers) {
            ## erstellen des neuen Nutzers, zu weisen der Lizenzen
            $vorname = $excelUser.Vorname
            $nachname = $excelUser.Nachname
           
            $dpname = $vorname.Substring(0, 1). + ". " + $nachname

            # Leerzeichen durch Bindestriche ersetzen
            $nachname = $nachname -replace " ", "-"

            # Umlaute im Nachnamen ersetzen
            $nachname = Replace-Umlauts -inputString $nachname.ToLower()
            $vorname = Replace-Umlauts -inputString $vorname.ToLower()
          
            ## Nutzername bzw Email zusammenabuen
            $mailNickname = $vorname.Substring(0, 1) + "." + $nachname
            $userPrincipalName = $mailNickname + $TargetUsername
               
           
            $lizenzfuerLuL = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $lizenzLuL
           
            $pw = "Hms" + ($excelUser.Geburtsdatum -replace ".", "") 
            $passwordProfile = @{
                forceChangePasswordNextSignIn = $true
                password                      = $pw
            }
          

            New-MgUser  -PasswordProfile $passwordProfile -GivenName $excelUser.Vorname -Surname $excelUser.Nachname -DisplayName $dpname -PhysicalDeliveryOfficeName $excelUser.Lehrer_Kuerzel  -JobTitle "Lehrer" -Country "Deutschland"  -UserPrincipalName $userPrincipalName -AccountEnabled -EmployeeId $excelUser.SAP_Personalnummer -MailNickName $mailNickname -UsageLocation "DE"
            Set-MgUserLicense -UserId $userPrincipalName -AddLicenses @{SkuId = $lizenzfuerLuL.SkuId } -RemoveLicenses @()
            
            $userAddToAD = [PSCustomObject]@{
                Nachname = $excelUser.Nachname
                Vorname  = $excelUser.Vorname
                Kuerzel  = $excelUser.Lehrer_Kuerzel
                Email    = $userPrincipalName
                Passwort = $passwort
            }
            $usersAddToAD += $userAddToAD

            $vorname = $excelUser.Vorname
            $nachname = $excelUser.Nachname
               
            $params = @{
                Message         = @{
                    Subject      = "Neuer Nutzer als Lehrer angelegt"
                    Body         = @{
                        ContentType = "HTML"  
                        Content     = "Hallo,<br>ein neuer Lehrer wurde soeben in MS365/Teams hinzugefügt.<br><br>-----------------------------<br><br><b>Name:</b> $vorname $nachname <br><b>E-Mail:</b> $userPrincipalName<br><b>Passwort:</b> $pw<br><br>-----------------------------<br><br>Bitte leite diese Daten an den/die neue Kolleg/in weiter.<br><br><i>Das ist eine automatisch generierte E-Mail. Bitte antworte nicht auf diese E-Mail.</i>"
                    }
                    ToRecipients = @(
                        @{
                            EmailAddress = @{
                                Address = $MailIT
                            }
                        }
                    )
                }
                SaveToSentItems = "false"
            }
            # A UPN can also be used as -UserId.
                    
            Send-MgUserMail -UserId  $MailAbsender -BodyParameter $params


        }
    }
    else {
        #Entfernen des Datensatzes aus $Users um Suchlaufzeit zu verkürzen
        # $users.Remove($user)
    }

    
    $i++
    Write-Host ("---------------------------------------------------------")
}


# Fehlende Benutzer aus LUSD in CSV-Datei schreiben
$missingUsersinADCsvPath = $ordnerPfad + "\nutzerInImportAberNichtInAD.csv"
$missingUsersinAD | Export-Csv -Path $missingUsersinADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

if ($generateNewUsers) {
    $usersAddToADCsvPath = $ordnerPfad + "\neuErstellteNutzer_$timestamp.csv"
    $usersAddToAD | Export-Csv -Path $usersAddToADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
}

#### Neu Nutzer laden, da potentiell neu angelegt.

$users = Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -like "*$TargetUsername" } | Sort-Object -Property Surname
$AllTeamsInOrg = Get-Team 


$i = 0
# Schleife durch alle Benutzer
foreach ($user in $users) {
    $i++

    $userFound = $false
     
    Write-Host ("" + $i + " von " + $anzahlUserInAD )
    Write-Host (" Pruefen:  " + $user.GivenName + ", " + $user.Surname  )

    if (($user.JobTitle -notlike "*Lehrer*") -and ($user.JobTitle -notlike "*Förderschullehrer*")) {
        Write-Host ("---------------------------------------------------------")
        continue
    }


    # Schleife durch die Excel-Daten
    foreach ($excelUser in $excelData) {
        [string]$GivenName = $user.GivenName | Out-String
        [string]$Vorname = $excelUser.Vorname | Out-String
         
         
        # Überprüfen, ob Vorname und Nachname übereinstimmen oder IDs gleich oder Kürzel.
       
       
        if ( ($user.PhysicalDeliveryOfficeName -eq $excelUser.Lehrer_Kuerzel) -or ($GivenName -eq $Vorname -and $user.Surname -eq $excelUser.Nachname) -or ($excelUser.SAP_Personalnummer -and ($user.ExtensionProperty.employeeId -eq $excelUser.SAP_Personalnummer)) ) {
           
            $userFound = $true
           
            if ($excelUser.Klassenlehrer_Klasse) {
                $klasse = $excelUser.Klassenlehrer_Klasse
            }
            else {
                $klasse = " "
            }
           
            if ($excelUser.Klassenlehrer_Vertreter_Klasse) {
                $stvklasse = $excelUser.Klassenlehrer_Vertreter_Klasse
            }
            else {
                $stvklasse = " "
            }
            if ($excelUser.SAP_Personalnummer) {
                $personalnummer = $excelUser.SAP_Personalnummer
            }
            else {
                $personalnummer = " "
            }


            
            $extensionProps = New-Object System.Collections.Generic.Dictionary"[String,String]"
            $extensionProps.Add("employeeId", $personalnummer)

            # Department auf den Wert der Klasse setzen
            Set-AzureADUser -ObjectId $user.ObjectId -Department $klasse -Country "Deutschland"  -State " " -PhysicalDeliveryOfficeName $excelUser.Lehrer_Kuerzel -ExtensionProperty $extensionProps -CompanyName " "
           
            Write-Host ("Eigenschaften für Benutzer " + $user.UserPrincipalName + " gesetzt.")

            # hinzufügen zu den Gruppen 1-3

            # Benutzer zum Teams-Team hinzufügen
            try {
                $membershipType = "Member"
                Add-TeamUser -GroupId $GroupID1 -User $user.UserPrincipalName -Role $membershipType
                Add-TeamUser -GroupId $GroupID2 -User $user.UserPrincipalName -Role $membershipType
                Add-TeamUser -GroupId $GroupID3 -User $user.UserPrincipalName -Role $membershipType

                Write-Host ("Benutzer " + $user.UserPrincipalName + " wurde als " + $membershipType + " zum Team Lehrer et. al. hinzugefügt.")
            }
            catch {
                Write-Host ("Fehler beim Hinzufügen von Benutzer " + $user.UserPrincipalName + " zum Team " + $_.Exception.Message)
            }

              # Überprüfen, ob der aktuelle Klassenname einem Muster im Array entspricht
                $erstellen = $true
                foreach ($pattern in $klassenArray) {
                    if ($klasse -like $pattern) {
                        Write-Host "$klasse überspringen..."
                        $erstellen = $false
                        break
                    }
                }
                # Wenn ein Muster getroffen wurde, bleibt $erstellen $true


            if ($gruppenerstellen -and $erstellen -and $excelUser.Klassenlehrer_Klasse) {
                    
                  
                # Überprüfen, ob ein Microsoft Teams-Team mit dem Namen der Klasse existiert

                $classTeamName = $excelUser.Klassenlehrer_Klasse + " Klassengruppe " + $jahr
                $classTeam = $null
                   
                foreach ($team in $AllTeamsInOrg) {
                    if ($team.DisplayName -eq $classTeamName) {
                        $classTeam = $team
                        break
                    }
                }
                                                      
                if (!$classTeam) {
               
                    # Team erstellen #-Visibility Private -Template "EDU_Class"
                    $classTeam = New-Team  -DisplayName $classTeamName  -MailNickName $classTeamName.replace(" ", "") -Template "EDU_Class"
                         
                    # Überprüfe, ob das Team erfolgreich erstellt wurde
                    if ($?) {
                        Write-Host ("Team " + $classTeam.DisplayName + " wurde erstellt.")
                        $AllTeamsInOrg += $classTeam
                    }
                    else {
                        Write-Output ("Fehler beim Erstellen des Klassenteams.")
                    }                              
                }


                # Benutzer zum Teams-Team hinzufügen
                try {
                    $membershipType = "Owner"
                    Add-TeamUser -GroupId $classTeam.GroupId -User $user.UserPrincipalName -Role $membershipType
                    try {
                        Remove-TeamUser -GroupId $classTeam.GroupId -User office365@hmsdtzb.onmicrosoft.com 
                    }
                    catch {

                    }
                    Write-Host ("Benutzer " + $user.UserPrincipalName + " wurde als " + $membershipType + " zum Team " + $classTeam.DisplayName + " hinzugefügt.")
                }
                catch {
                    Write-Host ("Fehler beim Hinzufügen von Benutzer " + $user.UserPrincipalName + " zum Team " + $classTeam.DisplayName + ": " + $_.Exception.Message)
                }

            }
            
            
        }
    }
    
    # Wenn der Benutzer nicht in der Excel-Datei gefunden wurde, zur Liste der fehlenden Benutzer hinzufügen
    if (!$userFound -and $user.JobTitle -eq "Lehrer") {
        $missingUser = [PSCustomObject]@{
            Name              = $user.DisplayName
            Nachname          = $user.Surname
            Vorname           = $user.GivenName
            ObjectId          = $user.ObjectId
            UserPrincipalName = $user.UserPrincipalName
        }
        $missingUsers += $missingUser


        
        #Wenn Löschen $true, dann löschen
        if ($deleteOldUsers -and $user.JobTitle -eq "Lehrer") {
            Remove-AzureADUser -ObjectId $user.ObjectId
            $usersDeletedFromAD += $missingUser
            Write-Host("Nutzer "+$user.UserPrincipalName+" wurde gelöscht.")
        }

    }
  Write-Host ("---------------------------------------------------------")
}

# Fehlende Benutzer in CSV-Datei schreiben
$missingUsersCsvPath = $ordnerPfad + "\nutzerInAdAberNichtInImport.csv"
$missingUsers | Export-Csv -Path $missingUsersCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

if ($deleteOldUsers) {
    $usersDeletedFromADCsvPath = $ordnerPfad + "\geloeschteNutzer_$timestamp.csv"
    $usersDeletedFromAD | Export-Csv -Path $usersDeletedFromADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
}
# Verbindung zu Azure AD trennen
Disconnect-AzureAD

# Beenden der Transkription und Schließen der Logdatei
Stop-Transcript

function Replace-Umlauts {
    param (
        [string]$inputString
    )

    $umlautMap = @{
        "ä" = "ae"
        "ö" = "oe"
        "ü" = "ue"
        "ß" = "ss"
       
    }

    foreach ($umlaut in $umlautMap.Keys) {
        $inputString = $inputString.Replace($umlaut, $umlautMap[$umlaut])
    }

    return $inputString
}
