Import-Module -Name ActiveDirectory


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Accept", "application/json, application/vnd.io.beekeeper.chats.chat+json;version=1")
$headers.Add("Authorization", "Token 73e61683-627e-409d-ab75-0e1d6ea12699")
$headers.Add("Content-Type", "application/json")
$headers.Add("charset", "utf-8")



$TimeStamp = Get-Date -Format "dd-MM-yyyy_hh-mm-ss"
Start-Transcript -Path "C:\Logs\V2_$TimeStamp.log"
$users = Import-Csv -Path "\\pkhsrv241\ADSync\personen.csv" -Encoding Default -Delimiter ';' -Header "Familienname","Vorname","Titel","Titel2","Eintritttsdatum","Austrittsdatum","PersNr","Abteilung","Taetigkeit","ORGTX","KSTNummer","Betrieb"


### OU Mapping für user Move
    $ouMap = @{
    'Administration Allgemein'      = 'Administration Allgemein'
    'Administration Pflege'         = 'Administration Pflege'
    'Anästhesie'                    = 'Anästhesie'
    'Ambulantes WM,Struma'          = 'Anästhesie'
    'Augenzentrum'                  = 'Augenzentrum'
    'Betriebsrat'                   = 'Administration Allgemein'
    'Kloster der Kreuzschwestern'   = 'Betagtenheim'
    'KOFÜ'                          = 'KOFÜ'
    'Küche'                         = 'Küche'
    'IT-Service'                    = 'IT'
    'Materialwirtschaft (Logistik,' = 'Einkauf und Lager'
    'OP'                            = 'OP'
    'Patientenadministration'       = 'Patientenadministration'
    'Personalwohnung'               = 'Haustechnik'
    'Physikalische Therapie'        = 'Therapie'
    'Praktikant (Ang.)'             = 'Praktikanten'
    'Praktikant (Arb.)'             = 'Praktikanten'
    'Reinigung'                     = 'Service'
    'Röntgen'                       = 'Röntgen_MRI'
    'Seniorenwohnheim'              = 'Betagtenheim'
    'Service (Hotelkomponente)'     = 'Service'
    'Station 2.Ost'                 = '2. West_Ost,OU=Stationen'
    'Station 2.West'                = '2. West_Ost,OU=Stationen'
    'Station 3.Ost'                 = '3. West_Ost,OU=Stationen'
    'Station 3.West'                = '3. West_Ost,OU=Stationen'
    'Station 4.Ost'                 = '4. West_Ost,OU=Stationen'
    'Station 4.West'                = '4. West_Ost,OU=Stationen'
    'Ärzteschaft Hochrum ohne Admin'= 'Ärzte'
    'Controlling'                   = 'Controlling'
    'Werkstätte/Haustechnik'        = 'Haustechnik'
}

    ##ADMapping für Gruppen Zuordnung (Muss ich mal fertig machen lol)
    $configADGruppen = @{
    "Station 2.Ost"          = "secUsersPflege"
    "Station 2.West"         = "secUsersPflege"
    "Station 3.Ost"          = "secUsersPflege"
    "Station 3.West"         = "secUsersPflege"
    "Station 4.Ost"          = "secUsersPflege"
    "Station 4.West"         = "secUsersPflege"
    "Anästhesie"         = "secUsersAnästesie-Ambulanz"
    "Physikalische Therapie"  = "secUsersTherapie"
    }




            # Gruppen-Chat-Mapping basierend auf dem Department-Attribut
        $ChatMap = @{
            #PflegeStationär
            'Station 4.West'                = @('b5e157c6-83a2-4546-beb2-80c4729df5b0')
            'Station 4.Ost'                 = @('e64dee70-5af4-40bd-9b80-527db1b71a87')
            'Station 3.West'                = @('5a5e4cc9-a74e-43fd-ae5e-ca467dce0489')
            'Station 3.Ost'                 = @('e4d8542c-58c9-4196-9b91-d1ccaae427be')
            'Station 2.West'                = @('ea90b4ad-2cdc-4800-a704-4ecb3bc5ec06')
            'Station 2.Ost'                 = @('ea90b4ad-2cdc-4800-a704-4ecb3bc5ec06')
            'Seniorenwohnheim'              = @('00fdf5a3-e18b-4a4a-955d-e5984b21e34a')

            #ML
            'Physikalische Therapie'        = @('1f5e743e-2110-48ac-a953-e7fd6ff2b2a6')
            'OP'                            = @('c301f261-f26d-4550-8f91-990d0b6e1bf2')
            'Ambulantes WM,Struma'          = @('414cb612-d21a-4aec-bfee-d9e142fdc570')
            'Anästhesie'                    = @('414cb612-d21a-4aec-bfee-d9e142fdc570')
           #'Augenzentrum'                  = @('1f5e743e-2110-48ac-a953-e7fd6ff2b2a6') Nur ein User?
            'Labor'                         = @('414cb612-d21a-4aec-bfee-d9e142fdc570','f0b10fd5-f8cc-4a43-9c75-a17e6a55ac3b')
            'Röntgen'                       = @('2917ad74-b531-46eb-93e9-d6c2971c1fbb')


            #DL
            'Küche'                        = @('af167853-ec87-4d94-9a27-557777eb0ef5')
            'Service (Hotelkomponente)'    = @('b1e3c4d7-6337-4ed9-bbf7-357fedda0d0c')
            'Reinigung'                    = @('b9a6cb10-cd61-4d27-a455-f0b94f6f7ab8')

            #AD
            'Ärzteschaft Hochrum ohne Admin'       = @('a24af50e-5884-4552-9579-51650366a6e2')

            #PM
            #'Patientenadministration'      = @('PM Patientenmanagement')
            'Aufnahme'                      = @('894346ee-dda6-412a-a729-e28986be11d2')
            'Codierung'                     = @('2b2125b1-20bd-4eb6-8512-03d15092b063')
            'Ärztesekretariat'              = @('0b26fbe8-0a6d-4a6c-8185-c2c18f4b7d59')
            'Rezeption'                     = @('2c6ddf45-586c-4b5f-a3e9-9aa074d386bb')

        }

            # Beekeeper Gruppen-Mapping basierend auf dem Department-Attribut
        $BeeGroupMap = @{
            #PflegeStationär
            'Station 4.West'        = @('PS Pflege Stationär', 'PS 4. West')
            'Station 4.Os'          = @('PS Pflege Stationär', 'PS 4. Ost')
            'Station 3.West'        = @('PS Pflege Stationär', 'PS 3. West')
            'Station 3.Ost'         = @('PS Pflege Stationär', 'PS 3. Ost')
            'Station 2.West'        = @('PS Pflege Stationär', 'PS 2. Bereich')
            'Station 2.Ost'         = @('PS Pflege Stationär', 'PS 2. Bereich')
            'Seniorenwohnheim'      = @('PS Pflege Stationär', 'PS Betreutes Wohnen')

            #ML
            'Physikalische Therapie'         = @('ML Medizinische Leitungsstellen', 'ML Therapie')
            'OP'                             = @('ML Medizinische Leitungsstellen', 'ML OP')
            'Ambulantes WM,Struma'           = @('ML Medizinische Leitungsstellen', 'ML Wundmanagement', 'ML Ambulanz/Anästhesie/Labor')
            'Augenzentrum'                   = @('ML Medizinische Leitungsstellen', 'ML ML Augenordination')
            'Labor'                          = @('ML Medizinische Leitungsstellen', 'ML Ambulanz/Anästhesie/Labor', 'AD Labor')
            'Anästhesie'                     = @('ML Medizinische Leitungsstellen', 'ML Ambulanz/Anästhesie/Labor')
            'Röntgen'                        = @('ML Medizinische Leitungsstellen', 'ML Radiologie')

            #DL
            'Küche'                        = @('DL Dienstleistungen', 'DL Küche')
            'Service (Hotelkomponente)'    = @('DL Dienstleistungen', 'DL Service')
            'Reinigung'                    = @('DL Dienstleistungen', 'DL Reinigung')

            #AD
            'Ärzteschaft Hochrum ohne Admin'       = @('AD Stationsärzte')

            #PM
            'Patientenadministration'       = @('PM Patientenmanagement')
            'Aufnahme'                      = @('PM Patientenmanagement','PM Aufname')
            'Codierung'                     = @('PM Patientenmanagement','PM Codierung')
            'Ärztesekretariat'              = @('PM Patientenmanagement','PM Ärztesekretariat')
            'Rezeption'                     = @('PM Patientenmanagement','PM Rezeption')

            #GF (für Mapping zu Inkonsistent)


            #FIBU / LV
            #'Patientenadministration'       = @('FIBU / LV')
            #'Patientenadministration'       = @('FIBU Finanzbuchhaltung')
            #'Patientenadministration'       = @('FIBU Lohnverrechnung')

            #PA
            'Patientenverrechnung'           = @('PA Patientenabrechnung')
            'Leitung Patientenverrechnung'   = @('PA Patientenabrechnung')

            #HR
            'Leitung Personalbüro / Assistenz PD'        = @('HR Personalabteilung')
            'Personalbüro'                               = @('HR Personalabteilung')

            #CO&IT
            'Leitung IT und Controlling'                          = @('CO&IT Controlling','CO & IT IT')
            'Assistenz Controlling'                               = @('CO&IT Controlling')
            'IT-Administrator'                                    = @('CO & IT IT')


        }



##Martin Witting = 10177
##Martin Birner = 11309 & 10769
##Chrisitne Jeggler = 10053
##Michaela Manzl = "10827
##Laura Unterberger B.A.
$PersIDSkipProcessing = @("10177", "11309", "10769", "10053", "10827", "11317", "10204")


######################################################



# Hauptfunktion für die AD-Synchronisierung
function PKH-AdSync {
    # Iteriere durch alle Benutzer
    foreach ($User in $users) {
        Start-Sleep -Milliseconds 250
        # Extrahiere Benutzerinformationen aus dem Eingabeobjekt
        $Familienname = $User.Familienname
        $Vorname = $User.Vorname
        $Titel = $User.Titel
        $Titel2 = $User.Titel2
        $Eintrittsdatum = $User.Eintritttsdatum
        $Austrittsdatum = $User.Austrittsdatum
        $PersNr = $User.PersNr.Substring(3)
        $Tätigkeit = $User.Taetigkeit ##Tätigkeit
        $Abteilung = $User.Abteilung ##KSTName
        try
        {
        $KSTNummer = $User.KSTNummer.Substring(5)
        }
        catch{
        #$User
        }
        #$Betrieb = $User.Betrieb
        $ORGTX = $User.ORGTX ##Abteilung

        
        #pause
        # Generiere weitere Benutzerinformationen
        $Vollname = "$Vorname $Familienname"
        $Beschreibung = "$Abteilung ($KSTNummer) -- $ORGTX ($Tätigkeit) -- ($PersNr)"
        $GenPW = (PKH-RandomPassword 8)
        $InitPassword = ConvertTo-SecureString $GenPW -AsPlainText -Force

        # Füge Titel zum Vollnamen hinzu, falls vorhanden
        if ($Titel) {
            $Vollname = "$Titel $Vollname"
 
        }
        if ($Titel2) {
            $Vollname = "$Vollname, $Titel2"

        }

        # Überspringe die Verarbeitung für bestimmte PersNr
        if ($PersIDSkipProcessing -notcontains $PersNr) {
            # Suche nach vorhandenen AD-Benutzern mit der entsprechenden PersNr
            if ($ADUSER = Get-ADUser -Filter "EmployeeID -like '$PersNr'" -Properties Department, Title, AccountExpirationDate, Enabled, employeeID, sAMAccountName, mail, employeeNumber, homepage) {
                Write-Host "Gefunden: $Vollname mit ID $PersNr" -ForegroundColor Green
                # Aktualisiere vorhandenen AD-Benutzer
                PKH-UpdateAD $Vollname $KSTNummer $Beschreibung $Vollname $ADUSER $Tätigkeit $Abteilung $PersNr $AustrittsDatum $Eintrittsdatum $Vorname $Familienname $GenPW $Titel $Titel2 $ORGTX

                #PKH-UPDATEBEEKEEPER FUNKTION 
            } else {
                Write-Host "Nicht gefunden: $Vollname mit ID $PersNr" -ForegroundColor Red
                # Erstelle neuen AD-Benutzer
                PKH-CreateAD $Vollname $Beschreibung $Vorname $Familienname $OUPath $InitPassword $KSTNummer $Abteilung $Tätigkeit $Vollname $PersNr $AustrittsDatum $Eintrittsdatum $GenPW $Titel $Titel2 $ORGTX

                #PKH-CREATEBEEKEEPER FUNKTION 
            }
        } else {
            Write-Host "++++ $PersNr wird nicht verarbeitet" -ForegroundColor Red
        }
    }
    
}

# Hauptfunktion zum Aktualisieren von AD-User-Informationen
function PKH-UpdateAD($VollName, $KostenStelle, $Beschreibung, $Displayname, $ADUSER, $MitarbKreisbez, $Abteilung, $PersNr, $Austritt, $Eintritt, $Vorname, $Nachname, $GenPW, $Titel1, $Titel2, $ORGTX) {     
    # Sammle weitere Informationen
    $Login = $ADUSER.SamAccountname
    $BeekeeperID = PKH-GetBeekeeperID($Login)

    # Aktualisiere Attribute, wenn sie nicht vorhanden sind
    Update-ADAttributeIfNotSet $ADUSER Department $Abteilung
    Update-ADAttributeIfNotSet $ADUSER EmployeeID $PersNr
    Update-ADAttributeIfNotSet $ADUSER HomePage $KostenStelle
    Update-ADAttributeIfNotSet $ADUSER Description $Beschreibung


    # Verarbeite Austrittsdatum
    Process-Austrittsdatum $ADUSER $Austritt $Eintritt $VollName $MitarbKreisbez $Abteilung $KostenStelle $GenPW ($ADUSER.SamAccountname)

    # Aktualisiere BeekeeperID, wenn sie nicht im AD vorhanden ist
    Update-Beekeeper $ADUSER $BeekeeperID

}

# Funktion zum Erstellen eines neuen Active Directory-Benutzers
function PKH-CreateAD($VollName, $Beschreibung, $Vorname, $Nachname, $OUPath, $InitPassword, $KostenStelle, $Abteilung, $MitarbKreisbez, $Personalname, $PersNr, $Austritt, $Eintritt, $GenPW, $ORGTX) {
    

    # Erstelle den Login-Namen aus den Initialen des Vornamens und des Nachnamens
    $Login = $Vorname[0] + "." + $Nachname
    $BeeLogin = Convert-Umlauts ($Vorname[0] + "." + $Nachname) 

    # Überprüfe, ob der Login-Name bereits existiert, und erstelle ggf. einen neuen Login-Namen
    if(Get-ADUser -Filter {SamAccountName -eq $Login}) {
        $Login = $Vorname[0] + $Vorname[1] + "." + $Nachname
        Write-Host "+++++ Login already exists nutze $Login" -ForegroundColor Red
    }
    
    # Wenn das Austrittsdatum erreicht ist, erstelle den Benutzer nicht
    if($Austritt -and (Get-Date) -ge [Datetime]::ParseExact($Austritt, 'dd.MM.yyyy', $null)) {
        Write-Host "+++++ $VollName Ausgetreten und wird nicht erstellt +++++" -ForegroundColor Red
        return
    }
    
    # Erstelle den neuen Benutzer
    Write-Host "+++++ $VollName wird angelegt +++++" -ForegroundColor Green

    # Parameter für New-ADUser
    $params = @{
        Name = $VollName
        GivenName = $Vorname
        Surname = $Nachname
        SamAccountName = $Login
        UserPrincipalName = ($Login + "@pk-hochrum.com")
        EmployeeID = $PersNr
        Path = "OU=DomBenutzer,OU=User,OU=HRUM,OU=SANATORIUM,DC=sanatorium,DC=int"
        AccountPassword = $InitPassword
        Description = $Beschreibung
        HomePage = $KostenStelle
        Department = $Abteilung
        Title = $MitarbKreisbez
        EmailAddress = ($Login + "@pk-hochrum.com")
        DisplayName = $VollName
        ChangePasswordAtLogo = $true
        Enabled = $true
    }
    
    # Erstelle den neuen AD-Benutzer
    New-ADUser @params

    # Warte für 5 Sekunden
    Start-Sleep -Seconds 2

    # Hole AD-Benutzerinformationen basierend auf dem Login
    $ADUSER = Get-ADUser -Identity $Login -Properties *

     # Erstelle Beekeeper-Profil für den Benutzer
     PKH-CreateBeekeeperProfile $Vorname $Nachname $Login $KostenStelle $PersNr $User.ORGTX $MitarbKreisbez $GenPW ($Login + "@pk-hochrum.com") $Titel $Titel2

     # Hole Beekeeper-ID für den Benutzer
     $BeekeeperID = PKH-GetBeekeeperID($BeeLogin)

     #Füge nutzer zu Gruppenchat und Bee Gruppen hinzu 
     PKH-AddMemberToGroupChat $BeekeeperID $ADUser.Department
     PKH-UpdateBeekeeperGroup $BeekeeperID $ADUser.Department

      # Prüfe, ob die Beekeeper-ID gefunden wurde
      if ($BeekeeperID -ne "Not Found") {
           # Setze die Mitarbeiternummer für den AD-Benutzer auf die Beekeeper-ID
          Set-ADUser -Identity $ADUSER -employeeNumber $BeekeeperID

          # Gebe die aktualisierte Mitarbeiternummer aus
          Write-Host "+++++ Update employeeNumber zu $BeekeeperID"
     }
  

    # Sende Benachrichtigungen und füge den Benutzer zu Gruppen hinzu
<#
    $Nachricht = @" 
    Hallo, $USERNAME <br /> <br /> <b> $Vollname </b> wurde erstellt <br /> Login: <b> $Login </b> <br /> Password: <b> $GenPW </b> <br /> E-Mail: <b> $Login@pk-hochrum.com </b> <br /> <br /> Eintritt: <b> $Eintritt </b> <br /> Austritt: <b> $Austritt </b> <br /> PersonalNR: <b> $PersNr </b> <br /> Abteilung: <b> $Abteilung </b> <br /> KST: <b> $KostenStelle </b> <br />  <br /> LG die IT
"@

    Send-BeekeeperChatMessage -UserId "d375258f-6e49-4591-99a6-49e35a3d7dc9" -MessageBody $Nachricht

    #>

    Start-Sleep -Milliseconds 1000

    $File = PKH-LoginPDF $BeekeeperID $PersNr
    PKH-SendMail $VollName $Eintritt $Austritt $PersNr $MitarbKreisbez $ORGTX $KostenStelle "erstellt" $GenPW $Login $File
    #PKH-addADGroup $ADUSER $Abteilung $MitarbKreisbez
    #PKH-MoveAD $ADUSER $Abteilung $MitarbKreisbez
}

# Funktion zum Hinzufügen eines AD-Benutzers zu einer Gruppe basierend auf der Abteilung
function PKH-addADGroup($ADUSER, $Abteilung, $MitarbKreisbez) {
    # Iteriere durch die Gruppen im Mapping
    foreach ($group in $configADGruppen.GetEnumerator()) {
        # Wenn die Abteilung des Benutzers der Gruppe im Mapping entspricht
        if ($Abteilung -eq $group.Name) {
            # Füge den Benutzer zur entsprechenden Gruppe hinzu
            Add-ADGroupMember -Identity $group.Value -Members $ADUSER.SamAccountName
            Write-Host "+++++ Update Gruppe $($group.Value)"
            break
        }
    }
}

# Funktion zum Verschieben eines AD-Benutzers in die entsprechende OU basierend auf der Abteilung oder dem MitarbKreisbez
function PKH-MoveAD($ADUSER, $Abteilung, $MitarbKreisbez) {
    $ou = $ouMap[$Abteilung]
    if ($ou) {
        $ouPath = "OU=$ou,OU=Abteilungen,OU=User,OU=HRUM,OU=SANATORIUM,DC=sanatorium,DC=int"
        try {
            Move-ADObject -Identity $ADUSER -TargetPath $ouPath
            Write-Host "+++++ Abteilung: Verschiebe user $ADUSER nach OU: $ouPath" -ForegroundColor Green
        } catch {
            # Fehler beim Verschieben kann hier optional behandelt werden
        }
    } else {
        $ou = $ouMap[$MitarbKreisbez]
        if ($ou) {
            $ouPath = "OU=$ou,OU=Abteilungen,OU=User,OU=HRUM,OU=SANATORIUM,DC=sanatorium,DC=int"
            try {
                Move-ADObject -Identity $ADUSER -TargetPath $ouPath
                Write-Host "+++++ MitarbKreisbez: Verschiebe user $ADUSER nach OU: $ouPath" -ForegroundColor Green
        } catch {
            # Fehler beim Verschieben kann hier optional behandelt werden
        }
    } else {
        Write-Host "OU für '$Abteilung' oder '$MitarbKreisbez' nicht erkannt" -ForegroundColor Red
    }
  }
}

# Funktion zum Aktualisieren eines AD-Attributs, wenn es noch nicht gesetzt ist
function Update-ADAttributeIfNotSet($ADUser, $AttributeName, $AttributeValue) {
    if (!$ADUser.$AttributeName) {
        Set-ADUser -Identity $ADUSER -Replace @{$AttributeName = $AttributeValue}
        Write-Host "+++++ Update $AttributeName zu " $AttributeValue
    }
}

# Funktion zum Aktualisieren der Beekeeper-ID
function Update-Beekeeper($ADUser, $BeekeeperID) {
    if (!$ADUser.employeeNumber) {
        if ($BeekeeperID -ne "Not Found") {
            Set-ADUser -Identity $ADUSER -employeeNumber $BeekeeperID
            Write-Host "+++++ Update employeeNumber zu" $BeekeeperID
        }
        else
        {
        
            if($ADUser.Enabled)
            {

                  PKH-CreateBeekeeperProfile ($ADUSER.givenname) ($ADUser.surname) ($ADUSER.SamAccountname) ($ADUser.homepage) ($ADUser.employeeID) $ORGTX <#($ADUser.Department)#> ($ADUser.title) ("Hochrum_"+$ADUser.employeeID+"!") ($Login + "@pk-hochrum.com") $Titel $Titel2
                  Write-Host "+++++ Erstelle Beekeeper User"  (Convert-Umlauts $ADUSER.SamAccountname)

                  $BeekeeperID = PKH-GetBeekeeperID($Login)
                  Set-ADUser -Identity $ADUSER -employeeNumber $BeekeeperID
                  Write-Host "+++++ Update employeeNumber zu" $BeekeeperID

                  #pause

            }

        }
    } else {

        PKH-UpdateBeekeperUser ($ADUser.employeeNumber) ($ADUSER.SamAccountname) ($ADUser.givenname) ($ADUser.surname) ($ADUser.title) ($ADUser.Department) ($ADUser.mail) ($ADUser.homepage) $Titel $Titel2 $Eintritt $Austritt
        
        #PKH-AddMemberToGroupChat $BeekeeperID $ADUser.Department
        #PKH-UpdateBeekeeperGroup $BeekeeperID $ADUser.Department
        Write-Host "+++++ Update Beekeeper Infos"
    }
}


function Process-Austrittsdatum($ADUser, $Austritt, $Eintritt, $VollName, $MitarbKreisbez, $Abteilung, $KostenStelle, $GenPW, $Login) {
    
    # Prüfe ob ein Austrittsdatum gesetzt ist
    if ($Austritt) {
        
        # Hole das aktuelle Ablaufdatum des Benutzers aus dem Active Directory
        $ADUserExpiration = Get-ADUser -Identity $ADUser -Properties AccountExpirationDate | Select-Object -ExpandProperty AccountExpirationDate

        # Prüfe ob das aktuelle Ablaufdatum dem Austrittsdatum entspricht
        if ($ADUserExpiration -eq [Datetime]::ParseExact($Austritt, 'dd.MM.yyyy', $null)) {
            Write-Host "Das Austrittsdatum des Benutzers ist unverändert. Weitere Verarbeitung wird übersprungen." -ForegroundColor Yellow
            
            # Prüfe ob das Austrittsdatum erreicht ist und der Benutzer noch aktiv ist, deaktiviere den Benutzer und versende eine E-Mail
            $today = Get-Date
            if ($today -ge [Datetime]$ADUserExpiration) {
                if ($ADUser.Enabled) {
                    DisableAndMoveADUser $ADUser
                    PKH-SendMail $VollName $Eintritt $Austritt $PersNr $MitarbKreisbez $Abteilung $KostenStelle "deaktiviert" "" $Login ""
                    return
                }
            }
            return
        }

        # Setze oder Aktualisiere das Ablaufdatum des Kontos
        Write-Host "+++++ Setze Austrittsdatum auf" $Austritt -ForegroundColor Red
        Set-ADAccountExpiration -Identity $ADUser -DateTime ([Datetime]::ParseExact($Austritt, 'dd.MM.yyyy', $null))

        # Vergleiche das Austrittsdatum mit dem aktuellen Datum
        $AustrittDate = [Datetime]::ParseExact($Austritt, 'dd.MM.yyyy', $null)
        $today = Get-Date

        # Wenn das Austrittsdatum erreicht ist und der Benutzer noch aktiv ist, deaktiviere den Benutzer und versende eine E-Mail
        if ($today -ge [Datetime]$AustrittDate) {
            if ($ADUser.Enabled) {
                DisableAndMoveADUser $ADUser
                PKH-SendMail $VollName $Eintritt $Austritt $PersNr $MitarbKreisbez $Abteilung $KostenStelle "deaktiviert" "" $Login ""
                return
            }
        } else { 
            # Wenn das Austrittsdatum noch nicht erreicht ist und der Benutzer deaktiviert ist, aktiviere den Benutzer
            EnableADUserIfDisabled $ADUser 
        }

    } else {
        # Wenn kein Austrittsdatum gesetzt ist, entferne das Ablaufdatum und aktiviere den Benutzer
        if ($ADUser.AccountExpirationDate) {
            Write-Host "+++++ Ablaufdatum wird entfernt..." -ForegroundColor Green
            Clear-ADAccountExpiration -Identity $ADUser
        }
        
        # Wenn der Benutzer deaktiviert ist, aktiviere den Benutzer und versende eine Benachrichtigung
        if (!$ADUser.Enabled) {
            EnableADUserAndSendNotification $ADUser $VollName $Eintritt $MitarbKreisbez $Abteilung $KostenStelle $GenPW $Login
        }
    }
}

function DisableAndMoveADUser($ADUser) {
    Write-Host "+++++ User wird Deaktiviert" -ForegroundColor Yellow
    Disable-ADAccount -Identity $ADUser

    PKH-RemoveMemberFromGroupChat ($ADUser.employeeNumber) $Abteilung

    PKH-UpdateBeekeperUserState ($ADUser.employeeNumber) $false
    Write-Host "+++++ User wird Verschoben" -ForegroundColor Yellow
    Move-ADObject -Identity $ADUser -TargetPath "OU=Deaktivierte_User,OU=User,OU=HRUM,OU=SANATORIUM,DC=sanatorium,DC=int"



}

# Funktion zum Aktivieren eines AD-Benutzers, falls deaktiviert
function EnableADUserIfDisabled($ADUser) {
    if (!$ADUser.Enabled) {
        EnableADUserAndSendNotification $ADUser $VollName $Eintritt $MitarbKreisbez $Abteilung $KostenStelle $GenPW $Login
        PKH-AddMemberToGroupChat ($ADUser.employeeNumber) $Abteilung
    }
}

# Funktion zum Reaktivieren eines AD-Benutzers und Senden einer Benachrichtigung
function EnableADUserAndSendNotification($ADUser, $VollName, $Eintritt, $MitarbKreisbez, $Abteilung, $KostenStelle, $GenPW, $Login) {
    Write-Host "+++++ $VollName wird reaktiviert" -ForegroundColor Green
    Enable-ADAccount -Identity $ADUser
    PKH-UpdateBeekeperUserState ($ADUser.employeeNumber) $true
    Set-ADAccountPassword -Identity $ADUser -NewPassword (ConvertTo-SecureString -AsPlainText $GenPW -Force) -Reset
    Set-ADUser -Identity $ADUSER -ChangePasswordAtLogon $true

    $File = PKH-LoginPDF $BeekeeperID $PersNr
    PKH-SendMail $VollName $Eintritt $AustrittString $PersNr $MitarbKreisbez $Abteilung $KostenStelle "reaktiviert" $GenPW $Login $File
}

function Convert-Umlauts {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    $replaceMap = @{
        'ä' = 'ae';
        'ö' = 'oe';
        'ü' = 'ue';
        'ß' = 'ss';
        'à' = 'a';
        'á' = 'a';
        'â' = 'a';
        'ã' = 'a';
        'å' = 'a';
        'æ' = 'ae';
        'ç' = 'c';
        'è' = 'e';
        'é' = 'e';
        'ê' = 'e';
        'ë' = 'e';
        'ì' = 'i';
        'í' = 'i';
        'î' = 'i';
        'ï' = 'i';
        'ð' = 'd';
        'ñ' = 'n';
        'ò' = 'o';
        'ó' = 'o';
        'ô' = 'o';
        'õ' = 'o';
        'ø' = 'o';
        'ù' = 'u';
        'ú' = 'u';
        'û' = 'u';
        'ý' = 'y';
        'þ' = 'th';
        'ÿ' = 'y';
        '.' = '_';
        '-' = '_';
    }

    foreach ($key in $replaceMap.Keys) {
        $Text = $Text.Replace($key, $replaceMap[$key])
    }

    return $Text
}

function PKH-SendMail($VollnameMitTitel,$Eintrittsdatum,$Austrittsdatum,$PersNr,$Tätigkeit,$Abteilung,$KSTNummer,$aktion,$InitPassword,$Login,$File)
{


  $mailbody = Get-Content -Path "\\PKHSRV241\ADSync\mailbody.html" |Out-String
  $body = $mailbody.replace('@@VollnamemitTitel',$VollnameMitTitel).replace('@@Eintrittsdatum',$Eintrittsdatum).replace('@@Austrittsdatum',$Austrittsdatum).replace('@@PersNr',$PersNr).replace('@@Tätigkeit',$Tätigkeit).replace('@@Abteilung',$Abteilung).replace('@@KSTNummer',$KSTNummer).replace('@@aktion',$aktion).replace('@@InitPassword',$InitPassword).replace('@@Login',$Login).Replace('@@MailAdress',$Login + "@pk-hochrum.com")

    $MailMessage = @{
        To = "b.mayerl@pk-hochrum.com"
        Bcc = "admins@pk-hochrum.com","m.krause@pk-hochrum.com"
        From = "support@pk-hochrum.com"
        Subject =  "+++ " + $VollnameMitTitel + " (" + $PersNr + ")" +" +++"
            Body = $body
            Smtpserver = "exch-int.privatklinik-hochrum.com"
            BodyAsHtml = $true
            Encoding = “UTF8”
            Attachment =  "$File"
}

  Send-MailMessage @MailMessage

}

function PKH-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [ValidateRange(4,[int]::MaxValue)]
        [int] $length,
        [int] $upper = 1,
        [int] $lower = 1,
        [int] $numeric = 1,
        [int] $special = 1
    )
    if($upper + $lower + $numeric + $special -gt $length) {
        throw "number of upper/lower/numeric/special char must be lower or equal to length"
    }
    $uCharSet = "ABCDEFGHJKMNOPQRSTUVWXYZ"
    $lCharSet = "abcdefghjkmnopqrstuvwxyz"
    $nCharSet = "0123456789"
    $sCharSet = "!?"
    $charSet = ""
    if($upper -gt 0) { $charSet += $uCharSet }
    if($lower -gt 0) { $charSet += $lCharSet }
    if($numeric -gt 0) { $charSet += $nCharSet }
    if($special -gt 0) { $charSet += $sCharSet }
    
    $charSet = $charSet.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
    $rng.GetBytes($bytes)
 
    $result = New-Object char[]($length)
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
    }
    $password = (-join $result)
    $valid = $true
    if($upper   -gt ($password.ToCharArray() | Where-Object {$_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
    if($lower   -gt ($password.ToCharArray() | Where-Object {$_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
    if($numeric -gt ($password.ToCharArray() | Where-Object {$_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
    if($special -gt ($password.ToCharArray() | Where-Object {$_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }
 
    if(!$valid) {
         $password = PKH-RandomPassword $length $upper $lower $numeric $special
    }
    return $password
}

function PKH-GetBeekeeperID($Username) {

    $Username = Convert-Umlauts $Username

    try {
        $URL = "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/by_name/" + $Username
        $Get_Bee_ID = Invoke-RestMethod $URL -Method 'GET' -Headers $headers
        $Bee_ID = $Get_Bee_ID.id
        Write-Host "$Bee_ID gefunden" -ForegroundColor Yellow
        return $Bee_ID
    } catch { 

        $Exception = PKH-ParseErrorForResponseBody($_)
        $Exception = $Exception.error

        if($Exception -eq "Not Found")
        { 
            Write-Host "PKH-GetBeekeeperID: $Username not found" -ForegroundColor Magenta
            return $Exception
        }
        else
        {
            return $Exception
        }
    }
}

function PKH-GetBeekeeperUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Username
    )

    # Bereinige den Benutzernamen
    #$Username = $Username -replace '\.|-', '_' -replace '[äöü]', '${_}' -replace 'ß', 'ss' -replace 'ó', 'ss'
    $Username = Convert-Umlauts $Username

    try {
        # Senden der Anfrage, um den Beekeeper-Benutzer abzurufen
        $USER = Invoke-RestMethod -Uri "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/by_name/$Username" `
            -Method 'GET' -Headers $headers

        # Konvertiere das Ergebnis in JSON und anschließend zurück in ein PowerShell-Objekt
        $USER = $USER | ConvertTo-Json
        $USER = ConvertFrom-Json $USER

        return $USER
    } catch {
        # Behandle Fehler und gebe eine Fehlermeldung aus
        $Exception = PKH-ParseErrorForResponseBody $_
        $Exception = $Exception.error
        Write-Host "PKH-GetBeekeeperUser: $Exception" -ForegroundColor Red
        #pause
        return $Exception
    }
}

function PKH-UpdateBeekeperUser {
    param(
        $Bee_ID, $Username, $vorname, $nachname, $position,
        $abteilung, $mail, $kostenstelle, $Titel1, $Titel2, $Eintritt, $Austritt
    )

    $Beeusername = Convert-Umlauts $Username

    $customFields = @(
        @{key = 'firstname'; value = "$Titel1 $vorname"}
        @{key = 'lastname'; value = "$nachname"}
        @{key = 'position'; value = $position}
        @{key = 'abteilung'; value = $abteilung}
        @{key = 'kostenstelle'; value = $kostenstelle}
        @{key = 'eintrittsdatum'; value = $Eintritt}
        @{key = 'austrittsdatum'; value = $Austritt}
    )

    $body = @{
        email         = $mail
        name          = $Beeusername
        tenantuserid  = "SANATORIUM\$Username"
        custom_fields = $customFields
    } | ConvertTo-Json -Depth 3

    $body

    try {
        $Update = Invoke-RestMethod -Uri "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/$Bee_ID" `
            -Method 'PUT' -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))
        $Update = $Update | ConvertTo-Json
        $Update = ConvertFrom-Json $Update
    } catch {
        $Exception = PKH-ParseErrorForResponseBody($_)
        $Exception = $Exception.error
        Write-Host "PKH-UpdateBeekeperUser ($Beeusername) : $Exception" -ForegroundColor Red
        return $Exception
    }
}

function PKH-UpdateBeekeperAvatar($Bee_ID, $Beeusername) {
    $url = "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/$Bee_ID"
    $body = @{
        avatar = "https://robohash.org/$Beeusername.png"
    } | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $url -Method 'PUT' -Headers $headers -Body $body
    } catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Error updating avatar for Beekeeper user with ID $Bee_ID : $ErrorMessage" -ForegroundColor Red
    }
}

function PKH-UpdateBeekeperUserState {
    param(
        $Bee_ID,
        [bool]$State
    )

    # Erstelle das Body-Objekt basierend auf dem Statuswert
    $body = @{
        suspended = -not $State
    } | ConvertTo-Json

    try {
        # Führe die Anfrage aus, um den Benutzerstatus in Beekeeper zu aktualisieren
        $Update = Invoke-RestMethod -Uri "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/$Bee_ID" `
            -Method 'PUT' -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))

        # Konvertiere die Antwort in JSON und dann in ein PowerShell-Objekt
        $Update = $Update | ConvertTo-Json
        $Update = ConvertFrom-Json $Update

        # Gebe das aktualisierte Benutzerobjekt zurück
        return $Update
    } catch {
        # Behandle Fehler und gebe eine Fehlermeldung aus
        $Exception = PKH-ParseErrorForResponseBody($_)
        $Exception = $Exception.error
        Write-Host "PKH-UpdateBeekeperUserState: $Exception" -ForegroundColor Red
        return $Exception
    }
}

function PKH-UpdateBeekeeperGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $BeekeeperUserID,
        [Parameter(Mandatory = $true)] [string] $Department
    )

        # Finde die entsprechenden Gruppen basierend auf dem Department
        $BeekeeperGroups = $BeeGroupMap[$Department]

        if (-not $BeekeeperGroups) {
            Write-Host "No matching groups found for department: $Department" -ForegroundColor Yellow
            return
        }

        foreach ($BeekeeperGroup in $BeekeeperGroups) {
            try {
            # Erstelle den Anfragekörper für die API
            $body = @{
                group = @{
                    name = $BeekeeperGroup
                }
            } | ConvertTo-Json

            # Führe die API-Anfrage aus, um den Benutzer zur Gruppe hinzuzufügen
            $AddUserToGroupUrl = "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/$BeekeeperUserID/group_memberships"
            $Update = Invoke-RestMethod -Uri $AddUserToGroupUrl -Method 'POST' -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))


            # Gebe eine Erfolgsmeldung aus
            Write-Host "Updated group for Beekeeper user with ID $BeekeeperUserID to $BeekeeperGroup" -ForegroundColor Green

                } catch {
             # Behandle Fehler und gebe eine Fehlermeldung aus
             $Exception = PKH-ParseErrorForResponseBody $_
             $Exception = $Exception.error
             Write-Host "PKH-UpdateBeekeeperGroup: $Exception" -ForegroundColor Red

            }

        }
}

function PKH-AddMemberToGroupChat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $BeekeeperUserId,
        [Parameter(Mandatory = $true)] [string] $Department,
        [Parameter(Mandatory = $false)] [string] $Role = "MEMBER"
    )

        # Finde die entsprechenden Gruppen-Chat-IDs basierend auf dem Department
        $ChatIds = $ChatMap[$Department]

        if (-not $ChatIds) {
            Write-Host "No matching group chats found for department: $Department" -ForegroundColor Yellow
            return
        }

        #$headers.Add("Accept", "application/json, application/vnd.io.beekeeper.chats.chat+json;version=1")

        foreach ($ChatId in $ChatIds) {

            try {

            # Erstelle den Anfragekörper für die API
            $body = @{
                user_id = $BeekeeperUserId
                role    = $Role
            } | ConvertTo-Json


            # Führe die API-Anfrage aus, um den Benutzer zum Gruppen-Chat hinzuzufügen
            $AddMemberToChatUrl = "https://privatklinik-hochrum.de.beekeeper.io/api/2/chats/groups/$ChatId/members"
            $Update = Invoke-RestMethod -Uri $AddMemberToChatUrl -Method 'POST' -Headers $headers -Body $body
            #$Update | ConvertTo-Json

            # Gebe eine Erfolgsmeldung aus
            Write-Host "Added member with ID $BeekeeperUserId to group chat $ChatId for department $Department" -ForegroundColor Green

        } catch {
            # Behandle Fehler und gebe eine Fehlermeldung aus
            $Exception = PKH-ParseErrorForResponseBody $_
            $Exception = $Exception.error
            Write-Host "PKH-AddMemberToGroupChat: $Exception" -ForegroundColor Red
          
            #return $Exception
    }

  }

}

function PKH-RemoveMemberFromGroupChat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $BeekeeperUserId,
        [Parameter(Mandatory = $true)] [string] $Department,
        [Parameter(Mandatory = $false)] [string] $Role = "MEMBER"
    )

        # Finde die entsprechenden Gruppen-Chat-IDs basierend auf dem Department
        $ChatIds = $ChatMap[$Department]

        if (-not $ChatIds) {
            Write-Host "No matching group chats found for department: $Department" -ForegroundColor Yellow
            return
        }

        foreach ($ChatId in $ChatIds) {

            try {


            # Führe die API-Anfrage aus, um den Benutzer zum Gruppen-Chat hinzuzufügen
            $AddMemberToChatUrl = "https://privatklinik-hochrum.de.beekeeper.io/api/2/chats/groups/$ChatId/members/$BeekeeperUserId"
            $Update = Invoke-RestMethod -Uri $AddMemberToChatUrl -Method 'DELETE' -Headers $headers
            $Update | ConvertTo-Json

            # Gebe eine Erfolgsmeldung aus
            Write-Host "Added member with ID $BeekeeperUserId to group chat $ChatId for department $Department" -ForegroundColor Green

        } catch {
            # Behandle Fehler und gebe eine Fehlermeldung aus
            $Exception = PKH-ParseErrorForResponseBody $_
            $Exception = $Exception.error
            Write-Host "PKH-AddMemberToGroupChat: $Exception" -ForegroundColor Red
          
    }

  }

}

function PKH-CreateBeekeeperProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Vorname,
        [Parameter(Mandatory = $true)] [string] $Familienname,
        [Parameter(Mandatory = $true)] [string] $ADusername,
        [Parameter(Mandatory = $true)] [string] $KSTNummer,
        [Parameter(Mandatory = $true)] [string] $PersNr,
        [Parameter(Mandatory = $true)] [string] $Abteilung,
        [Parameter(Mandatory = $true)] [string] $Tätigkeit,
        [Parameter(Mandatory = $true)] [string] $GenPW,
        [Parameter(Mandatory = $true)] [string] $email,
        [Parameter(Mandatory = $false)] [string] $Titel1,
        [Parameter(Mandatory = $false)] [string] $Titel2
    )

    # Bereinige und erstelle den Beekeeper-Benutzernamen
    #$Beeusername = $ADusername -replace '\.|-|ó', '_' -replace '[äöüß]', '${_}' -replace '\s', ''
     $Beeusername = Convert-Umlauts $ADusername

    # Erstelle das Body-Objekt für die Anfrage
    $body = @{
        name         = $Beeusername
        language     = 'de'
        role         = 'member'
        tenantuserid = "SANATORIUM\$ADusername"
        suspended    = $false
        password     = ""
        email        = $email
        custom_fields = @(
            @{key = 'firstname'; value = "$Titel1 $Vorname"}
            @{key = 'lastname'; value = "$Familienname $Titel2"}
            @{key = 'abteilung'; value = $Abteilung}
            @{key = 'position'; value = $Tätigkeit}
            @{key = 'personalnr'; value = $PersNr}
            @{key = 'kostenstelle'; value = $KSTNummer}
        )
    } | ConvertTo-Json -Depth 3

    $body
    #pause

    try {
        # Sende die Anfrage, um einen neuen Beekeeper-Benutzer zu erstellen
        $create = Invoke-RestMethod -Uri 'https://privatklinik-hochrum.de.beekeeper.io/api/2/users' `
            -Method 'POST' -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))

        Write-Host "Beekeeper User mit der ID: $Beeusername wurde erstellt" -ForegroundColor Green

        # Aktualisiere den Avatar des Benutzers
        PKH-UpdateBeekeperAvatar (PKH-GetBeekeeperID $Beeusername) $Beeusername
    } catch {
        # Behandle Fehler und gebe eine Fehlermeldung aus
        $Exception = PKH-ParseErrorForResponseBody $_
        $Exception = $Exception.error
        Write-Host "PKH-CreateBeekeeperProfile: $Exception" -ForegroundColor Red
        return $Exception
    }
}

function PKH-DeleteBeekeeperProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $ID
    )

    try {
        # Senden der Anfrage, um den Beekeeper-Benutzer zu löschen
        $response = Invoke-RestMethod -Uri "https://privatklinik-hochrum.de.beekeeper.io/api/2/users/$ID" `
            -Method Delete -Headers $headers

        # Gebe eine Erfolgsmeldung aus
        Write-Host "Beekeeper user with ID $ID has been deleted." -ForegroundColor Green

        return $true
    } catch {
        # Behandle Fehler und gebe eine Fehlermeldung aus
        $Exception = PKH-ParseErrorForResponseBody $_
        Write-Host "PKH-DeleteBeekeeperProfile error: $Exception" -ForegroundColor Red

        return $false
    }
}

function PKH-ParseErrorForResponseBody($PKHError) {
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        if ($PKHError.Exception.Response) {  
            $Reader = New-Object System.IO.StreamReader($PKHError.Exception.Response.GetResponseStream())
            $Reader.BaseStream.Position = 0
            $Reader.DiscardBufferedData()
            $ResponseBody = $Reader.ReadToEnd()
            if ($ResponseBody.StartsWith('{')) {
                $ResponseBody = $ResponseBody | ConvertFrom-Json
            }
            return $ResponseBody
        }
    }
    else {
        return $PKHError.ErrorDetails.Message
    }
}

function Send-BeekeeperChatMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,
        [Parameter(Mandatory=$true)]
        [string]$MessageBody,
        [Parameter()]
        [string]$AttachmentJson = $null
    )

    $url = "https://privatklinik-hochrum.de.beekeeper.io/api/2/chats/users/$UserId/messages"


    $body = @{
        body = $MessageBody
    }

    if ($AttachmentJson) {
        $body.attachment = ConvertFrom-Json $AttachmentJson
    }

   # $response = Invoke-RestMethod -Method "Post" -Uri $url -Headers $headers -Body ($body | ConvertTo-Json)

    if ($response) {
        Write-Host "Message sent to user $UserId"
    }
}

function Get-ImageBase64FromUrl([Uri]$url) {
    $b = Invoke-WebRequest $url -UseBasicParsing

    $type = $b.Headers["Content-Type"];
    $base64 = [convert]::ToBase64String($b.Content);

    return "$base64";
}

function PKH-CreateBeekeeperQR {
    param (
        [string]$UserID
    )

        # Sende die Anfrage, um einen neuen Beekeeper-Benutzer zu erstellen
        $createQR = Invoke-RestMethod -Uri "https://privatklinik-hochrum.de.beekeeper.io/api/2/tokens/$UserID" -Method 'POST' -Headers $headers
        #Write-Host "Beekeeper User mit der ID: $Beeusername wurde erstellt" -ForegroundColor Green
       #$createQR.qr_url
       $QR = Get-ImageBase64FromUrl($createQR.qr_url)

       return $QR

}

function Convert-MhtToPdf {
    param (
        [string]$sourceFilePath,
        [string]$destinationFilePath,
        [string]$UserID
    )

    $QRCODE = PKH-CreateBeekeeperQR $UserID
    # Read the content of the MHT file
    $content = Get-Content -Path $sourceFilePath -Raw




    $content = $content.Replace('@@VollnamemitTitel',$Vollname).replace('@@employeeID',$PersNr).replace('@@InitPassword',$GenPW).replace('@@Login',$Login).replace('@@MailAdress',$Login+"@pk-hochrum.com").Replace("@@QRCode",$QRCODE)

    # Add more search and replace operations if needed
    # $findText2 = "Variable2"
    # $replaceText2 = "Replacement2"
    # $content = $content.Replace($findText2, $replaceText2)

    # Save the modified content to a new temporary file
    $tempFilePath = [System.IO.Path]::GetTempFileName()
    $content | Out-File -FilePath $tempFilePath -Encoding default

    # Use Word to export the temporary file to PDF
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $docTemp = $wordApp.Documents.Open($tempFilePath)
    $docTemp.ExportAsFixedFormat($destinationFilePath, 17)  # 17 is the value for PDF format
    $docTemp.Close()
    $wordApp.Quit()

    # Clean up the temporary file
     #$tempFilePath
    Remove-Item $tempFilePath -Force

    Write-Host "MHT to PDF conversion completed. PDF saved at: $destinationFilePath"
}

function PKH-LoginPDF($UserID,$PersNr){

    $sourceFilePath = "D:\ADSync\LoginInstruction\MA.mht"
    $destinationFilePath = "D:\ADSync\LoginInstruction\out\$PersNr.pdf"

    # Replace variables in the MHT file and save the modified content as a new PDF
    Convert-MhtToPdf -sourceFilePath $sourceFilePath -destinationFilePath $destinationFilePath -UserID $UserID

    return $destinationFilePath
    
}

#PKH-InitBeekeeperHeader
PKH-AdSync
Stop-Transcript