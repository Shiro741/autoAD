Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

function CapitalizeLastName {
    $textboxNom.Text = $textboxNom.Text.ToUpper()
}

function CapitalizeFirstName {
    if (-not [string]::IsNullOrWhiteSpace($textboxPrenom.Text)) {
        $textboxPrenom.Text = $textboxPrenom.Text.Substring(0,1).ToUpper() + $textboxPrenom.Text.Substring(1).ToLower()
    }
}

function ValidatePassword {
    if ($textboxPassword.Text -ne $textboxConfirmPassword.Text) {
        [System.Windows.Forms.MessageBox]::Show("Les mots de passe ne correspondent pas. Veuillez réessayer.", "Erreur de confirmation de mot de passe", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}

function Get-ADOrganizationalUnits {
    $ous = @{}
    $searcher = [adsisearcher]"(objectCategory=organizationalUnit)"
    $searcher.FindAll() | ForEach-Object {
        $ouName = $_.Properties.name[0]
        $ouDistinguishedName = $_.Properties.distinguishedname[0]
        $ous[$ouName] = $ouDistinguishedName
    }
    return $ous
}

function ValidateManager {
    param (
        [string]$commonName
    )

    try {
        $managers = Get-ADUser -Filter "CN -eq '$commonName'" -Properties * -ErrorAction Stop
        foreach ($manager in $managers) {
            Write-Host "Manager trouvé :"
            $manager | Format-List * | Out-String | Write-Host
        }
        return $managers
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la récupération des managers : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

function ValidateUserUniqueness {
    param (
        [string]$userPrincipalName
    )

    try {
        $existingUser = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -ErrorAction Stop
        if ($existingUser) {
            [System.Windows.Forms.MessageBox]::Show("Un utilisateur avec ce UPN existe déjà. Veuillez choisir un autre nom.", "Erreur de validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    } catch {
        return $true
    }
    return $true
}

function ShowExpirationForm {
    param (
        [string]$userPrincipalName,
        [string]$description,
        [datetime]$dateArrivee
    )

    $formExpiration = New-Object System.Windows.Forms.Form
    $formExpiration.Text = "Définir la date d'expiration"
    $formExpiration.Size = New-Object System.Drawing.Size(300, 200)
    $formExpiration.StartPosition = "CenterParent"

    $labelExpiration = New-Object System.Windows.Forms.Label
    $labelExpiration.Text = "Date d'expiration (facultatif) :"
    $labelExpiration.Location = New-Object System.Drawing.Point(10, 20)
    $labelExpiration.AutoSize = $true
    $formExpiration.Controls.Add($labelExpiration)

    $datePickerExpiration = New-Object System.Windows.Forms.DateTimePicker
    $datePickerExpiration.Location = New-Object System.Drawing.Point(10, 50)
    $datePickerExpiration.Size = New-Object System.Drawing.Size(260, 20)
    $formExpiration.Controls.Add($datePickerExpiration)

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Text = "OK"
    $buttonOk.Location = New-Object System.Drawing.Point(110, 90)
    $buttonOk.Size = New-Object System.Drawing.Size(75, 25)
    $buttonOk.Add_Click({
        $expirationDate = $datePickerExpiration.Value
        $formExpiration.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $formExpiration.Close()

        try {
            $user = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -SearchBase $selectedOUPath -Properties DistinguishedName -ErrorAction Stop
            if ($null -eq $user) {
                throw "Utilisateur introuvable : $userPrincipalName"
            }

            $dn = $user.DistinguishedName
            $expirationDescription = ""
            if ($expirationDate -gt [datetime]::MinValue) {
                $expirationDescription = "Expire le : $($expirationDate.ToString('dd/MM/yyyy'))"
                Set-ADUser -Identity $dn -AccountExpirationDate $expirationDate -Description "$description, $expirationDescription"
            } else {
                Set-ADUser -Identity $dn -Description "$description"
            }
            [System.Windows.Forms.MessageBox]::Show("Date d'expiration définie pour l'utilisateur $userPrincipalName.", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de la définition de la date d'expiration : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $formExpiration.Controls.Add($buttonOk)

    $formExpiration.ShowDialog()
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Formulaire d'ajout d'utilisateur CSTB"
$form.Size = New-Object System.Drawing.Size(400, 700)
$form.StartPosition = "CenterScreen"

$labelNom = New-Object System.Windows.Forms.Label
$labelNom.Text = "Nom :"
$labelNom.Location = New-Object System.Drawing.Point(10, 20)
$labelNom.AutoSize = $true
$form.Controls.Add($labelNom)

$textboxNom = New-Object System.Windows.Forms.TextBox
$textboxNom.Location = New-Object System.Drawing.Point(150, 20)
$textboxNom.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxNom)

$labelPrenom = New-Object System.Windows.Forms.Label
$labelPrenom.Text = "Prénom :"
$labelPrenom.Location = New-Object System.Drawing.Point(10, 60)
$labelPrenom.AutoSize = $true
$form.Controls.Add($labelPrenom)

$textboxPrenom = New-Object System.Windows.Forms.TextBox
$textboxPrenom.Location = New-Object System.Drawing.Point(150, 60)
$textboxPrenom.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxPrenom)

$labelTypeContrat = New-Object System.Windows.Forms.Label
$labelTypeContrat.Text = "Type de contrat :"
$labelTypeContrat.Location = New-Object System.Drawing.Point(10, 100)
$labelTypeContrat.AutoSize = $true
$form.Controls.Add($labelTypeContrat)

$dropdownTypeContrat = New-Object System.Windows.Forms.ComboBox
$dropdownTypeContrat.Location = New-Object System.Drawing.Point(150, 100)
$dropdownTypeContrat.Size = New-Object System.Drawing.Size(200, 20)
$dropdownTypeContrat.Items.Add("CDI")
$dropdownTypeContrat.Items.Add("CDD")
$dropdownTypeContrat.Items.Add("Intérim")
$dropdownTypeContrat.Items.Add("Externe")
$dropdownTypeContrat.Add_SelectedIndexChanged({
    if ($dropdownTypeContrat.SelectedItem -eq "CDI") {
        $datePickerArrivalDate.Enabled = $false
        $datePickerArrivalDate.Value = [datetime]::MinValue
    } else {
        $datePickerArrivalDate.Enabled = $true
    }
})
$form.Controls.Add($dropdownTypeContrat)

$labelArrivalDate = New-Object System.Windows.Forms.Label
$labelArrivalDate.Text = "Date d'arrivée :"
$labelArrivalDate.Location = New-Object System.Drawing.Point(10, 140)
$labelArrivalDate.AutoSize = $true
$form.Controls.Add($labelArrivalDate)

$datePickerArrivalDate = New-Object System.Windows.Forms.DateTimePicker
$datePickerArrivalDate.Location = New-Object System.Drawing.Point(150, 140)
$datePickerArrivalDate.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($datePickerArrivalDate)

$labelOU = New-Object System.Windows.Forms.Label
$labelOU.Text = "Unité Organisationnelle :"
$labelOU.Location = New-Object System.Drawing.Point(10, 180)
$labelOU.AutoSize = $true
$form.Controls.Add($labelOU)

$dropdownOU = New-Object System.Windows.Forms.ComboBox
$dropdownOU.Location = New-Object System.Drawing.Point(150, 180)
$dropdownOU.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($dropdownOU)

$ouDictionary = Get-ADOrganizationalUnits
$dropdownOU.Items.AddRange($ouDictionary.Keys)

$dropdownOU.Add_SelectedIndexChanged({
    $selectedOUName = $dropdownOU.SelectedItem
    $selectedOUPath = $ouDictionary[$selectedOUName]
})

$labelPassword = New-Object System.Windows.Forms.Label
$labelPassword.Text = "Mot de passe :"
$labelPassword.Location = New-Object System.Drawing.Point(10, 220)
$labelPassword.AutoSize = $true
$form.Controls.Add($labelPassword)

$textboxPassword = New-Object System.Windows.Forms.TextBox
$textboxPassword.Location = New-Object System.Drawing.Point(150, 220)
$textboxPassword.Size = New-Object System.Drawing.Size(200, 20)
$textboxPassword.PasswordChar = "*"
$form.Controls.Add($textboxPassword)

$labelConfirmPassword = New-Object System.Windows.Forms.Label
$labelConfirmPassword.Text = "Confirmation :"
$labelConfirmPassword.Location = New-Object System.Drawing.Point(10, 260)
$labelConfirmPassword.AutoSize = $true
$form.Controls.Add($labelConfirmPassword)

$textboxConfirmPassword = New-Object System.Windows.Forms.TextBox
$textboxConfirmPassword.Location = New-Object System.Drawing.Point(150, 260)
$textboxConfirmPassword.Size = New-Object System.Drawing.Size(200, 20)
$textboxConfirmPassword.PasswordChar = "*"
$form.Controls.Add($textboxConfirmPassword)

$labelVille = New-Object System.Windows.Forms.Label
$labelVille.Text = "Ville :"
$labelVille.Location = New-Object System.Drawing.Point(10, 300)
$labelVille.AutoSize = $true
$form.Controls.Add($labelVille)

$dropdownVille = New-Object System.Windows.Forms.ComboBox
$dropdownVille.Location = New-Object System.Drawing.Point(150, 300)
$dropdownVille.Size = New-Object System.Drawing.Size(200, 20)
$dropdownVille.Items.Add("MLV")
$dropdownVille.Items.Add("GRE")
$dropdownVille.Items.Add("NAN")
$dropdownVille.Items.Add("SOP")
$form.Controls.Add($dropdownVille)

$labelFonction = New-Object System.Windows.Forms.Label
$labelFonction.Text = "Fonction :"
$labelFonction.Location = New-Object System.Drawing.Point(10, 340)
$labelFonction.AutoSize = $true
$form.Controls.Add($labelFonction)

$textboxFonction = New-Object System.Windows.Forms.TextBox
$textboxFonction.Location = New-Object System.Drawing.Point(150, 340)
$textboxFonction.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxFonction)

$labelDirection = New-Object System.Windows.Forms.Label
$labelDirection.Text = "Direction :"
$labelDirection.Location = New-Object System.Drawing.Point(10, 380)
$labelDirection.AutoSize = $true
$form.Controls.Add($labelDirection)

$textboxDirection = New-Object System.Windows.Forms.TextBox
$textboxDirection.Location = New-Object System.Drawing.Point(150, 380)
$textboxDirection.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxDirection)

$labelEntreprise = New-Object System.Windows.Forms.Label
$labelEntreprise.Text = "Entreprise :"
$labelEntreprise.Location = New-Object System.Drawing.Point(10, 420)
$labelEntreprise.AutoSize = $true
$form.Controls.Add($labelEntreprise)

$textboxEntreprise = New-Object System.Windows.Forms.TextBox
$textboxEntreprise.Location = New-Object System.Drawing.Point(150, 420)
$textboxEntreprise.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxEntreprise)

$labelManager = New-Object System.Windows.Forms.Label
$labelManager.Text = "Manager :"
$labelManager.Location = New-Object System.Drawing.Point(10, 460)
$labelManager.AutoSize = $true
$form.Controls.Add($labelManager)

$textboxManager = New-Object System.Windows.Forms.TextBox
$textboxManager.Location = New-Object System.Drawing.Point(150, 460)
$textboxManager.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxManager)

$buttonCheckManager = New-Object System.Windows.Forms.Button
$buttonCheckManager.Location = New-Object System.Drawing.Point(360, 460)
$buttonCheckManager.Size = New-Object System.Drawing.Size(25, 20)
$buttonCheckManager.Text = "?"
$buttonCheckManager.Add_Click({
    $managerCommonName = $textboxManager.Text
    if ($managerCommonName -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Veuillez entrer le nom commun du manager.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $managers = ValidateManager -commonName $managerCommonName
    if ($managers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun manager trouvé pour le nom commun : $managerCommonName", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } elseif ($managers.Count -eq 1) {
        $manager = $managers[0]
        $managerCN = $manager.CN
        $textboxManager.Text = $managerCN
        [System.Windows.Forms.MessageBox]::Show("Manager trouvé : $($manager.Name), CN: $managerCN", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        # Logique pour plusieurs managers trouvés, similaire à l'existant
    }
})

$form.Controls.Add($buttonCheckManager)

$buttonValider = New-Object System.Windows.Forms.Button
$buttonValider.Location = New-Object System.Drawing.Point(150, 500)
$buttonValider.Size = New-Object System.Drawing.Size(100, 30)
$buttonValider.Text = "Valider"
$buttonValider.Add_Click({
    CapitalizeLastName
    CapitalizeFirstName

    if (-not (ValidatePassword)) {
        return
    }

    $nom = $textboxNom.Text
    $prenom = $textboxPrenom.Text
    $typeContrat = $dropdownTypeContrat.SelectedItem

    $nomUtilisateur = "$prenom.$nom"
    $userPrincipalName = "$prenom.$nom@cstb.fr"
    $displayName = "$prenom $nom"

    if ($typeContrat -eq "Externe") {
        $nomExterne = "${nom}_Externe"
        $userPrincipalName = "$prenom.$nomExterne@cstb.fr"
    }

    if (-not (ValidateUserUniqueness -userPrincipalName $userPrincipalName)) {
        return
    }

    try {
        $selectedOUName = $dropdownOU.SelectedItem
        $selectedOUPath = $ouDictionary[$selectedOUName]
        $motDePasse = $textboxPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $dateArrivee = $datePickerArrivalDate.Value
        $formattedArrivalDate = $dateArrivee.ToString("dd/MM/yyyy")
        $ville = $dropdownVille.SelectedItem
        $fonction = $textboxFonction.Text
        $direction = $textboxDirection.Text
        $entreprise = $textboxEntreprise.Text
        $description = "$typeContrat, Date d'arrivée : $formattedArrivalDate"
        $managerCN = $textboxManager.Text  # Utiliser le CN du champ Manager

        if (-not $selectedOUPath) {
            throw "Unité Organisationnelle non sélectionnée ou introuvable."
        }

        if (-not $managerCN) {
            throw "Manager non valide ou introuvable."
        }

        Write-Host "Creating user in OU: $selectedOUPath"
        Write-Host "UserPrincipalName: $userPrincipalName"
        Write-Host "SamAccountName: $prenom.$nom"
        
        New-ADUser -Name $nomUtilisateur -GivenName $prenom -SamAccountName "$prenom.$nom" -Surname $nom -AccountPassword $motDePasse -UserPrincipalName $userPrincipalName -Enabled $true -DisplayName $displayName -Description $description -Path $selectedOUPath -City $ville -Title $fonction -Department $direction -Company $entreprise -Manager $managerCN -ChangePasswordAtLogon $true
        
        Write-Host "User $userPrincipalName created. Waiting for replication..."

        # Temps d'attente pour la réplication
        Start-Sleep -Seconds 15

        $maxRetries = 5
        $retryCount = 0
        $userFound = $false

        while (-not $userFound -and $retryCount -lt $maxRetries) {
            try {
                Write-Host "Searching for user: $userPrincipalName in $selectedOUPath, attempt $($retryCount + 1)"
                $user = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -SearchBase $selectedOUPath -Properties DistinguishedName, Description -ErrorAction Stop
                if ($null -ne $user) {
                    $userFound = $true
                    Write-Host "User $nomUtilisateur found in $selectedOUPath on attempt $($retryCount + 1)"
                }
            } catch {
                Write-Host "User not found in $selectedOUPath, retrying in 10 seconds..."
                Start-Sleep -Seconds 10
                $retryCount++
            }
        }

        if (-not $userFound) {
            throw "Utilisateur $nomUtilisateur introuvable après $maxRetries tentatives dans $selectedOUPath."
        }

        if ($null -eq $user) {
            throw "L'utilisateur $userPrincipalName n'a pas été trouvé ou est $null."
        }

        Write-Host "User object: $($user | Out-String)"

        # Mise à jour de la description pour inclure type de contrat, date d'arrivée et date d'expiration
        $updatedDescription = $description
        if ($user.Description -notlike "*Date d'arrivée*") {
            $updatedDescription = "$user.Description, $description"
        } elseif ($user.Description -match "Date d'arrivée : \d{2}/\d{2}/\d{4}") {
            $updatedDescription = [regex]::Replace($user.Description, "Date d'arrivée : \d{2}/\d{2}/\d{4}", "Date d'arrivée : $formattedArrivalDate")
        }

        Set-ADUser -Identity $user.DistinguishedName -Description $updatedDescription

        # Attendre 10 secondes avant d'afficher le formulaire de départ
        Start-Sleep -Seconds 10

        if ($typeContrat -ne "CDI") {
            try {
                ShowExpirationForm -userPrincipalName $userPrincipalName -description $updatedDescription -dateArrivee $dateArrivee
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Erreur lors de la définition de la date d'expiration : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Utilisateur $nomUtilisateur créé avec succès dans l'OU $selectedOUName.", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la création de l'utilisateur : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($buttonValider)

$form.ShowDialog() | Out-Null




############################################ Partie serveur de mail ############################################

# Chemin du script d'envoi d'email
$emailScriptPath = "C:\Scripts\SendWelcomeEmail.ps1"

# Définir la variable $nomExterne
$nomExterne = "${nom}_Externe"

# Informations pour l'email
if ($typeContrat -eq "Externe") {
    $emailTo = "$prenom.$nomExterne@cstb.fr"
} else {
    $emailTo = "$prenom.$nom@cstb.fr"
}
$emailSubject = "Bienvenue dans l'entreprise"
$emailBody = @"
Bonjour $prenom,

Bienvenue dans l'entreprise CSTB. 
Votre compte a été créé avec succès.


Cordialement,
L'équipe CSU
"@

# Enregistrer les informations de l'email dans un fichier pour le script différé
$emailInfoPath = "C:\Scripts\EmailInfo_$($nomUtilisateur).xml"
$emailInfo = @{
    To = $emailTo
    Subject = $emailSubject
    Body = $emailBody
}
$emailInfo | Export-Clixml -Path $emailInfoPath

# Créer la tâche planifiée pour envoyer l'email dans 24 heures
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File $emailScriptPath -EmailInfoPath $emailInfoPath"
$trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddHours(24))
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "SendWelcomeEmail_$($nomUtilisateur)" -Description "Envoie un email de bienvenue à $nomUtilisateur 24 heures après la création de son compte"

