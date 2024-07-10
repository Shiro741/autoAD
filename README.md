# autoAD
# add user with powershell app

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

function ShowExpirationForm {
    param (
        [string]$userPrincipalName,
        [string]$description,
        [datetime]$dateArrivee
    )

    $expirationForm = New-Object System.Windows.Forms.Form
    $expirationForm.Text = "Définir la date d'expiration"
    $expirationForm.Size = New-Object System.Drawing.Size(300, 200)
    $expirationForm.StartPosition = "CenterScreen"

    $labelExpirationDate = New-Object System.Windows.Forms.Label
    $labelExpirationDate.Text = "Date de départ :"
    $labelExpirationDate.Location = New-Object System.Drawing.Point(10, 20)
    $labelExpirationDate.AutoSize = $true
    $expirationForm.Controls.Add($labelExpirationDate)

    $datePickerExpirationDate = New-Object System.Windows.Forms.DateTimePicker
    $datePickerExpirationDate.Location = New-Object System.Drawing.Point(150, 20)
    $datePickerExpirationDate.Size = New-Object System.Drawing.Size(120, 20)
    $expirationForm.Controls.Add($datePickerExpirationDate)

    $buttonSaveExpiration = New-Object System.Windows.Forms.Button
    $buttonSaveExpiration.Location = New-Object System.Drawing.Point(100, 60)
    $buttonSaveExpiration.Size = New-Object System.Drawing.Size(100, 30)
    $buttonSaveExpiration.Text = "Enregistrer"
    $buttonSaveExpiration.Add_Click({
        $expirationDate = $datePickerExpirationDate.Value
        $formattedArrivalDate = $dateArrivee.ToString("dd/MM/yyyy")
        $formattedDepartureDate = $expirationDate.ToString("dd/MM/yyyy")
        # Vérification et suppression de l'ancienne date d'arrivée dans la description
        $updatedDescription = $description -replace ",? Date d'arrivée : \d{2}/\d{2}/\d{4}", ""
        $updatedDescription = "$updatedDescription, Date d'arrivée : $formattedArrivalDate, Date de départ : $formattedDepartureDate"
        SetAccountExpiration -expirationDate $expirationDate -userPrincipalName $userPrincipalName -description $updatedDescription
        [System.Windows.Forms.MessageBox]::Show("Date d'expiration mise à jour avec succès.", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $expirationForm.Close()
    })
    $expirationForm.Controls.Add($buttonSaveExpiration)

    $expirationForm.ShowDialog() | Out-Null
}

function SetAccountExpiration {
    param (
        [datetime]$expirationDate,
        [string]$userPrincipalName,
        [string]$description
    )

    if ($expirationDate -ne [datetime]::MinValue) {
        try {
            $user = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'"
            if ($user) {
                Set-ADUser -Identity $user.DistinguishedName -AccountExpirationDate $expirationDate -Description $description
            } else {
                throw "Utilisateur non trouvé pour définir la date d'expiration."
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de la définition de la date d'expiration du compte : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function ValidateUserUniqueness {
    param (
        [string]$userPrincipalName
    )

    $user = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'"
    if ($user) {
        [System.Windows.Forms.MessageBox]::Show("Le nom d'utilisateur $userPrincipalName existe déjà. Veuillez choisir un autre nom.", "Erreur de validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}

function ValidateManager {
    param (
        [string]$lastName
    )

    $managers = Get-ADUser -Filter "Surname -eq '$lastName'"
    return $managers
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Formulaire d'ajout d'utilisateur"
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

$ouDictionary = Get-ADOrganizationalUnits
$dropdownOU.Items.AddRange($ouDictionary.Keys)

$dropdownOU.Add_SelectedIndexChanged({
    $selectedOUName = $dropdownOU.SelectedItem
    $selectedOUPath = $ouDictionary[$selectedOUName]
})

$form.Controls.Add($dropdownOU)

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
$dropdownVille.Items.Add("Ville 1")
$dropdownVille.Items.Add("Ville 2")
$dropdownVille.Items.Add("Ville 3")
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
    $managerInput = $textboxManager.Text
    $managers = ValidateManager -lastName $managerInput
    if ($managers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun manager trouvé pour le nom de famille : $managerInput", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } elseif ($managers.Count -eq 1) {
        $manager = $managers[0]
        $textboxManager.Text = $manager.UserPrincipalName
        [System.Windows.Forms.MessageBox]::Show("Manager trouvé : $($manager.Name)", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        # Afficher une liste de choix si plusieurs managers sont trouvés
        $listForm = New-Object System.Windows.Forms.Form
        $listForm.Text = "Sélectionner un Manager"
        $listForm.Size = New-Object System.Drawing.Size(400, 300)
        $listForm.StartPosition = "CenterParent"

        $labelList = New-Object System.Windows.Forms.Label
        $labelList.Text = "Sélectionnez un manager :"
        $labelList.Location = New-Object System.Drawing.Point(10, 20)
        $labelList.AutoSize = $true
        $listForm.Controls.Add($labelList)

        $listBoxManagers = New-Object System.Windows.Forms.ListBox
        $listBoxManagers.Location = New-Object System.Drawing.Point(10, 50)
        $listBoxManagers.Size = New-Object System.Drawing.Size(360, 150)
        $managers | ForEach-Object {
            $listBoxManagers.Items.Add($_.Name + " (" + $_.UserPrincipalName + ")")
        }
        $listForm.Controls.Add($listBoxManagers)

        $buttonSelectManager = New-Object System.Windows.Forms.Button
        $buttonSelectManager.Location = New-Object System.Drawing.Point(150, 220)
        $buttonSelectManager.Size = New-Object System.Drawing.Size(100, 30)
        $buttonSelectManager.Text = "Sélectionner"
        $buttonSelectManager.Add_Click({
            if ($listBoxManagers.SelectedItem -ne $null) {
                $selectedIndex = $listBoxManagers.SelectedIndex
                $selectedManager = $managers[$selectedIndex]
                $textboxManager.Text = $selectedManager.UserPrincipalName
                $textboxManager.Tag = $selectedManager.DistinguishedName
                [System.Windows.Forms.MessageBox]::Show("Manager sélectionné : $($selectedManager.Name)", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $listForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un manager.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
        $listForm.Controls.Add($buttonSelectManager)

        $listForm.ShowDialog() | Out-Null
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
    $nomUtilisateur = "$prenom.$nom"
    $userPrincipalName = "$prenom.$nom@$($env:USERDNSDOMAIN)"
    
    if (-not (ValidateUserUniqueness -userPrincipalName $userPrincipalName)) {
        return
    }

    try {
        $selectedOUName = $dropdownOU.SelectedItem
        $selectedOUPath = $ouDictionary[$selectedOUName]
        $motDePasse = $textboxPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $dateArrivee = $datePickerArrivalDate.Value
        $ville = $dropdownVille.SelectedItem
        $fonction = $textboxFonction.Text
        $direction = $textboxDirection.Text
        $entreprise = $textboxEntreprise.Text
        $typeContrat = $dropdownTypeContrat.SelectedItem
        $description = ""
        $managerUPN = $textboxManager.Tag

        if ($typeContrat -eq "CDI") {
            $description = "Membre de $direction"
        } else {
            $formattedArrivalDate = $dateArrivee.ToString("dd/MM/yyyy")
            $description = "$typeContrat, Date d'arrivée : $formattedArrivalDate"
        }

        if (-not $selectedOUPath) {
            throw "Unité Organisationnelle non sélectionnée ou introuvable."
        }

        if (-not $managerUPN) {
            throw "Manager non valide ou introuvable."
        }

        Write-Host "Creating user in OU: $selectedOUPath"
        New-ADUser -Name $nomUtilisateur -GivenName $prenom -SamAccountName "$prenom.$nom" -Surname $nom -AccountPassword $motDePasse -UserPrincipalName $userPrincipalName -Enabled $true -Description $description -Path $selectedOUPath -City $ville -Title $fonction -Department $direction -Company $entreprise -Manager $managerUPN -ChangePasswordAtLogon $true
        Start-Sleep -Seconds 5  # Wait for AD replication

        $user = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'"
        if ($user) {
            Write-Host "User $nomUtilisateur found"

            if ($typeContrat -ne "CDI") {
                ShowExpirationForm -userPrincipalName $userPrincipalName -description $description -dateArrivee $dateArrivee
            } else {
                [System.Windows.Forms.MessageBox]::Show("Utilisateur $nomUtilisateur créé avec succès dans l'OU $selectedOUName.", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Host "User $nomUtilisateur not found after retries"
            throw "Utilisateur $nomUtilisateur introuvable après création."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la création de l'utilisateur : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($buttonValider)

$form.ShowDialog() | Out-Null
