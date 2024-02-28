# Charger le module Active Directory si pas déjà chargé
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}

# Fonction pour remplacer les caractères accentués
function Remove-Diacritics {
    param (
        [string]$text
    )
    $normalized = $text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    $normalized.ToCharArray() | ForEach-Object {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($_)
        }
    }
    return $builder.ToString()
}

# Fonction pour récupérer le suffixe UPN principal du domaine actuel
function Get-DomainUPNSuffix {
    $domain = Get-ADDomain
    return $domain.DNSRoot
}

# Fonction pour récupérer les suffixes UPN configurés au niveau de la forêt
function Get-ForestUPNSuffixes {
    $forest = Get-ADForest
    return $forest.UPNSuffixes
}

# Fonction pour sélectionner le fichier CSV
function Select-CSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $result = $openFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    } else {
        return $null
    }
}

# Fonction pour sélectionner l'OU et saisir le mot de passe
function Select-ADOUAndPassword {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Sélection de l''OU et saisie du mot de passe'
    $form.Size = New-Object System.Drawing.Size(520, 350)
    $form.StartPosition = 'CenterScreen'


    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(480, 20)
    $label.Text = 'Entrer le nom de L''OU à rechercher :'
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 35)
    $textBox.Size = New-Object System.Drawing.Size(480, 20)
    $form.Controls.Add($textBox)

    $ouLabel = New-Object System.Windows.Forms.Label
    $ouLabel.Location = New-Object System.Drawing.Point(10, 60)
    $ouLabel.Size = New-Object System.Drawing.Size(480, 20)
    $ouLabel.Text = 'Sélectionner l''OU de destination :'
    $form.Controls.Add($ouLabel)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 85)
    $listBox.Size = New-Object System.Drawing.Size(480, 200)
    $listBox.BackColor = 'White'
    $form.Controls.Add($listBox)

    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Location = New-Object System.Drawing.Point(10, 290)
    $passwordLabel.Size = New-Object System.Drawing.Size(480, 20)
    $passwordLabel.Text = 'Entrer le mot de passe à appliquer :'
    $form.Controls.Add($passwordLabel)

    $passwordBox = New-Object System.Windows.Forms.TextBox
    $passwordBox.Location = New-Object System.Drawing.Point(10, 315)
    $passwordBox.Size = New-Object System.Drawing.Size(480, 20)
    $form.Controls.Add($passwordBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(390, 345)
    $okButton.Size = New-Object System.Drawing.Size(100, 25)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $textBox.Add_TextChanged({
        $listBox.Items.Clear()
        $searchText = $textBox.Text
        if ($searchText.Length -gt 0) {
            $OUs = Get-ADOrganizationalUnit -Filter "Name -like '*$searchText*'" -Properties Name | Select-Object -ExpandProperty DistinguishedName
            foreach ($ou in $OUs) {
                $listBox.Items.Add($ou)
            }
        }
    })

    $form.ShowDialog() | Out-Null

    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $listBox.SelectedItem, $passwordBox.Text
    } else {
        return $null, $null
    }
}

# Fonction pour sélectionner le suffixe UPN
function Select-UPNSuffix {
    Add-Type -AssemblyName System.Windows.Forms
    $formUPN = New-Object System.Windows.Forms.Form
    $formUPN.Text = 'Sélection du suffixe UPN'
    $formUPN.Size = New-Object System.Drawing.Size(320, 200)
    $formUPN.StartPosition = 'CenterScreen'

    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(10, 10)
    $comboBox.Size = New-Object System.Drawing.Size(290, 20)
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $domainUPNSuffix = Get-DomainUPNSuffix
    $comboBox.Items.Add($domainUPNSuffix)

    $upnSuffixes = Get-ForestUPNSuffixes
    foreach ($suffix in $upnSuffixes) {
        if (-not [string]::IsNullOrWhiteSpace($suffix) -and $suffix -ne $domainUPNSuffix) {
            $comboBox.Items.Add($suffix)
        }
    }

    $formUPN.Controls.Add($comboBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(220, 130)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formUPN.Controls.Add($okButton)
    $formUPN.AcceptButton = $okButton

    $formUPN.ShowDialog() | Out-Null

    if ($formUPN.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $comboBox.SelectedItem.ToString().Trim()
    } else {
        return $null
    }
}
# Début du script principal

$csvPath = Select-CSVFile
if ($csvPath -eq $null) {
    Write-Host "Aucun fichier CSV sélectionné, arrêt du script."
    exit
}

$selectedOU, $password = Select-ADOUAndPassword
if ($selectedOU -eq $null -or $password -eq $null) {
    Write-Host "OU non sélectionné ou mot de passe non fourni, arrêt du script."
    exit
}

$selectedUPNSuffix = Select-UPNSuffix
if ($selectedUPNSuffix -eq $null) {
    Write-Host "Suffixe UPN non sélectionné, arrêt du script."
    exit
}

# Assurez-vous que le suffixe UPN est correct
if ($selectedUPNSuffix -match '(\S+)') {
    $selectedUPNSuffix = $matches[1]
}

# Vérifiez si le suffixe UPN n'est pas null avant d'accéder à l'index
if ($selectedUPNSuffix -ne $null) {
    # Affichez le suffixe UPN pour la vérification
    Write-Host "Suffixe UPN sélectionné pour la création de l'utilisateur: $selectedUPNSuffix"
} else {
    Write-Host "Suffixe UPN non sélectionné, arrêt du script."
    exit
}

$users = Import-Csv -Path $csvPath -Delimiter ';'
foreach ($user in $users) {
    $formattedGivenName = ($user.Prénom.Substring(0,1).ToUpper() + $user.Prénom.Substring(1).ToLower()).Replace(" ", "")
    $formattedSurname = $user.Nom.Replace(" ", "").ToUpper()

    # Utiliser la fonction pour supprimer les accents
    $cleanGivenName = Remove-Diacritics -text $user.Prénom
    $cleanSurname = Remove-Diacritics -text $user.Nom

    $samAccountName = ("{0}.{1}" -f $cleanGivenName.Replace(" ", "").ToLower(), $cleanSurname.Replace(" ", "").ToLower()).Substring(0,[Math]::Min(20, $cleanGivenName.Length + $cleanSurname.Length))
    
    # Nettoyez le UPN si nécessaire
    $fullUPN = "$samAccountName@$selectedUPNSuffix"
    $cleanUPN = $fullUPN -replace '0 1 ', ''  # Enlève '0 1 ' s'il apparaît

    $ADUserParams = @{
        Enabled               = $true
        Path                  = $selectedOU
        AccountPassword       = (ConvertTo-SecureString -AsPlainText $password -Force)
        PasswordNeverExpires  = $false
        ChangePasswordAtLogon = $true
        UserPrincipalName     = $cleanUPN
        SamAccountName        = $samAccountName
        Name                  = "$formattedGivenName $formattedSurname"
        GivenName             = $formattedGivenName
        Surname               = $formattedSurname
        DisplayName           = "$formattedGivenName $formattedSurname"
        EmailAddress          = $cleanUPN
        City                  = $user.Ville
        PostalCode            = $user.'Code postal'
        State                 = $user.État
        Title                 = $user.Titre
        Department            = $user.Service
    }

    try {
        New-ADUser @ADUserParams -ErrorAction Stop
        Write-Host "Utilisateur créé : $($ADUserParams['GivenName']) $($ADUserParams['Surname'])" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de la création de l'utilisateur : $($ADUserParams['GivenName']) $($ADUserParams['Surname'])" -ForegroundColor Red
        Write-Host "Détail de l'erreur : $($_.Exception.Message)" -ForegroundColor Red
    }
}


# Optionnel : Déclenchez une synchronisation avec Microsoft 365 si nécessaire
# Start-ADSyncSyncCycle -PolicyType Delta

Write-Host "Script terminé." -ForegroundColor Yellow
