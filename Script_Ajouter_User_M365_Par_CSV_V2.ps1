# 1) Faire son export CSV depuis M365 en ayant filtrer le type d'user voulut
# 2) Avoir un csv bien délimité: ouvrir excel -> ouvrir csv -> Délimité -> Cocher seulement Virgule -> Standart -> Terminer
# 3) Je vous conseille personnellement de supprimer le contenu de la colonne "Adresses proxy" dans votre CSV export M365 pour ne pas avoir de problèmes de synchro smtp
# 4) Les SuffixUPN de votre domaine vous seront automatiquement affiché, choisissez le bon
# 5) Penser à mettre un MDP en accord avec vos GPO

# Charger le module Active Directory si pas déjà chargé
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}

# remplace accents
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

# récup suffixUPN principal du domaine
function Get-DomainUPNSuffix {
    $domain = Get-ADDomain
    return $domain.DNSRoot
}

# récup suffixUPN configurés en plus
function Get-ForestUPNSuffixes {
    $forest = Get-ADForest
    return $forest.UPNSuffixes
}

# Selection CSV
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

# Selection OU et MDP
function Select-ADOUAndPassword {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Sélection de l''OU et saisie du mot de passe'
    $form.Size = New-Object System.Drawing.Size(520, 430)
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

# Selection SuffixUPN
function Select-UPNSuffix {
    Add-Type -AssemblyName System.Windows.Forms
    $formUPN = New-Object System.Windows.Forms.Form
    $formUPN.Text = 'Sélection du suffixe UPN'
    $formUPN.Size = New-Object System.Drawing.Size(215, 125)
    $formUPN.StartPosition = 'CenterScreen'

    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(10, 10)
    $comboBox.Size = New-Object System.Drawing.Size(180, 20)
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
    $okButton.Location = New-Object System.Drawing.Point(115, 40)
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

# Debut script princip

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

# Continu script si suffix upn
if ($selectedUPNSuffix -ne $null) {
    $users = Import-Csv -Path $csvPath -Delimiter ';'
    foreach ($user in $users) {
        $formattedGivenName = ($user.Prénom.Substring(0,1).ToUpper() + $user.Prénom.Substring(1).ToLower()).Replace(" ", "")
        $formattedSurname = $user.Nom.Replace(" ", "").ToUpper()

        # Utilise fonction suppression accents
        $cleanGivenName = Remove-Diacritics -text $user.Prénom
        $cleanSurname = Remove-Diacritics -text $user.Nom

       
        $samAccountName = ("{0}.{1}" -f $cleanGivenName.Replace(" ", "").ToLower(), $cleanSurname.Replace(" ", "").ToLower())
        if ($samAccountName.Length -gt 20) {
        $samAccountName = $samAccountName.Substring(0, 20)
        }

        # Nettoi UPN si necessaire
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
            Write-Host "Utilisateur créé : $($ADUserParams['SamAccountName'])" -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de la création de l'utilisateur : $_" -ForegroundColor Red
        }
    }  # ferme foreach

    try {
        Start-ADSyncSyncCycle -PolicyType Delta
        Write-Host "Cycle de synchronisation Delta démarré." -ForegroundColor Green
    } catch [System.InvalidOperationException] {
        Write-Host "Un cycle de synchronisation est déjà en cours. Impossible de démarrer un nouveau cycle jusqu'à ce que celui-ci soit terminé." -ForegroundColor Yellow
    } catch {
        Write-Host "Une erreur inattendue s'est produite lors du démarrage du cycle de synchronisation Delta : $_" -ForegroundColor Red
    }
}  # ferme if ($selectedUPNSuffix -ne $null)
Write-Host "Script terminé." -ForegroundColor Yellow


#####################################
#  .-. .-')       (`-.   _  .-')    #
#  \  ( OO )    _(OO  )_( \( -O )   #
#  ,--. ,--.,--(_/   ,. \,------.   #
#  |  .'   /\   \   /(__/|   /`. '  #
#  |      /, \   \ /   / |  /  | |  #
#  |     ' _) \   '   /, |  |_.' |  #
#  |  .   \    \     /__)|  .  '.'  #
#  |  |\   \    \   /    |  |\  \   #
#  `--' '--'     `-'     `--' '--'  #
#####################################
