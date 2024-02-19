# 1) Faire son export CSV depuis M365 en ayant filtré le type d'user voulut
# 2) Avoir un csv bien délimité: ouvrir excel -> ouvrir csv -> Délimité -> Cocher seulement Virgule -> Standart -> Terminer
# 3) Penser à modifier @... ,Ligne 120, de votre UserPrincipalName par le votre, actuellement: "$samAccountName@groupe-test.com"
# 4) Penser à modifier @... ,Ligne 125, de votre EmailAddress par le votre, actuellement: "$samAccountName@groupe-test.com"
# 5) Penser à mettre un MDP en accord avec vos GPO

# Charger le module Active Directory si pas déjà chargé
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}

function Select-CSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.FileName
}

function Select-ADOUAndPassword {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Création d''utilisateurs à partir de csv M365'
    $form.Size = New-Object System.Drawing.Size(520,420)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = 'LightGray'

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(480,20)
    $label.Text = 'Entrer le nom de L''OU a rechercher:'
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,35)
    $textBox.Size = New-Object System.Drawing.Size(480,20)
    $form.Controls.Add($textBox)

    $ouLabel = New-Object System.Windows.Forms.Label
    $ouLabel.Location = New-Object System.Drawing.Point(10,60)
    $ouLabel.Size = New-Object System.Drawing.Size(480,20)
    $ouLabel.Text = 'Sélectionner l''OU de destination:'
    $form.Controls.Add($ouLabel)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,85)
    $listBox.Size = New-Object System.Drawing.Size(480,200)
    $listBox.BackColor = 'White'
    $form.Controls.Add($listBox)

    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Location = New-Object System.Drawing.Point(10,290)
    $passwordLabel.Size = New-Object System.Drawing.Size(480,20)
    $passwordLabel.Text = 'Entrer le mot de passe à appliquer:'
    $form.Controls.Add($passwordLabel)

    $passwordBox = New-Object System.Windows.Forms.TextBox
    $passwordBox.Location = New-Object System.Drawing.Point(10,315)
    $passwordBox.Size = New-Object System.Drawing.Size(480,20)
    #$passwordBox.PasswordChar = '*'
    $passwordBox.BackColor = 'White'
    $form.Controls.Add($passwordBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(390,345)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = 'OK'
    $okButton.BackColor = 'RoyalBlue'
    $okButton.ForeColor = 'White'
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
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
    }
}


$csvPath = Select-CSVFile
if ([string]::IsNullOrWhiteSpace($csvPath)) {
    Write-Host "Aucun CSV sélectionné, arrêt script."
    exit
}

$users = Import-Csv -Path $csvPath -Delimiter ';'

$selectedOU, $password = Select-ADOUAndPassword
if ($selectedOU -eq $null -or $password -eq "") {
    Write-Host "Aucune OU sélectionnée/mdp donné, arrêt script."
    exit
}

foreach ($user in $users) {
    # Ajout prénom en minuscule, nom en maj et build SamAccountName en min, le tout en enlevant les espaces
    $formattedGivenName = $user.Prénom.Replace(" ", "").ToLower()
    $formattedSurname = $user.Nom.Replace(" ", "").ToUpper()
    $samAccountName = ("{0}.{1}" -f $formattedGivenName, $user.Nom.Replace(" ", "").ToLower())

    # Creation paramètres de base pour chaque user
    $ADUserParams = @{
        Enabled               = $true
        Path                  = $selectedOU
        AccountPassword       = (ConvertTo-SecureString -AsPlainText $password -Force)
        PasswordNeverExpires  = $false
        ChangePasswordAtLogon = $true
        UserPrincipalName     = "$samAccountName@groupe-test.com"
        SamAccountName        = $samAccountName
        Name                  = "$formattedGivenName $formattedSurname"
        GivenName             = $formattedGivenName
        Surname               = $formattedSurname
        DisplayName           = "$formattedGivenName $formattedSurname"
        EmailAddress          = "$samAccountName@groupe-test.com"
        City                  = $user.Ville
        PostalCode            = $user.'Code postal'
        State                 = $user.État
        Title                 = $user.Titre
        Department            = $user.Service
    }

    # Tentative création de chaque user avec paramètres accumulés
    try {
        New-ADUser @ADUserParams -ErrorAction Stop
        Write-Host "Utilisateur créé : $($ADUserParams['GivenName']) $($ADUserParams['Surname'])"
    } catch {
        Write-Host "Erreur lors de la création de l'utilisateur : $($ADUserParams['GivenName']) $($ADUserParams['Surname'])"
        Write-Host "Détail de l'erreur : $($_.Exception.Message)"
    }
}

# Force une synchro avec M365
Start-ADSyncSyncCycle

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
