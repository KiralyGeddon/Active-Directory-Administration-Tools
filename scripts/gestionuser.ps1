#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Onglet Gestion Utilisateurs ###
#
# Description :
# Ce script definit l'interface et la logique de l'onglet "Gestion des Utilisateurs".
# Il est charge par main.ps1.
#-----------------------------------------------------------------------------------

# Assurez-vous que les variables globales sont accessibles
# $script:tabControl, $script:CurrentLanguage, $script:Localization, Get-LocalizedText
# $script:domainPath, $script:domainSuffix, Write-Log, Populate-ADDropdowns

#region Fonctions Logiques de l'onglet Gestion Utilisateurs

function Add-SingleUser {
    $prenom = $script:addUserPrenomTextBox.Text
    $nom = $script:addUserNomTextBox.Text
    $fonction = $script:addUserFonctionTextBox.Text

    if (-not $prenom -or -not $nom -or -not $fonction) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorAddUserFields"), "Erreur", "OK", "Error")
        return
    }

    $login = "$($prenom.Substring(0,1).ToLower()).$($nom.ToLower())"
    $email = "$login@$script:domainSuffix"
    $password = "Azerty123456!"
    $ouFonctionPath = "OU=$fonction,OU=Utilisateurs,$script:domainPath"

    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouFonctionPath'" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $fonction -Path "OU=Utilisateurs,$script:domainPath"
        Write-Log (Get-LocalizedText "LogFunctionOuCreated" $fonction)
    }

    if (Get-ADUser -Filter { SamAccountName -eq $login }) {
        Write-Log (Get-LocalizedText "LogUserExists" $login) "AVERTISSEMENT"
    } else {
        try {
            New-ADUser -Name "$nom $prenom" `
                       -GivenName $prenom `
                       -Surname $nom `
                       -SamAccountName $login `
                       -UserPrincipalName $email `
                       -Title $fonction `
                       -Path $ouFonctionPath `
                       -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                       -ChangePasswordAtLogon $true `
                       -Enabled $true -ErrorAction Stop
            Write-Log (Get-LocalizedText "LogUserCreatedManual" $login)
            $script:addUserPrenomTextBox.Clear()
            $script:addUserNomTextBox.Clear()
            $script:addUserFonctionTextBox.Clear()
            Populate-ADDropdowns
        } catch {
            Write-Log (Get-LocalizedText "LogErrorUserCreationManual" $login $_) "ERREUR"
        }
    }
}

function Remove-SpecificUser {
    $login = $script:deleteUserLoginComboBox.SelectedItem
    if (-not $login) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectUserToDelete"), "Erreur", "OK", "Error")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgConfirmDeleteUser" $login), (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        try {
            $user = Get-ADUser -Identity $login -ErrorAction Stop
            Remove-ADUser -Identity $user -Confirm:$false -ErrorAction Stop
            Write-Log (Get-LocalizedText "LogUserDeleted" $login)
            $script:deleteUserLoginComboBox.SelectedIndex = -1
            Populate-ADDropdowns
        } catch {
            Write-Log (Get-LocalizedText "LogErrorUserDeletion" $login $_) "ERREUR"
        }
    }
}

#endregion

#region Définition de l'interface graphique de l'onglet Gestion Utilisateurs

$script:tabPageUsers = New-Object System.Windows.Forms.TabPage
$script:tabPageUsers.Text = (Get-LocalizedText "UsersTabTitle")
$script:tabControl.TabPages.Add($script:tabPageUsers)

$script:addUserGroupBox = New-Object System.Windows.Forms.GroupBox
$script:addUserGroupBox.Text = (Get-LocalizedText "AddUserGroupBox")
$script:addUserGroupBox.Location = New-Object System.Drawing.Point(20, 20)
$script:addUserGroupBox.Size = New-Object System.Drawing.Size(270, 280)
$script:tabPageUsers.Controls.Add($script:addUserGroupBox)

$script:addUserPrenomLabel = New-Object System.Windows.Forms.Label; $script:addUserPrenomLabel.Text = (Get-LocalizedText "AddUserFirstNameLabel"); $script:addUserPrenomLabel.Location = New-Object System.Drawing.Point(15, 40); $script:addUserPrenomLabel.AutoSize = $true; $script:addUserGroupBox.Controls.Add($script:addUserPrenomLabel)
$script:addUserPrenomTextBox = New-Object System.Windows.Forms.TextBox; $script:addUserPrenomTextBox.Location = New-Object System.Drawing.Point(100, 37); $script:addUserPrenomTextBox.Size = New-Object System.Drawing.Size(150, 20); $script:addUserGroupBox.Controls.Add($script:addUserPrenomTextBox)

$script:addUserNomLabel = New-Object System.Windows.Forms.Label; $script:addUserNomLabel.Text = (Get-LocalizedText "AddUserLastNameLabel"); $script:addUserNomLabel.Location = New-Object System.Drawing.Point(15, 80); $script:addUserNomLabel.AutoSize = $true; $script:addUserGroupBox.Controls.Add($script:addUserNomLabel)
$script:addUserNomTextBox = New-Object System.Windows.Forms.TextBox; $script:addUserNomTextBox.Location = New-Object System.Drawing.Point(100, 77); $script:addUserNomTextBox.Size = New-Object System.Drawing.Size(150, 20); $script:addUserGroupBox.Controls.Add($script:addUserNomTextBox)

$script:addUserFonctionLabel = New-Object System.Windows.Forms.Label; $script:addUserFonctionLabel.Text = (Get-LocalizedText "AddUserFunctionLabel"); $script:addUserFonctionLabel.Location = New-Object System.Drawing.Point(15, 120); $script:addUserFonctionLabel.AutoSize = $true; $script:addUserGroupBox.Controls.Add($script:addUserFonctionLabel)
$script:addUserFonctionTextBox = New-Object System.Windows.Forms.TextBox; $script:addUserFonctionTextBox.Location = New-Object System.Drawing.Point(100, 117); $script:addUserFonctionTextBox.Size = New-Object System.Drawing.Size(150, 20); $script:addUserGroupBox.Controls.Add($script:addUserFonctionTextBox)

$script:addUserButton = New-Object System.Windows.Forms.Button; $script:addUserButton.Text = (Get-LocalizedText "AddUserButton"); $script:addUserButton.Location = New-Object System.Drawing.Point(60, 180); $script:addUserButton.Size = New-Object System.Drawing.Size(150, 30); $script:addUserButton.Add_Click({ Add-SingleUser }); $script:addUserGroupBox.Controls.Add($script:addUserButton)

$script:deleteUserGroupBox = New-Object System.Windows.Forms.GroupBox
$script:deleteUserGroupBox.Text = (Get-LocalizedText "DeleteUserGroupBox")
$script:deleteUserGroupBox.Location = New-Object System.Drawing.Point(310, 20)
$script:deleteUserGroupBox.Size = New-Object System.Drawing.Size(285, 280)
$script:tabPageUsers.Controls.Add($script:deleteUserGroupBox)

$script:deleteUserLoginLabel = New-Object System.Windows.Forms.Label; $script:deleteUserLoginLabel.Text = (Get-LocalizedText "DeleteUserSelectUserLabel"); $script:deleteUserLoginLabel.Location = New-Object System.Drawing.Point(15, 40); $script:deleteUserLoginLabel.AutoSize = $true; $script:deleteUserGroupBox.Controls.Add($script:deleteUserLoginLabel)

$script:deleteUserLoginComboBox = New-Object System.Windows.Forms.ComboBox
$script:deleteUserLoginComboBox.Location = New-Object System.Drawing.Point(15, 65)
$script:deleteUserLoginComboBox.Size = New-Object System.Drawing.Size(250, 20)
$script:deleteUserLoginComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:deleteUserGroupBox.Controls.Add($script:deleteUserLoginComboBox)

$script:deleteUserButton = New-Object System.Windows.Forms.Button; $script:deleteUserButton.Text = (Get-LocalizedText "DeleteUserButton"); $script:deleteUserButton.Location = New-Object System.Drawing.Point(65, 180); $script:deleteUserButton.Size = New-Object System.Drawing.Size(150, 30); $script:deleteUserButton.BackColor = "MistyRose"; $script:deleteUserButton.Add_Click({ Remove-SpecificUser }); $script:deleteUserGroupBox.Controls.Add($script:deleteUserButton)

#endregion