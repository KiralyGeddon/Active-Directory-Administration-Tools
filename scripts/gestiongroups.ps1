#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Onglet Gestion Groupes ###
#
# Description :
# Ce script definit l'interface et la logique de l'onglet "Gestion des Groupes".
# Il est charge par main.ps1.
#-----------------------------------------------------------------------------------

# Assurez-vous que les variables globales sont accessibles
# $script:tabControl, $script:CurrentLanguage, $script:Localization, Get-LocalizedText
# $script:domainPath, $script:domainSuffix, Write-Log, Populate-ADDropdowns

#region Fonctions Logiques de l'onglet Gestion Groupes

function Add-SpecificGroup {
    $groupName = $script:createGroupNameTextBox.Text
    if (-not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorCreateGroupField"), "Erreur", "OK", "Error")
        return
    }

    try {
        if(Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue) {
             Write-Log (Get-LocalizedText "LogGroupExists" $groupName) "AVERTISSEMENT"
             return
        }
        New-ADGroup -Name $groupName -GroupScope Global -Path $script:domainPath -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogGroupCreated" $groupName)
        $script:createGroupNameTextBox.Clear()
        Populate-ADDropdowns
    } catch {
        Write-Log (Get-LocalizedText "LogErrorGroupCreation" $groupName $_) "ERREUR"
    }
}

function Remove-SpecificGroup {
    $groupName = $script:deleteGroupNameComboBox.SelectedItem
    if (-not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectGroupToDelete"), "Erreur", "OK", "Error")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgConfirmDeleteGroup" $groupName), (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        try {
            Get-ADGroup -Identity $groupName -ErrorAction Stop | Remove-ADGroup -Confirm:$false -ErrorAction Stop
            Write-Log (Get-LocalizedText "LogGroupDeleted" $groupName)
            $script:deleteGroupNameComboBox.SelectedIndex = -1
            Populate-ADDropdowns
        } catch {
            Write-Log (Get-LocalizedText "LogErrorGroupDeletion" $groupName $_) "ERREUR"
        }
    }
}

function Add-UserToSpecificGroup {
    $login = $script:addUserToGroupLoginComboBox.SelectedItem
    $groupName = $script:addUserToGroupGroupNameComboBox.SelectedItem

    if (-not $login -or -not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorAddUserToGroupFields"), "Erreur", "OK", "Error")
        return
    }

    try {
        Add-ADGroupMember -Identity $groupName -Members $login -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogUserAddedToGroup" $login $groupName)
        $script:addUserToGroupLoginComboBox.SelectedIndex = -1
        $script:addUserToGroupGroupNameComboBox.SelectedIndex = -1
        Populate-ADDropdowns
    } catch {
        Write-Log (Get-LocalizedText "LogErrorAddUserToGroup" $login $groupName $_) "ERREUR"
    }
}

#endregion

#region Définition de l'interface graphique de l'onglet Gestion Groupes

$script:tabPageGroups = New-Object System.Windows.Forms.TabPage
$script:tabPageGroups.Text = (Get-LocalizedText "GroupsTabTitle")
$script:tabControl.TabPages.Add($script:tabPageGroups)

$script:createGroupGroupBox = New-Object System.Windows.Forms.GroupBox; $script:createGroupGroupBox.Text = (Get-LocalizedText "CreateGroupGroupBox"); $script:createGroupGroupBox.Location = New-Object System.Drawing.Point(20, 20); $script:createGroupGroupBox.Size = New-Object System.Drawing.Size(280, 130); $script:tabPageGroups.Controls.Add($script:createGroupGroupBox)
$script:createGroupNameLabel = New-Object System.Windows.Forms.Label; $script:createGroupNameLabel.Text = (Get-LocalizedText "CreateGroupNameLabel"); $script:createGroupNameLabel.Location = New-Object System.Drawing.Point(15, 30); $script:createGroupNameLabel.AutoSize = $true; $script:createGroupGroupBox.Controls.Add($script:createGroupNameLabel)
$script:createGroupNameTextBox = New-Object System.Windows.Forms.TextBox; $script:createGroupNameTextBox.Location = New-Object System.Drawing.Point(15, 55); $script:createGroupNameTextBox.Size = New-Object System.Drawing.Size(250, 20); $script:createGroupGroupBox.Controls.Add($script:createGroupNameTextBox)
$script:createGroupButton = New-Object System.Windows.Forms.Button; $script:createGroupButton.Text = (Get-LocalizedText "CreateGroupButton"); $script:createGroupButton.Location = New-Object System.Drawing.Point(90, 85); $script:createGroupButton.Size = New-Object System.Drawing.Size(100, 25); $script:createGroupButton.Add_Click({ Add-SpecificGroup }); $script:createGroupGroupBox.Controls.Add($script:createGroupButton)

$script:deleteGroupGroupBox = New-Object System.Windows.Forms.GroupBox; $script:deleteGroupGroupBox.Text = (Get-LocalizedText "DeleteGroupGroupBox"); $script:deleteGroupGroupBox.Location = New-Object System.Drawing.Point(315, 20); $script:deleteGroupGroupBox.Size = New-Object System.Drawing.Size(280, 130); $script:tabPageGroups.Controls.Add($script:deleteGroupGroupBox)
$script:deleteGroupNameLabel = New-Object System.Windows.Forms.Label; $script:deleteGroupNameLabel.Text = (Get-LocalizedText "DeleteGroupSelectGroupLabel"); $script:deleteGroupNameLabel.Location = New-Object System.Drawing.Point(15, 30); $script:deleteGroupNameLabel.AutoSize = $true; $script:deleteGroupGroupBox.Controls.Add($script:deleteGroupNameLabel)

$script:deleteGroupNameComboBox = New-Object System.Windows.Forms.ComboBox
$script:deleteGroupNameComboBox.Location = New-Object System.Drawing.Point(15, 55)
$script:deleteGroupNameComboBox.Size = New-Object System.Drawing.Size(250, 20)
$script:deleteGroupNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:deleteGroupGroupBox.Controls.Add($script:deleteGroupNameComboBox)

$script:deleteGroupButton = New-Object System.Windows.Forms.Button; $script:deleteGroupButton.Text = (Get-LocalizedText "DeleteGroupButton"); $script:deleteGroupButton.Location = New-Object System.Drawing.Point(90, 85); $script:deleteGroupButton.Size = New-Object System.Drawing.Size(100, 25); $script:deleteGroupButton.BackColor = "MistyRose"; $script:deleteGroupButton.Add_Click({ Remove-SpecificGroup }); $script:deleteGroupGroupBox.Controls.Add($script:deleteGroupButton)

$script:addUserToGroupBox = New-Object System.Windows.Forms.GroupBox; $script:addUserToGroupBox.Text = (Get-LocalizedText "AddUserToGroupGroupBox"); $script:addUserToGroupBox.Location = New-Object System.Drawing.Point(20, 170); $script:addUserToGroupBox.Size = New-Object System.Drawing.Size(575, 140); $script:tabPageGroups.Controls.Add($script:addUserToGroupBox)
$script:addUserToGroupLoginLabel = New-Object System.Windows.Forms.Label; $script:addUserToGroupLoginLabel.Text = (Get-LocalizedText "AddUserToGroupSelectUserLabel"); $script:addUserToGroupLoginLabel.Location = New-Object System.Drawing.Point(15, 30); $script:addUserToGroupLoginLabel.AutoSize = $true; $script:addUserToGroupBox.Controls.Add($script:addUserToGroupLoginLabel)

$script:addUserToGroupLoginComboBox = New-Object System.Windows.Forms.ComboBox
$script:addUserToGroupLoginComboBox.Location = New-Object System.Drawing.Point(15, 55)
$script:addUserToGroupLoginComboBox.Size = New-Object System.Drawing.Size(250, 20)
$script:addUserToGroupLoginComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:addUserToGroupBox.Controls.Add($script:addUserToGroupLoginComboBox)

$script:addUserToGroupGroupNameLabel = New-Object System.Windows.Forms.Label; $script:addUserToGroupGroupNameLabel.Text = (Get-LocalizedText "AddUserToGroupSelectGroupLabel"); $script:addUserToGroupGroupNameLabel.Location = New-Object System.Drawing.Point(300, 30); $script:addUserToGroupGroupNameLabel.AutoSize = $true; $script:addUserToGroupBox.Controls.Add($script:addUserToGroupGroupNameLabel)

$script:addUserToGroupGroupNameComboBox = New-Object System.Windows.Forms.ComboBox
$script:addUserToGroupGroupNameComboBox.Location = New-Object System.Drawing.Point(300, 55)
$script:addUserToGroupGroupNameComboBox.Size = New-Object System.Drawing.Size(250, 20)
$script:addUserToGroupGroupNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:addUserToGroupBox.Controls.Add($script:addUserToGroupGroupNameComboBox)

$script:addUserToGroupButton = New-Object System.Windows.Forms.Button; $script:addUserToGroupButton.Text = (Get-LocalizedText "AddUserToGroupButton"); $script:addUserToGroupButton.Location = New-Object System.Drawing.Point(212, 95); $script:addUserToGroupButton.Size = New-Object System.Drawing.Size(150, 30); $script:addUserToGroupButton.Add_Click({ Add-UserToSpecificGroup }); $script:addUserToGroupBox.Controls.Add($script:addUserToGroupButton)

#endregion