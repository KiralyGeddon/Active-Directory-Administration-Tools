#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Onglet Gestion des OU ###
#
# Description :
# Ce script definit l'interface et la logique de l'onglet "Gestion des Unites d'Organisation (OU)".
# Il est charge par main.ps1.
#-----------------------------------------------------------------------------------

# Assurez-vous que les variables globales sont accessibles
# $script:tabControl, $script:CurrentLanguage, $script:Localization, Get-LocalizedText
# $script:domainPath, Write-Log, Populate-ADDropdowns

#region Fonctions Logiques de l'onglet Gestion des OU

function Add-OU {
    $ouName = $script:createOuNameTextBox.Text
    $parentOuPathInput = $script:createOuParentPathTextBox.Text # User input for parent path

    if (-not $ouName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorOuNameRequired"), "Erreur", "OK", "Error")
        return
    }

    # Construct the full parent path
    $targetPath = $script:domainPath
    if ($parentOuPathInput) {
        # Split by comma to handle multiple levels like "RH,OU=Services"
        $pathParts = $parentOuPathInput -split ',' | ForEach-Object { "$($_.Trim())" }
        # Reconstruct in AD format (e.g., "OU=RH,OU=Services")
        $formattedParentPath = ($pathParts | ForEach-Object {
            if ($_ -notmatch "^OU=") { "OU=$_" } else { $_ }
        }) -join ','
        $targetPath = "$formattedParentPath,$script:domainPath"
    }

    try {
        # Check if parent OU path exists if specified
        if ($parentOuPathInput -and -not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$targetPath'" -ErrorAction SilentlyContinue)) {
            Write-Log (Get-LocalizedText "LogErrorOuCreation" $ouName "Le chemin parent specifie n'existe pas ou est invalide : $parentOuPathInput.") "ERREUR"
            [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "LogErrorOuCreation" $ouName "Le chemin parent specifie n'existe pas ou est invalide : $parentOuPathInput."), "Erreur", "OK", "Error")
            return
        }

        # Check if the OU already exists at the target path
        if (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $targetPath -ErrorAction SilentlyContinue) {
             Write-Log (Get-LocalizedText "LogOuExists" $ouName $targetPath) "AVERTISSEMENT"
             [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "LogOuExists" $ouName $targetPath), "Information", "OK", "Information")
             return
        }

        # Create the OU
        New-ADOrganizationalUnit -Name $ouName -Path $targetPath -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogOuCreated" $ouName $targetPath)
        $script:createOuNameTextBox.Clear()
        $script:createOuParentPathTextBox.Clear()
        Populate-ADDropdowns # Refresh all dropdowns including OUs
    } catch {
        Write-Log (Get-LocalizedText "LogErrorOuCreation" $ouName $_.Exception.Message) "ERREUR"
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "LogErrorOuCreation" $ouName $_.Exception.Message), "Erreur", "OK", "Error")
    }
}

function Remove-OU {
    $ouName = $script:deleteOuNameComboBox.SelectedItem
    
    if (-not $ouName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectOuToDelete"), "Erreur", "OK", "Error")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgConfirmDeleteOu" $ouName), (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        try {
            # Find the OU by name. If multiple OUs with same name exist, this will take the first one found or error if filter is too broad.
            # A more robust solution might require selecting by DN or providing parent path.
            $ouToDelete = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction Stop
            Remove-ADOrganizationalUnit -Identity $ouToDelete -Confirm:$false -ErrorAction Stop
            Write-Log (Get-LocalizedText "LogOuDeleted" $ouName)
            $script:deleteOuNameComboBox.SelectedIndex = -1
            Populate-ADDropdowns # Refresh OU list
        } catch {
            Write-Log (Get-LocalizedText "LogErrorOuDeletion" $ouName $_.Exception.Message) "ERREUR"
            [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "LogErrorOuDeletion" $ouName $_.Exception.Message), "Erreur", "OK", "Error")
        }
    }
}

#endregion

#region Définition de l'interface graphique de l'onglet Gestion des OU

$script:tabPageOU = New-Object System.Windows.Forms.TabPage
$script:tabPageOU.Text = (Get-LocalizedText "OuTabTitle")
$script:tabControl.TabPages.Add($script:tabPageOU)

# GroupBox pour la création d'OU
$script:createOuGroupBox = New-Object System.Windows.Forms.GroupBox
$script:createOuGroupBox.Text = (Get-LocalizedText "CreateOuGroupBox")
$script:createOuGroupBox.Location = New-Object System.Drawing.Point(20, 20)
$script:createOuGroupBox.Size = New-Object System.Drawing.Size(280, 190)
$script:tabPageOU.Controls.Add($script:createOuGroupBox)

$script:createOuNameLabel = New-Object System.Windows.Forms.Label
$script:createOuNameLabel.Text = (Get-LocalizedText "CreateOuNameLabel")
$script:createOuNameLabel.Location = New-Object System.Drawing.Point(15, 30)
$script:createOuNameLabel.AutoSize = $true
$script:createOuGroupBox.Controls.Add($script:createOuNameLabel)

$script:createOuNameTextBox = New-Object System.Windows.Forms.TextBox
$script:createOuNameTextBox.Location = New-Object System.Drawing.Point(15, 55)
$script:createOuNameTextBox.Size = New-Object System.Drawing.Size(250, 20)
$script:createOuGroupBox.Controls.Add($script:createOuNameTextBox)

$script:createOuParentPathLabel = New-Object System.Windows.Forms.Label
$script:createOuParentPathLabel.Text = (Get-LocalizedText "CreateOuParentPathLabel")
$script:createOuParentPathLabel.Location = New-Object System.Drawing.Point(15, 85)
$script:createOuParentPathLabel.AutoSize = $true
$script:createOuGroupBox.Controls.Add($script:createOuParentPathLabel)

$script:createOuParentPathTextBox = New-Object System.Windows.Forms.TextBox
$script:createOuParentPathTextBox.Location = New-Object System.Drawing.Point(15, 110)
$script:createOuParentPathTextBox.Size = New-Object System.Drawing.Size(250, 20)
$script:createOuGroupBox.Controls.Add($script:createOuParentPathTextBox)

$script:createOuButton = New-Object System.Windows.Forms.Button
$script:createOuButton.Text = (Get-LocalizedText "CreateOuButton")
$script:createOuButton.Location = New-Object System.Drawing.Point(90, 140)
$script:createOuButton.Size = New-Object System.Drawing.Size(100, 25)
$script:createOuButton.Add_Click({ Add-OU })
$script:createOuGroupBox.Controls.Add($script:createOuButton)

# GroupBox pour la suppression d'OU
$script:deleteOuGroupBox = New-Object System.Windows.Forms.GroupBox
$script:deleteOuGroupBox.Text = (Get-LocalizedText "DeleteOuGroupBox")
$script:deleteOuGroupBox.Location = New-Object System.Drawing.Point(315, 20)
$script:deleteOuGroupBox.Size = New-Object System.Drawing.Size(280, 190)
$script:tabPageOU.Controls.Add($script:deleteOuGroupBox)

$script:deleteOuNameLabel = New-Object System.Windows.Forms.Label
$script:deleteOuNameLabel.Text = (Get-LocalizedText "DeleteOuNameLabel")
$script:deleteOuNameLabel.Location = New-Object System.Drawing.Point(15, 30)
$script:deleteOuNameLabel.AutoSize = $true
$script:deleteOuGroupBox.Controls.Add($script:deleteOuNameLabel)

$script:deleteOuNameComboBox = New-Object System.Windows.Forms.ComboBox
$script:deleteOuNameComboBox.Location = New-Object System.Drawing.Point(15, 55)
$script:deleteOuNameComboBox.Size = New-Object System.Drawing.Size(250, 20)
$script:deleteOuNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:deleteOuGroupBox.Controls.Add($script:deleteOuNameComboBox)

$script:deleteOuButton = New-Object System.Windows.Forms.Button
$script:deleteOuButton.Text = (Get-LocalizedText "DeleteOuButton")
$script:deleteOuButton.Location = New-Object System.Drawing.Point(90, 140)
$script:deleteOuButton.Size = New-Object System.Drawing.Size(100, 25)
$script:deleteOuButton.BackColor = "MistyRose"
$script:deleteOuButton.Add_Click({ Remove-OU })
$script:deleteOuGroupBox.Controls.Add($script:deleteOuButton)

#endregion