#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool ###
#
# Version : 3.0 (Internationalise) - Refactorise
# Auteur : Kiraly Geddon aide par Gemini
#
# Description :
# Ce script fournit une interface graphique (GUI) pour effectuer des operations
# courantes dans Active Directory : import en masse, ajout/suppression d'utilisateurs
# et gestion de groupes. Refactorise en modules.
#-----------------------------------------------------------------------------------

#region 0. CONFIGURATION ET CHARGEMENT DES LIBRAIRIES

# Langue par defaut
$script:CurrentLanguage = "fr" # ou "en"

# Charger les fonctions utilitaires et la configuration de localisation
. "$PSScriptRoot\librairies\lib.ps1"

#endregion

#region 1. VeRIFICATION DES PReREQUIS (ADMIN & RSAT)

# --- A. Verification des droits d'Administrateur ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning (Get-LocalizedText "PrereqAdminNeeded")
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -File `"$PSCommandPath`""
    exit
}

# --- B. Verification et installation du module Active Directory (RSAT) ---
Add-Type -AssemblyName System.Windows.Forms

if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    $installChoice = [System.Windows.Forms.MessageBox]::Show(
        (Get-LocalizedText "PrereqModuleNotFound"),
        (Get-LocalizedText "PrereqMissingTitle"),
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($installChoice -eq 'Yes') {
        try {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallStarting"),
                "Information", "OK", "Information"
            )
            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallSuccess"),
                (Get-LocalizedText "PrereqInstallFinishedTitle"),
                "OK", "Information"
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallFailed" $_.Exception.Message),
                (Get-LocalizedText "PrereqInstallErrorTitle"),
                "OK", "Error"
            )
        }
        finally {
            exit
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "PrereqOperationCancelled"), (Get-LocalizedText "OperationCancelledTitle"), "OK", "Warning")
        exit
    }
}

#endregion

#region 2. DeTECTION DU DOMAINE ET CHARGEMENT DES MODULES

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $script:adDomain = Get-ADDomain
    $script:domainPath = $script:adDomain.DistinguishedName
    $script:domainSuffix = $script:adDomain.DNSRoot
}
catch {
    [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "ErrorCriticalAd"), "Erreur Critique", "OK", "Error")
    exit
}

Add-Type -AssemblyName System.Drawing
#endregion


#region Chargement des Assemblages .NET pour l'Interface Graphique (GUI)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion


#region Definition de l'Interface Graphique (GUI) Principale

$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = (Get-LocalizedText "MainWindowTitle" -Args $script:domainSuffix)
$script:mainForm.Size = New-Object System.Drawing.Size(650, 600)
$script:mainForm.StartPosition = 'CenterScreen'
$script:mainForm.FormBorderStyle = 'FixedDialog'
$script:mainForm.MaximizeBox = $false

$script:tabControl = New-Object System.Windows.Forms.TabControl
$script:tabControl.Location = New-Object System.Drawing.Point(10, 10)
$script:tabControl.Size = New-Object System.Drawing.Size(615, 350)
$script:mainForm.Controls.Add($script:tabControl)

# Charger les modules d'onglets
. "$PSScriptRoot\scripts\importcsv.ps1"
. "$PSScriptRoot\scripts\gestionuser.ps1"
. "$PSScriptRoot\scripts\gestiongoups.ps1"
. "$PSScriptRoot\scripts\informations.ps1"
. "$PSScriptRoot\scripts\gestionou.ps1" # Ajout du script de gestion des OU

# Zone de logs
$script:outputLabel = New-Object System.Windows.Forms.Label
$script:outputLabel.Text = (Get-LocalizedText "LogsLabel")
$script:outputLabel.Location = New-Object System.Drawing.Point(10, 370)
$script:outputLabel.AutoSize = $true
$script:mainForm.Controls.Add($script:outputLabel)

$script:outputTextBox = New-Object System.Windows.Forms.TextBox
$script:outputTextBox.Location = New-Object System.Drawing.Point(10, 390)
$script:outputTextBox.Size = New-Object System.Drawing.Size(615, 130)
$script:outputTextBox.Multiline = $true
$script:outputTextBox.ScrollBars = "Vertical"
$script:outputTextBox.ReadOnly = $true
$script:outputTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
$script:mainForm.Controls.Add($script:outputTextBox)

# Contrôle de selection de la langue
$script:languageLabel = New-Object System.Windows.Forms.Label
$script:languageLabel.Text = (Get-LocalizedText "LanguageLabel")
$script:languageLabel.Location = New-Object System.Drawing.Point(10, 530)
$script:languageLabel.AutoSize = $true
$script:mainForm.Controls.Add($script:languageLabel)

$script:languageComboBox = New-Object System.Windows.Forms.ComboBox
$script:languageComboBox.Location = New-Object System.Drawing.Point(80, 527)
$script:languageComboBox.Size = New-Object System.Drawing.Size(100, 20)
$script:languageComboBox.Items.AddRange(@("fr", "en")) # Ajouter les langues disponibles
$script:languageComboBox.SelectedItem = $script:CurrentLanguage # Selectionner la langue par defaut
$script:languageComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList # Empêche l'edition
$script:languageComboBox.Add_SelectedIndexChanged({
    $script:CurrentLanguage = $script:languageComboBox.SelectedItem # Mettre à jour la variable de langue globale
    Update-GuiLanguage # Appeler la fonction pour rafraîchir les textes
})
$script:mainForm.Controls.Add($script:languageComboBox)


# Fonction pour mettre à jour les textes de la GUI en fonction de la langue selectionnee
function Update-GuiLanguage {
    $script:mainForm.Text = (Get-LocalizedText "MainWindowTitle" -Args $script:domainSuffix)

    # Onglet Import CSV
    if ($script:tabPageImport) { $script:tabPageImport.Text = (Get-LocalizedText "ImportTabTitle") }
    if ($script:importLabel) { $script:importLabel.Text = (Get-LocalizedText "ImportLabelPath") }
    if ($script:browseButton) { $script:browseButton.Text = (Get-LocalizedText "BrowseButton") }
    if ($script:importButton) { $script:importButton.Text = (Get-LocalizedText "ImportButton") }

    # Onglet Gestion Utilisateurs
    if ($script:tabPageUsers) { $script:tabPageUsers.Text = (Get-LocalizedText "UsersTabTitle") }
    if ($script:addUserGroupBox) { $script:addUserGroupBox.Text = (Get-LocalizedText "AddUserGroupBox") }
    if ($script:addUserPrenomLabel) { $script:addUserPrenomLabel.Text = (Get-LocalizedText "AddUserFirstNameLabel") }
    if ($script:addUserNomLabel) { $script:addUserNomLabel.Text = (Get-LocalizedText "AddUserLastNameLabel") }
    if ($script:addUserFonctionLabel) { $script:addUserFonctionLabel.Text = (Get-LocalizedText "AddUserFunctionLabel") }
    if ($script:addUserButton) { $script:addUserButton.Text = (Get-LocalizedText "AddUserButton") }
    if ($script:deleteUserGroupBox) { $script:deleteUserGroupBox.Text = (Get-LocalizedText "DeleteUserGroupBox") }
    if ($script:deleteUserLoginLabel) { $script:deleteUserLoginLabel.Text = (Get-LocalizedText "DeleteUserSelectUserLabel") }
    if ($script:deleteUserButton) { $script:deleteUserButton.Text = (Get-LocalizedText "DeleteUserButton") }

    # Onglet Gestion Groupes
    if ($script:tabPageGroups) { $script:tabPageGroups.Text = (Get-LocalizedText "GroupsTabTitle") }
    if ($script:createGroupGroupBox) { $script:createGroupGroupBox.Text = (Get-LocalizedText "CreateGroupGroupBox") }
    if ($script:createGroupNameLabel) { $script:createGroupNameLabel.Text = (Get-LocalizedText "CreateGroupNameLabel") }
    if ($script:createGroupButton) { $script:createGroupButton.Text = (Get-LocalizedText "CreateGroupButton") }
    if ($script:deleteGroupGroupBox) { $script:deleteGroupGroupBox.Text = (Get-LocalizedText "DeleteGroupGroupBox") }
    if ($script:deleteGroupNameLabel) { $script:deleteGroupNameLabel.Text = (Get-LocalizedText "DeleteGroupSelectGroupLabel") }
    if ($script:deleteGroupButton) { $script:deleteGroupButton.Text = (Get-LocalizedText "DeleteGroupButton") }
    if ($script:addUserToGroupBox) { $script:addUserToGroupBox.Text = (Get-LocalizedText "AddUserToGroupGroupBox") }
    if ($script:addUserToGroupLoginLabel) { $script:addUserToGroupLoginLabel.Text = (Get-LocalizedText "AddUserToGroupSelectUserLabel") }
    if ($script:addUserToGroupGroupNameLabel) { $script:addUserToGroupGroupNameLabel.Text = (Get-LocalizedText "AddUserToGroupSelectGroupLabel") }
    if ($script:addUserToGroupButton) { $script:addUserToGroupButton.Text = (Get-LocalizedText "AddUserToGroupButton") }

    # Onglet Informations
    if ($script:tabPageInfo) { $script:tabPageInfo.Text = (Get-LocalizedText "InfoTabTitle") }
    if ($script:infoLabelTitle) { $script:infoLabelTitle.Text = (Get-LocalizedText "InfoLabelTitle") }
    if ($script:infoLabelVersion) { $script:infoLabelVersion.Text = (Get-LocalizedText "InfoLabelVersion") }
    if ($script:infoLabelAuthor) { $script:infoLabelAuthor.Text = (Get-LocalizedText "InfoLabelAuthor") }
    if ($script:infoLabelDescription) { $script:infoLabelDescription.Text = (Get-LocalizedText "InfoLabelDescription") }
    
    # Onglet Gestion des OU (NOUVEAU)
    if ($script:tabPageOU) { $script:tabPageOU.Text = (Get-LocalizedText "OuTabTitle") }
    if ($script:createOuGroupBox) { $script:createOuGroupBox.Text = (Get-LocalizedText "CreateOuGroupBox") }
    if ($script:createOuNameLabel) { $script:createOuNameLabel.Text = (Get-LocalizedText "CreateOuNameLabel") }
    if ($script:createOuParentPathLabel) { $script:createOuParentPathLabel.Text = (Get-LocalizedText "CreateOuParentPathLabel") }
    if ($script:createOuButton) { $script:createOuButton.Text = (Get-LocalizedText "CreateOuButton") }
    if ($script:deleteOuGroupBox) { $script:deleteOuGroupBox.Text = (Get-LocalizedText "DeleteOuGroupBox") }
    if ($script:deleteOuNameLabel) { $script:deleteOuNameLabel.Text = (Get-LocalizedText "DeleteOuNameLabel") }
    if ($script:deleteOuButton) { $script:deleteOuButton.Text = (Get-LocalizedText "DeleteOuButton") }


    # Autres elements globaux
    if ($script:outputLabel) { $script:outputLabel.Text = (Get-LocalizedText "LogsLabel") }
    if ($script:languageLabel) { $script:languageLabel.Text = (Get-LocalizedText "LanguageLabel") }
}

#endregion

#region Affichage de la fenêtre

$script:mainForm.Add_Shown({
    $script:mainForm.Activate()
    Update-GuiLanguage # Initialiser les textes avec la langue par defaut
    Populate-ADDropdowns # Remplir les dropdowns au demarrage, y compris les OUs
})

$script:tabControl.Add_SelectedIndexChanged({
    # Rafraîchir les listes seulement pour les onglets qui les utilisent
    if ($script:tabControl.SelectedTab -eq $script:tabPageUsers -or $script:tabControl.SelectedTab -eq $script:tabPageGroups -or $script:tabControl.SelectedTab -eq $script:tabPageOU) {
        Populate-ADDropdowns # Utiliser la fonction qui rafraîchit aussi les OUs
    }
})

[void]$script:mainForm.ShowDialog()
#endregion