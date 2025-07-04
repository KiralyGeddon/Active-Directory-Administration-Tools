#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Librairies Communes ###
#
# Description :
# Ce script contient les fonctions utilitaires et la configuration de localisation
# utilisees par l'application principale et les modules d'onglets.
#-----------------------------------------------------------------------------------

#region 0. CONFIGURATION ET CHAÎNES DE LANGUE

# Langue par defaut (definie dans main.ps1, mais incluse ici pour reference)
# $CurrentLanguage = "en" # ou "fr"

# Tableau de hachage pour les chaînes de texte par langue
$script:Localization = @{
    "fr" = @{
        "MainWindowTitle"           = "Outil de Gestion Active Directory - v3.0"
        "ImportTabTitle"            = "Import en Masse (CSV)"
        "ImportLabelPath"           = "Chemin vers le fichier .csv :"
        "BrowseButton"              = "Parcourir..."
        "ImportButton"              = "Lancer l'Importation"
        "UsersTabTitle"             = "Gestion des Utilisateurs"
        "AddUserGroupBox"           = "Ajouter un utilisateur"
        "AddUserFirstNameLabel"     = "Prenom :"
        "AddUserLastNameLabel"      = "Nom :"
        "AddUserFunctionLabel"      = "Fonction :"
        "AddUserButton"             = "Ajouter l'utilisateur"
        "DeleteUserGroupBox"        = "Supprimer un utilisateur"
        "DeleteUserSelectUserLabel" = "Selectionner l'utilisateur (Login) :"
        "DeleteUserButton"          = "Supprimer l'utilisateur"
        "GroupsTabTitle"            = "Gestion des Groupes"
        "CreateGroupGroupBox"       = "Creer un groupe"
        "CreateGroupNameLabel"      = "Nom du groupe :"
        "CreateGroupButton"         = "Creer"
        "DeleteGroupGroupBox"       = "Supprimer un groupe"
        "DeleteGroupSelectGroupLabel"= "Selectionner le groupe :"
        "DeleteGroupButton"         = "Supprimer"
        "AddUserToGroupGroupBox"    = "Ajouter un utilisateur à un groupe"
        "AddUserToGroupSelectUserLabel"= "Selectionner l'utilisateur (Login) :"
        "AddUserToGroupSelectGroupLabel"= "Selectionner le groupe :"
        "AddUserToGroupButton"      = "Ajouter au groupe"
        "LogsLabel"                 = "Logs des operations :"
        "LanguageLabel"             = "Langue :"
        "InfoTabTitle"              = "Informations"
        "InfoLabelTitle"            = "Outil de Gestion Active Directory"
        "InfoLabelVersion"          = "Version : 3.0"
        "InfoLabelAuthor"           = "Auteur : Kiraly Geddon, aide par Gemini"
        "InfoLabelDescription"      = "Description : Ce script fournit une interface graphique (GUI) pour effectuer des operations courantes dans Active Directory."

        # NOUVEAU: Gestion des OU
        "OuTabTitle"                = "Gestion des OU"
        "CreateOuGroupBox"          = "Créer une Unité d'Organisation (OU)"
        "CreateOuNameLabel"         = "Nom de l'OU :"
        "CreateOuParentPathLabel"   = "Chemin parent (optionnel, ex: RH,OU=Services) :"
        "CreateOuButton"            = "Créer l'OU"
        "DeleteOuGroupBox"          = "Supprimer une Unité d'Organisation (OU)"
        "DeleteOuNameLabel"         = "Sélectionner l'OU :"
        "DeleteOuButton"            = "Supprimer l'OU"
        "MsgErrorOuNameRequired"    = "Le nom de l'OU est requis."
        "LogOuExists"               = "L'OU '{0}' existe déjà sous '{1}'."
        "LogOuCreated"              = "OU '{0}' créée sous '{1}'."
        "LogErrorOuCreation"        = "Erreur lors de la création de l'OU '{0}' : {1}"
        "MsgErrorSelectOuToDelete"  = "Veuillez sélectionner une OU à supprimer."
        "MsgConfirmDeleteOu"        = "Êtes-vous sûr de vouloir supprimer définitivement l'OU '{0}' ? (Doit être vide)"
        "LogOuDeleted"              = "OU '{0}' supprimée avec succès."
        "LogErrorOuDeletion"        = "Impossible de supprimer l'OU '{0}'. Elle n'existe peut-être pas ou n'est pas vide. Erreur: {1}"
        "LogOUsRefreshed"           = "Liste des OUs rafraîchie."
        "LogErrorOUsRefresh"        = "Erreur lors du rafraîchissement de la liste des OUs : {0}"


        # Messages de log et d'erreur existants
        "LogRefreshLists"           = "Rafraichissement des listes utilisateurs, groupes et OUs..."
        "LogUsersRefreshed"         = "Liste des utilisateurs rafraichie."
        "LogErrorUsersRefresh"      = "Erreur lors du rafraichissement de la liste des utilisateurs : {0}"
        "LogGroupsRefreshed"        = "Liste des groupes rafraichie."
        "LogErrorGroupsRefresh"     = "Erreur lors du rafraichissement de la liste des groupes : {0}"
        "MsgErrorInvalidCsv"        = "Le chemin specifie est invalide ou ce n'est pas un fichier .csv."
        "MsgErrorCsvRead"           = "Impossible de lire le fichier CSV. Verifiez le format, l'encodage (UTF8) et le delimiteur (;)."
        "LogCsvImportStarted"       = "Importation du fichier {0} demarree."
        "LogOuCreated"              = "OU '{0}' creee à la racine du domaine." # This might conflict with the new OU specific one
        "LogFunctionOuCreated"      = "Creation de l'OU : {0}"
        "LogUserExists"             = "L'utilisateur {0} existe deja."
        "LogUserCreated"            = "Utilisateur cree : {0} ({1} {2})"
        "LogErrorUserCreation"      = "Erreur lors de la creation de l'utilisateur {0} : {1}"
        "LogImportFinished"         = "Importation terminee."
        "MsgErrorAddUserFields"     = "Tous les champs (Nom, Prenom, Fonction) sont requis."
        "LogUserCreatedManual"      = "Utilisateur cree manuellement : {0}"
        "LogErrorUserCreationManual"= "Erreur lors de la creation manuelle de {0} : {1}"
        "MsgErrorSelectUserToDelete"= "Veuillez selectionner un utilisateur a supprimer."
        "MsgConfirmDeleteUser"      = "Etes-vous sur de vouloir supprimer definitivement l'utilisateur '{0}' ?"
        "LogUserDeleted"            = "Utilisateur '{0}' supprime avec succes."
        "LogErrorUserDeletion"      = "Impossible de supprimer l'utilisateur '{0}'. Il n'existe peut-etre pas ou une erreur est survenue. Erreur: {1}"
        "MsgErrorCreateGroupField"  = "Le champ 'Nom du groupe' est requis."
        "LogGroupExists"            = "Le groupe '{0}' existe deja."
        "LogGroupCreated"           = "Groupe '{0}' cree avec succes."
        "LogErrorGroupCreation"     = "Erreur lors de la creation du groupe '{0}' : {1}"
        "MsgErrorSelectGroupToDelete"= "Veuillez selectionner un groupe a supprimer."
        "MsgConfirmDeleteGroup"     = "Etes-vous sur de vouloir supprimer definitivement le groupe '{0}' ?"
        "LogGroupDeleted"           = "Groupe '{0}' supprime avec succes."
        "LogErrorGroupDeletion"     = "Impossible de supprimer le groupe '{0}'. Il n'existe peut-être pas ou une erreur est survenue. Erreur: {1}"
        "MsgErrorAddUserToGroupFields"= "Les champs 'Login utilisateur' et 'Nom du groupe' sont requis."
        "LogUserAddedToGroup"       = "Utilisateur '{0}' ajoute au groupe '{1}'."
        "LogErrorAddUserToGroup"    = "Erreur lors de l'ajout de '{0}' a '{1}'. Verifiez que l'utilisateur et le groupe existent. Erreur: {2}"
        "PrereqAdminNeeded"         = "Les droits d'administrateur sont requis. Tentative de relancement du script en mode eleve..."
        "PrereqModuleNotFound"      = "Le module Active Directory (RSAT) est introuvable sur cet ordinateur.`n`nCe composant est indispensable au fonctionnement du script.`n`nVoulez-vous tenter de l'installer maintenant ?`n(Cela necessite une connexion Internet et peut prendre quelques minutes)"
        "PrereqInstallStarting"     = "L'installation va commencer. Veuillez patienter..."
        "PrereqInstallSuccess"      = "L'installation des outils Active Directory a reussi !`n`nVeuillez relancer le script pour continuer."
        "PrereqInstallFailed"       = "L'installation a echoue. Assurez-vous d'avoir une connexion Internet et que votre systeme est à jour.`n`nErreur : {0}"
        "PrereqOperationCancelled"  = "Le script ne peut pas continuer sans ce composant."
        "ErrorCriticalAd"           = "Impossible de charger le module AD ou de detecter le domaine."
        "ConfirmDeletionTitle"      = "Confirmation de suppression"
        "PrereqMissingTitle"        = "Composant Manquant"
        "PrereqInstallFinishedTitle"= "Installation Terminee"
        "PrereqInstallErrorTitle"   = "Erreur d'Installation"
        "OperationCancelledTitle"   = "Operation Annulee"

    }
    "en" = @{
        "MainWindowTitle"           = "Active Directory Management Tool - v3.0"
        "ImportTabTitle"            = "Bulk Import (CSV)"
        "ImportLabelPath"           = "Path to .csv file :"
        "BrowseButton"              = "Browse..."
        "ImportButton"              = "Start Import"
        "UsersTabTitle"             = "User Management"
        "AddUserGroupBox"           = "Add User"
        "AddUserFirstNameLabel"     = "First Name :"
        "AddUserLastNameLabel"      = "Last Name :"
        "AddUserFunctionLabel"      = "Function :"
        "AddUserButton"             = "Add User"
        "DeleteUserGroupBox"        = "Delete User"
        "DeleteUserSelectUserLabel" = "Select User (Login) :"
        "DeleteUserButton"          = "Delete User"
        "GroupsTabTitle"            = "Group Management"
        "CreateGroupGroupBox"       = "Create Group"
        "CreateGroupNameLabel"      = "Group Name :"
        "CreateGroupButton"         = "Create"
        "DeleteGroupGroupBox"       = "Delete Group"
        "DeleteGroupSelectGroupLabel"= "Select Group :"
        "DeleteGroupButton"         = "Delete"
        "AddUserToGroupGroupBox"    = "Add User to Group"
        "AddUserToGroupSelectUserLabel"= "Select User (Login) :"
        "AddUserToGroupSelectGroupLabel"= "Select Group :"
        "AddUserToGroupButton"      = "Add to Group"
        "LogsLabel"                 = "Operation Logs :"
        "LanguageLabel"             = "Language :"
        "InfoTabTitle"              = "Information"
        "InfoLabelTitle"            = "Active Directory Management Tool"
        "InfoLabelVersion"          = "Version : 3.0"
        "InfoLabelAuthor"           = "Author : Kiraly Geddon, assisted by Gemini"
        "InfoLabelDescription"      = "Description : This script provides a graphical interface (GUI) to perform common Active Directory operations."

        # NEW: OU Management
        "OuTabTitle"                = "OU Management"
        "CreateOuGroupBox"          = "Create Organizational Unit (OU)"
        "CreateOuNameLabel"         = "OU Name :"
        "CreateOuParentPathLabel"   = "Parent Path (optional, e.g.: HR,OU=Services) :"
        "CreateOuButton"            = "Create OU"
        "DeleteOuGroupBox"          = "Delete Organizational Unit (OU)"
        "DeleteOuNameLabel"         = "Select OU :"
        "DeleteOuButton"            = "Delete OU"
        "MsgErrorOuNameRequired"    = "OU name is required."
        "LogOuExists"               = "OU '{0}' already exists under '{1}'."
        "LogOuCreated"              = "OU '{0}' created under '{1}'."
        "LogErrorOuCreation"        = "Error creating OU '{0}' : {1}"
        "MsgErrorSelectOuToDelete"  = "Please select an OU to delete."
        "MsgConfirmDeleteOu"        = "Are you sure you want to permanently delete OU '{0}'? (Must be empty)"
        "LogOuDeleted"              = "OU '{0}' deleted successfully."
        "LogErrorOuDeletion"        = "Could not delete OU '{0}'. It may not exist or is not empty. Error: {1}"
        "LogOUsRefreshed"           = "OU list refreshed."
        "LogErrorOUsRefresh"        = "Error refreshing OU list: {0}"


        # Existing Log and Error Messages
        "LogRefreshLists"           = "Refreshing user, group and OU lists..."
        "LogUsersRefreshed"         = "User list refreshed."
        "LogErrorUsersRefresh"      = "Error refreshing user list: {0}"
        "LogGroupsRefreshed"        = "Group list refreshed."
        "LogErrorGroupsRefresh"     = "Error refreshing group list: {0}"
        "MsgErrorInvalidCsv"        = "Invalid path or not a .csv file."
        "MsgErrorCsvRead"           = "Could not read CSV file. Check format, encoding (UTF8), and delimiter (;)."
        "LogCsvImportStarted"       = "CSV file {0} import started."
        "LogOuCreated"              = "OU '{0}' created at domain root." # This might conflict with the new OU specific one
        "LogFunctionOuCreated"      = "Creating OU: {0}"
        "LogUserExists"             = "User {0} already exists."
        "LogUserCreated"            = "User created: {0} ({1} {2})"
        "LogErrorUserCreation"      = "Error creating user {0}: {1}"
        "LogImportFinished"         = "Import finished."
        "MsgErrorAddUserFields"     = "All fields (First Name, Last Name, Function) are required."
        "LogUserCreatedManual"      = "User created manually: {0}"
        "LogErrorUserCreationManual"= "Error manually creating {0}: {1}"
        "MsgErrorSelectUserToDelete"= "Please select a user to delete."
        "MsgConfirmDeleteUser"      = "Are you sure you want to permanently delete user '{0}' ?"
        "LogUserDeleted"            = "User '{0}' deleted successfully."
        "LogErrorUserDeletion"      = "Could not delete user '{0}'. It may not exist or an error occurred. Error: {1}"
        "MsgErrorCreateGroupField"  = "The 'Group Name' field is required."
        "LogGroupExists"            = "Group '{0}' already exists."
        "LogGroupCreated"           = "Group '{0}' created successfully."
        "LogErrorGroupCreation"     = "Error creating group '{0}': {1}"
        "MsgErrorSelectGroupToDelete"= "Please select a group to delete."
        "MsgConfirmDeleteGroup"     = "Are you sure you want to permanently delete group '{0}' ?"
        "LogGroupDeleted"           = "Group '{0}' deleted successfully."
        "LogErrorGroupDeletion"     = "Could not delete group '{0}'. It may not exist or an error occurred. Error: {1}"
        "MsgErrorAddUserToGroupFields"= "The 'User Login' and 'Group Name' fields are required."
        "LogUserAddedToGroup"       = "User '{0}' added to group '{1}'."
        "LogErrorAddUserToGroup"    = "Error adding '{0}' to '{1}'. Verify user and group exist. Error: {2}"
        "PrereqAdminNeeded"         = "Administrator rights are required. Attempting to relaunch script in elevated mode..."
        "PrereqModuleNotFound"      = "The Active Directory module (RSAT) is not found on this computer.`n`nThis component is essential for the script to function.`n`nDo you want to try to install it now?`n(This requires an internet connection and may take a few minutes)"
        "PrereqInstallStarting"     = "Installation will begin. Please wait..."
        "PrereqInstallSuccess"      = "Active Directory tools installation successful! `n`nPlease restart the script to continue."
        "PrereqInstallFailed"       = "Installation failed. Ensure you have an internet connection and your system is up to date.`n`nError : {0}"
        "PrereqOperationCancelled"  = "The script cannot continue without this component."
        "ErrorCriticalAd"           = "Unable to load AD module or detect domain."
        "ConfirmDeletionTitle"      = "Confirm Deletion"
        "PrereqMissingTitle"        = "Missing Component"
        "PrereqInstallFinishedTitle"= "Installation Finished"
        "PrereqInstallErrorTitle"   = "Installation Error"
        "OperationCancelledTitle"   = "Operation Cancelled"
    }
}

# Fonction utilitaire pour obtenir la chaîne localisee
function Get-LocalizedText {
    param([string]$Key, [object[]]$Args)
    $text = $script:Localization[$script:CurrentLanguage].$Key
    if ($null -eq $text) {
        # Fallback to English if key is not found in current language
        $text = $script:Localization["en"].$Key
        if ($null -eq $text) {
            # Fallback to key itself if not found in English either
            $text = $Key
        }
    }
    if ($Args) {
        return [string]::Format($text, $Args)
    }
    return $text
}

# Fonction de log (outputTextBox doit être défini globalement dans main.ps1)
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    if ($script:outputTextBox) {
        $script:outputTextBox.AppendText("[$Type] - $(Get-Date -Format 'HH:mm:ss') - $Message`r`n")
    } else {
        Write-Host "[$Type] - $(Get-Date -Format 'HH:mm:ss') - $Message"
    }
}

# Fonction pour populer les ComboBox AD (dépend de $domainSuffix et des comboboxes définies dans les modules)
function Populate-ADDropdowns {
    Write-Log (Get-LocalizedText "LogRefreshLists")

    # Clear relevant comboboxes - global access
    if ($script:deleteUserLoginComboBox) { $script:deleteUserLoginComboBox.Items.Clear() }
    if ($script:addUserToGroupLoginComboBox) { $script:addUserToGroupLoginComboBox.Items.Clear() }
    if ($script:deleteGroupNameComboBox) { $script:deleteGroupNameComboBox.Items.Clear() }
    if ($script:addUserToGroupGroupNameComboBox) { $script:addUserToGroupGroupNameComboBox.Items.Clear() }
    if ($script:deleteOuNameComboBox) { $script:deleteOuNameComboBox.Items.Clear() } # Clear OU ComboBox

    # Populate Users
    try {
        $adUsers = Get-ADUser -Filter * -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName | Sort-Object
        foreach ($userLogin in $adUsers) {
            if ($script:deleteUserLoginComboBox) { [void]$script:deleteUserLoginComboBox.Items.Add($userLogin) }
            if ($script:addUserToGroupLoginComboBox) { [void]$script:addUserToGroupLoginComboBox.Items.Add($userLogin) }
        }
        Write-Log (Get-LocalizedText "LogUsersRefreshed")
    }
    catch {
        Write-Log (Get-LocalizedText "LogErrorUsersRefresh" $_.Exception.Message) "ERREUR"
    }

    # Populate Groups
    try {
        $adGroups = Get-ADGroup -Filter * -Properties Name | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($groupName in $adGroups) {
            if ($script:deleteGroupNameComboBox) { [void]$script:deleteGroupNameComboBox.Items.Add($groupName) }
            if ($script:addUserToGroupGroupNameComboBox) { [void]$script:addUserToGroupGroupNameComboBox.Items.Add($groupName) }
        }
        Write-Log (Get-LocalizedText "LogGroupsRefreshed")
    }
    catch {
        Write-Log (Get-LocalizedText "LogErrorGroupsRefresh" $_.Exception.Message) "ERREUR"
    }

    # Populate OUs (NEW)
    try {
        $adOUs = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($ouName in $adOUs) {
            if ($script:deleteOuNameComboBox) { [void]$script:deleteOuNameComboBox.Items.Add($ouName) }
        }
        Write-Log (Get-LocalizedText "LogOUsRefreshed")
    }
    catch {
        Write-Log (Get-LocalizedText "LogErrorOUsRefresh" $_.Exception.Message) "ERREUR"
    }
}

#endregion