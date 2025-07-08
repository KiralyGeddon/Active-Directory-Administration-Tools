#-----------------------------------------------------------------------------------
#          Administration Tool ######
#
# Version : 2.9 (Internationalise)
# Auteur : Kiraly Geddon aide par Gemini
#
# Description :
# Ce script fournit une interface graphique (GUI) pour effectuer des operations
# courantes dans Active Directory : import en masse, ajout/suppression d'utilisateurs
# et gestion de groupes.
#
# Nouveautes de la v2.0 :
#   - Detection automatique du domaine Active Directory.
#   - Ajout de commentaires detailles pour chaque partie du code.
#
# Nouveautes de la v2.5 :
#   - Verification des droits d'administrateur et auto-elevation si necessaire.
#   - Verification de la presence des outils RSAT pour AD et proposition
#     d'installation automatique si manquants.
#
# Nouveautes pour v2.6 :
#   - Remplacement des TextBox par des ComboBox pour la selection d'utilisateurs/groupes existants.
#   - Correction de l'erreur "Multiple ambiguous overloads found for "Font"".
#   - Ajout de la fonction Populate-ADDropdowns pour remplir les ComboBox.
#   - Rafraîchissement automatique des ComboBox lors de la selection des onglets pertinents.
#
# Nouveautes pour v2.7 :
#   - Correction des problèmes d'accents (necessite d'enregistrer le script en UTF-8 sans BOM).
#   - Ajout d'une gestion multilingue (Français/Anglais) avec selection via ComboBox.
#   - Toutes les chaînes de l'interface et des logs sont externalisees pour la traduction.
#
# Nouveautes pour v2.8 :
#   - Corrections de quelques erreurs à l'import CSV.
# Nouveautes de la v2.9 :
#   - Ajout d'un onglet pour la création et la suppression des Unités d'Organisation (OU).
#   - Mise à jour de la version et du changelog.
#   - Amélioration de la robustesse et de l'expérience utilisateur.
#-----------------------------------------------------------------------------------

#region 0. CONFIGURATION ET CHAÎNES DE LANGUE

# Langue par defaut du script. Peut être "en" (Anglais) ou "fr" (Français).
$CurrentLanguage = "en" # ou "fr"

# Tableau de hachage pour stocker toutes les chaînes de texte de l'interface et des messages de log.
# Cela permet une internationalisation facile du script.
$Localization = @{
    # Définition des textes pour la langue Française
    "fr" = @{
        "MainWindowTitle"           = "Outil de Gestion Active Directory - v2.9"
        "ImportTabTitle"            = "Import en Masse (CSV)"
        "ImportLabelPath"           = "Chemin vers le fichier .csv :"
        "BrowseButton"              = "Parcourir..."
        "ImportButton"              = "Lancer l'Importation"
        "UsersTabTitle"             = "Gestion des Utilisateurs"
        "AddUserGroupBox"           = "Ajouter un utilisateur"
        "AddUserFirstNameLabel"     = "Prénom :"
        "AddUserLastNameLabel"      = "Nom :"
        "AddUserFunctionLabel"      = "Fonction :"
        "AddUserButton"             = "Ajouter l'utilisateur"
        "DeleteUserGroupBox"        = "Supprimer un utilisateur"
        "DeleteUserSelectUserLabel" = "Sélectionner l'utilisateur (Login) :"
        "DeleteUserButton"          = "Supprimer l'utilisateur"
        "GroupsTabTitle"            = "Gestion des Groupes"
        "CreateGroupGroupBox"       = "Créer un groupe"
        "CreateGroupNameLabel"      = "Nom du groupe :"
        "CreateGroupButton"         = "Créer"
        "DeleteGroupGroupBox"       = "Supprimer un groupe"
        "DeleteGroupSelectGroupLabel"= "Sélectionner le groupe :"
        "DeleteGroupButton"         = "Supprimer"
        "AddUserToGroupGroupBox"    = "Ajouter un utilisateur à un groupe"
        "AddUserToGroupSelectUserLabel"= "Sélectionner l'utilisateur (Login) :"
        "AddUserToGroupSelectGroupLabel"= "Sélectionner le groupe :"
        "AddUserToGroupButton"      = "Ajouter au groupe"
        "OUTabTitle"                = "Gestion des OUs" # NOUVEAU
        "CreateOUGroupBox"          = "Créer une Unité d'Organisation" # NOUVEAU
        "CreateOUNameLabel"         = "Nom de la nouvelle OU :" # NOUVEAU
        "CreateOUParentLabel"       = "Emplacement (OU parente) :" # NOUVEAU
        "CreateOUButton"            = "Créer l'OU" # NOUVEAU
        "DeleteOUGroupBox"          = "Supprimer une Unité d'Organisation" # NOUVEAU
        "DeleteOUSelectLabel"       = "Sélectionner l'OU à supprimer :" # NOUVEAU
        "DeleteOURecursiveLabel"    = "Suppression récursive (si l'OU contient des objets)" # NOUVEAU
        "DeleteOUButton"            = "Supprimer l'OU" # NOUVEAU
        "LogsLabel"                 = "Logs des opérations :"
        "LanguageLabel"             = "Langue :"
        "InformationTabTitle"       = "Information"
        "PrereqMissingTitle"        = "Module Active Directory manquant"
        "PrereqInstallFinishedTitle"= "Installation terminée"
        "PrereqInstallErrorTitle"   = "Erreur d'installation"
        "OperationCancelledTitle"   = "Opération annulée"
        "ConfirmDeletionTitle"      = "Confirmer la suppression"

        # Messages de log et d'erreur
        "LogRefreshLists"           = "Rafraîchissement des listes (Utilisateurs, Groupes, OUs)..."
        "LogUsersRefreshed"         = "Liste des utilisateurs rafraîchie."
        "LogErrorUsersRefresh"      = "Erreur lors du rafraîchissement de la liste des utilisateurs : {0}"
        "LogGroupsRefreshed"        = "Liste des groupes rafraîchie."
        "LogErrorGroupsRefresh"     = "Erreur lors du rafraîchissement de la liste des groupes : {0}"
        "LogOURefreshed"            = "Liste des OUs rafraîchie." # NOUVEAU
        "LogErrorOURefresh"         = "Erreur lors du rafraîchissement de la liste des OUs : {0}" # NOUVEAU
        "MsgErrorInvalidCsv"        = "Le chemin spécifié est invalide ou ce n'est pas un fichier .csv."
        "MsgErrorCsvRead"           = "Impossible de lire le fichier CSV. Vérifiez le format, l'encodage (UTF8) et le délimiteur (;)."
        "LogCsvImportStarted"       = "Importation du fichier {0} démarrée."
        "LogOuCreated"              = "OU '{0}' créée dans '{1}'." # Mis à jour
        "LogFunctionOuCreated"      = "Création de l'OU : {0}"
        "LogUserExists"             = "L'utilisateur {0} existe déjà."
        "LogUserCreated"            = "Utilisateur créé : {0} ({1} {2})"
        "LogErrorUserCreation"      = "Erreur lors de la création de l'utilisateur {0} : {1}"
        "LogImportFinished"         = "Importation terminée."
        "MsgErrorAddUserFields"     = "Tous les champs (Nom, Prénom, Fonction) sont requis."
        "LogUserCreatedManual"      = "Utilisateur créé manuellement : {0}"
        "LogErrorUserCreationManual"= "Erreur lors de la création manuelle de {0} : {1}"
        "MsgErrorSelectUserToDelete"= "Veuillez sélectionner un utilisateur à supprimer."
        "MsgConfirmDeleteUser"      = "Êtes-vous sûr de vouloir supprimer définitivement l'utilisateur '{0}' ?"
        "LogUserDeleted"            = "Utilisateur '{0}' supprimé avec succès."
        "LogErrorUserDeletion"      = "Impossible de supprimer l'utilisateur '{0}'. Il n'existe peut-être pas ou une erreur est survenue. Erreur: {1}"
        "MsgErrorCreateGroupField"  = "Le champ 'Nom du groupe' est requis."
        "LogGroupExists"            = "Le groupe '{0}' existe déjà."
        "LogGroupCreated"           = "Groupe '{0}' créé avec succès."
        "LogErrorGroupCreation"     = "Erreur lors de la création du groupe '{0}' : {1}"
        "MsgErrorSelectGroupToDelete"= "Veuillez sélectionner un groupe à supprimer."
        "MsgConfirmDeleteGroup"     = "Êtes-vous sûr de vouloir supprimer définitivement le groupe '{0}' ?"
        "LogGroupDeleted"           = "Groupe '{0}' supprimé avec succès."
        "LogErrorGroupDeletion"     = "Impossible de supprimer le groupe '{0}'. Il n'existe peut-être pas ou une erreur est survenue. Erreur: {1}"
        "MsgErrorAddUserToGroupFields"= "Les champs 'Login utilisateur' et 'Nom du groupe' sont requis."
        "LogUserAddedToGroup"       = "Utilisateur '{0}' ajouté au groupe '{1}'."
        "LogErrorAddUserToGroup"    = "Erreur lors de l'ajout de '{0}' à '{1}'. Vérifiez que l'utilisateur et le groupe existent. Erreur: {2}"
        "MsgErrorCreateOUFields"    = "Le nom de l'OU et son emplacement parent sont requis." # NOUVEAU
        "LogOUExists"               = "L'OU '{0}' existe déjà dans '{1}'." # NOUVEAU
        "LogErrorOUCreation"        = "Erreur lors de la création de l'OU '{0}' : {1}" # NOUVEAU
        "MsgErrorSelectOUToDelete"  = "Veuillez sélectionner une OU à supprimer." # NOUVEAU
        "MsgConfirmDeleteOU"        = "Êtes-vous sûr de vouloir supprimer définitivement l'OU '{0}' ? ATTENTION : Ceci est irréversible." # NOUVEAU
        "MsgConfirmDeleteOURecursive"= "La suppression récursive effacera l'OU '{0}' ET TOUS LES OBJETS qu'elle contient (utilisateurs, groupes, autres OUs...). Confirmez-vous cette action ? " # NOUVEAU
        "LogOUDeleted"              = "OU '{0}' supprimée avec succès." # NOUVEAU
        "LogErrorOUDeletion"        = "Erreur lors de la suppression de l'OU '{0}'. Elle est peut-être protégée contre la suppression accidentelle ou une autre erreur est survenue : {1}" # NOUVEAU
        "PrereqAdminNeeded"         = "Les droits d'administrateur sont requis. Tentative de relancement du script en mode élevé..."
        "PrereqModuleNotFound"      = "Le module Active Directory (RSAT) est introuvable sur cet ordinateur.`n`nCe composant est indispensable au fonctionnement du script.`n`nVoulez-vous tenter de l'installer maintenant ?`n(Cela nécessite une connexion Internet et peut prendre quelques minutes)"
        "PrereqInstallStarting"     = "L'installation va commencer. Veuillez patienter..."
        "PrereqInstallSuccess"      = "L'installation des outils Active Directory a réussi !`n`nVeuillez relancer le script pour continuer."
        "PrereqInstallFailed"       = "L'installation a échoué. Assurez-vous d'avoir une connexion Internet et que votre système est à jour.`n`nErreur : {0}"
        "PrereqOperationCancelled"  = "Le script ne peut pas continuer sans ce composant."
        "ErrorCriticalAd"           = "Impossible de charger le module AD ou de détecter le domaine."
    }
    # Définition des textes pour la langue Anglaise
    "en" = @{
        "MainWindowTitle"           = "Active Directory Administration Tool - v2.9"
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
        "OUTabTitle"                = "OU Management" # NEW
        "CreateOUGroupBox"          = "Create an Organizational Unit" # NEW
        "CreateOUNameLabel"         = "New OU Name:" # NEW
        "CreateOUParentLabel"       = "Location (Parent OU):" # NEW
        "CreateOUButton"            = "Create OU" # NEW
        "DeleteOUGroupBox"          = "Delete an Organizational Unit" # NEW
        "DeleteOUSelectLabel"       = "Select OU to delete:" # NEW
        "DeleteOURecursiveLabel"    = "Recursive delete (if OU contains objects)" # NEW
        "DeleteOUButton"            = "Delete OU" # NEW
        "LogsLabel"                 = "Operation Logs :"
        "LanguageLabel"             = "Language :"
        "InformationTabTitle"       = "Information"
        "PrereqMissingTitle"        = "Missing Active Directory Module"
        "PrereqInstallFinishedTitle"= "Installation Finished"
        "PrereqInstallErrorTitle"   = "Installation Error"
        "OperationCancelledTitle"   = "Operation Cancelled"
        "ConfirmDeletionTitle"      = "Confirm Deletion"

        # Log and Error Messages
        "LogRefreshLists"           = "Refreshing lists (Users, Groups, OUs)..."
        "LogUsersRefreshed"         = "User list refreshed."
        "LogErrorUsersRefresh"      = "Error refreshing user list: {0}"
        "LogGroupsRefreshed"        = "Group list refreshed."
        "LogErrorGroupsRefresh"     = "Error refreshing group list: {0}"
        "LogOURefreshed"            = "OU list refreshed." # NEW
        "LogErrorOURefresh"         = "Error refreshing OU list: {0}" # NEW
        "MsgErrorInvalidCsv"        = "Invalid path or not a .csv file."
        "MsgErrorCsvRead"           = "Could not read CSV file. Check format, encoding (UTF8), and delimiter (;)."
        "LogCsvImportStarted"       = "CSV file {0} import started."
        "LogOuCreated"              = "OU '{0}' created in '{1}'." # Updated
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
        "MsgErrorCreateOUFields"    = "OU Name and Parent Location are required." # NEW
        "LogOUExists"               = "OU '{0}' already exists in '{1}'." # NEW
        "LogErrorOUCreation"        = "Error creating OU '{0}': {1}" # NEW
        "MsgErrorSelectOUToDelete"  = "Please select an OU to delete." # NEW
        "MsgConfirmDeleteOU"        = "Are you sure you want to permanently delete the OU '{0}'? WARNING: This is irreversible." # NEW
        "MsgConfirmDeleteOURecursive"= "Recursive deletion will erase the OU '{0}' AND ALL OBJECTS within it (users, groups, other OUs...). Do you confirm this action?" # NEW
        "LogOUDeleted"              = "OU '{0}' deleted successfully." # NEW
        "LogErrorOUDeletion"        = "Error deleting OU '{0}'. It might be protected from accidental deletion or another error occurred: {1}" # NEW
        "PrereqAdminNeeded"         = "Administrator rights are required. Attempting to relaunch script in elevated mode..."
        "PrereqModuleNotFound"      = "The Active Directory module (RSAT) is not found on this computer.`n`nThis component is essential for the script to function.`n`nDo you want to try to install it now?`n(This requires an internet connection and may take a few minutes)"
        "PrereqInstallStarting"     = "Installation will begin. Please wait..."
        "PrereqInstallSuccess"      = "Active Directory tools installation successful! `n`nPlease restart the script to continue."
        "PrereqInstallFailed"       = "Installation failed. Ensure you have an internet connection and your system is up to date.`n`nError: {0}"
        "PrereqOperationCancelled"  = "The script cannot continue without this component."
        "ErrorCriticalAd"           = "Unable to load AD module or detect domain."
    }
}

# Fonction utilitaire pour obtenir la chaîne localisée en fonction de la langue actuelle.
function Get-LocalizedText {
    param([string]$Key, [object[]]$Args)
    $text = $Localization[$CurrentLanguage].$Key
    # Si la clé n'est pas trouvée dans la langue actuelle, on essaie l'anglais comme fallback.
    if ($null -eq $text) {
        $text = $Localization["en"].$Key
        # Si la clé n'est pas trouvée non plus en anglais, on retourne la clé elle-même.
        if ($null -eq $text) {
            $text = $Key
        }
    }
    # Si des arguments sont fournis, on formate la chaîne.
    if ($Args) {
        try {
            # On passe directement le tableau $Args. Le paramètre [object[]]$Args garantit que c'est déjà un tableau.
            return [string]::Format($text, $Args)
        }
        catch [System.FormatException] {
            # Ce bloc catch empêche le script de s'arrêter si le nombre de marqueurs {0}, {1}...
            # dans la chaîne de localisation ne correspond pas au nombre d'arguments fournis.
            Write-Warning "Erreur de formatage pour la clé '$Key'. Incompatibilité entre les marqueurs et les arguments. Texte brut : `"$text`""
            return $text
        }
    }
    return $text
}

#endregion

#region 1. VeRIFICATION DES PReREQUIS (ADMIN & RSAT)

# --- A. Verification des droits d'Administrateur ---
# Vérifie si le script est exécuté avec des droits d'administrateur.
# Si ce n'est pas le cas, il tente de relancer le script en mode élevé (avec l'UAC).
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning (Get-LocalizedText "PrereqAdminNeeded") # Affiche un avertissement localisé.
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -File `"$PSCommandPath`"" # Relance le script en tant qu'administrateur.
    exit # Quitte l'instance actuelle du script.
}

# --- B. Verification et installation du module Active Directory (RSAT) ---
# Charge l'assemblage System.Windows.Forms pour pouvoir utiliser les boîtes de dialogue.
Add-Type -AssemblyName System.Windows.Forms

# Vérifie si le module ActiveDirectory est disponible (fait partie des Outils d'Administration Serveur RSAT).
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    # Si le module n'est pas trouvé, propose de l'installer.
    $installChoice = [System.Windows.Forms.MessageBox]::Show(
        (Get-LocalizedText "PrereqModuleNotFound"), # Message d'information localisé.
        (Get-LocalizedText "PrereqMissingTitle"), # Titre de la boîte de dialogue localisé.
        [System.Windows.Forms.MessageBoxButtons]::YesNo, # Boutons Oui/Non.
        [System.Windows.Forms.MessageBoxIcon]::Question # Icône de question.
    )

    # Si l'utilisateur choisit 'Oui' pour installer.
    if ($installChoice -eq 'Yes') {
        try {
            # Informe l'utilisateur que l'installation va commencer.
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallStarting"),
                "Information", "OK", "Information"
            )
            # Tente d'ajouter la fonctionnalité RSAT pour Active Directory.
            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
            # Informe l'utilisateur que l'installation a réussi.
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallSuccess"),
                (Get-LocalizedText "PrereqInstallFinishedTitle"), # Titre localisé.
                "OK", "Information"
            )
        }
        catch {
            # En cas d'erreur lors de l'installation, affiche un message d'erreur.
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedText "PrereqInstallFailed" $_.Exception.Message),
                (Get-LocalizedText "PrereqInstallErrorTitle"), # Titre localisé.
                "OK", "Error"
            )
        }
        finally {
            exit # Quitte le script après l'installation (réussie ou échouée).
        }
    }
    else {
        # Si l'utilisateur choisit 'Non', annule l'opération et quitte le script.
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "PrereqOperationCancelled"), (Get-LocalizedText "OperationCancelledTitle"), "OK", "Warning") # Message et titre localisés.
        exit
    }
}

#endregion

#region 2. DeTECTION DU DOMAINE ET CHARGEMENT DES MODULES

try {
    # Importe le module ActiveDirectory.
    Import-Module ActiveDirectory -ErrorAction Stop
    # Détecte le domaine Active Directory actuel.
    $adDomain = Get-ADDomain
    # Récupère le chemin distingué du domaine (ex: DC=mondomaine,DC=lan).
    $domainPath = $adDomain.DistinguishedName
    # Récupère le suffixe DNS du domaine (ex: mondomaine.lan).
    $domainSuffix = $adDomain.DNSRoot
}
catch {
    # En cas d'erreur critique (module AD introuvable ou domaine non détecté), affiche un message et quitte.
    [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "ErrorCriticalAd"), "Erreur Critique", "OK", "Error")
    exit
}

# Charge l'assemblage System.Drawing pour les fonctionnalités graphiques.
Add-Type -AssemblyName System.Drawing
#endregion


#region Chargement des Assemblages .NET pour l'Interface Graphique (GUI)
# Ces lignes sont redondantes avec la section 1.B mais ne posent pas de problème.
# Elles s'assurent que les bibliothèques nécessaires à la GUI sont chargées.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion

#region Fonctions Logiques (Modularite)

# Fonction pour écrire des messages dans la zone de logs de l'interface.
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    $outputTextBox.AppendText("[$Type] - $(Get-Date -Format 'HH:mm:ss') - $Message`r`n")
}

# Fonction pour remplir les ComboBox des listes d'utilisateurs et de groupes AD.
function Populate-ADDropdowns {
    Write-Log (Get-LocalizedText "LogRefreshLists") # Informe que les listes sont en cours de rafraîchissement.

    # Efface les éléments existants dans les ComboBox avant de les remplir.
    if ($deleteUserLoginComboBox) { $deleteUserLoginComboBox.Items.Clear() }
    if ($addUserToGroupLoginComboBox) { $addUserToGroupLoginComboBox.Items.Clear() }
    if ($deleteGroupNameComboBox) { $deleteGroupNameComboBox.Items.Clear() }
    if ($addUserToGroupGroupNameComboBox) { $addUserToGroupGroupNameComboBox.Items.Clear() }

    # Tente de récupérer et de remplir la liste des utilisateurs.
    try {
        $adUsers = Get-ADUser -Filter * -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName | Sort-Object
        foreach ($userLogin in $adUsers) {
            if ($deleteUserLoginComboBox) { [void]$deleteUserLoginComboBox.Items.Add($userLogin) }
            if ($addUserToGroupLoginComboBox) { [void]$addUserToGroupLoginComboBox.Items.Add($userLogin) }
        }
        Write-Log (Get-LocalizedText "LogUsersRefreshed") # Informe que la liste des utilisateurs est rafraîchie.
    }
    catch {
        Write-Log (Get-LocalizedText "LogErrorUsersRefresh" $_.Exception.Message) "ERREUR" # Affiche une erreur en cas de problème.
    }

    # Tente de récupérer et de remplir la liste des groupes.
    try {
        $adGroups = Get-ADGroup -Filter * -Properties Name | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($groupName in $adGroups) {
            if ($deleteGroupNameComboBox) { [void]$deleteGroupNameComboBox.Items.Add($groupName) }
            if ($addUserToGroupGroupNameComboBox) { [void]$addUserToGroupGroupNameComboBox.Items.Add($groupName) }
        }
        Write-Log (Get-LocalizedText "LogGroupsRefreshed") # Informe que la liste des groupes est rafraîchie.
    }
    catch {
        Write-Log (Get-LocalizedText "LogErrorGroupsRefresh" $_.Exception.Message) "ERREUR" # Affiche une erreur en cas de problème.
    }
}

# Fonction pour démarrer l'importation d'utilisateurs à partir d'un fichier CSV.
function Start-CsvUserImport {
    $csvPath = $csvPathTextBox.Text # Récupère le chemin du fichier CSV depuis la TextBox.
    # Vérifie si le chemin est valide et si c'est bien un fichier .csv.
    if (-not (Test-Path $csvPath) -or ($csvPath -notlike "*.csv")) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorInvalidCsv"), "Erreur", "OK", "Error")
        return
    }

    # Tente de lire le fichier CSV.
    try {
        $CSVData = Import-CSV -Path $csvPath -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogCsvImportStarted" $csvPath) # Informe que l'importation a démarré.
    }
    catch {
        Write-Log (Get-LocalizedText "MsgErrorCsvRead") "ERREUR" # Affiche une erreur si la lecture échoue.
        return
    }

    # Vérifie si l'Unité d'Organisation (OU) "Utilisateurs" existe, sinon la crée.
    if (-not(Get-ADOrganizationalUnit -Filter "Name -eq 'Utilisateurs'" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name "Utilisateurs" -Path $domainPath # Crée l'OU à la racine du domaine détecté.
        Write-Log (Get-LocalizedText "LogOuCreated" "Utilisateurs") # Informe de la création de l'OU.
    }

    # Parcourt chaque ligne du fichier CSV pour créer les utilisateurs.
    foreach ($Utilisateur in $CSVData) {
        $prenom = $Utilisateur.Prenom
        $nom = $Utilisateur.Nom
        $fonction = $Utilisateur.Fonction.Trim()
        # Génère le login (sAMAccountName) : première lettre du prénom + nom de famille.
        $login = "$($prenom.Substring(0,1).ToLower()).$($nom.ToLower())"
        # Génère l'adresse e-mail en utilisant le suffixe DNS du domaine détecté.
        $email = "$login@$domainSuffix"
        $password = "Azerty123456!" # Mot de passe par défaut.
        # Définit le chemin de l'OU spécifique à la fonction de l'utilisateur.
        $ouFonctionPath = "OU=$fonction,OU=Utilisateurs,$domainPath"

        # Vérifie si l'OU de la fonction existe, sinon la crée.
        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouFonctionPath'" -ErrorAction SilentlyContinue)) {
            New-ADOrganizationalUnit -Name $fonction -Path "OU=Utilisateurs,$domainPath" # Crée l'OU sous "Utilisateurs".
            Write-Log (Get-LocalizedText "LogFunctionOuCreated" $fonction) # Informe de la création de l'OU de fonction.
        }

        # Vérifie si l'utilisateur existe déjà.
        if (Get-ADUser -Filter { SamAccountName -eq $login }) {
            Write-Log (Get-LocalizedText "LogUserExists" $login) "AVERTISSEMENT" # Informe que l'utilisateur existe déjà.
        }
        else {
            # Tente de créer le nouvel utilisateur.
            try {
                New-ADUser -Name "$nom $prenom" `
                           -DisplayName "$nom $prenom" `
                           -GivenName $prenom `
                           -Surname $nom `
                           -SamAccountName $login `
                           -UserPrincipalName $email `
                           -EmailAddress $email `
                           -Title $fonction `
                           -Path $ouFonctionPath `
                           -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                           -ChangePasswordAtLogon $true `
                           -Enabled $true -ErrorAction Stop
                Write-Log (Get-LocalizedText "LogUserCreated" $login $nom $prenom) # Informe de la création réussie.
            }
            catch {
                Write-Log (Get-LocalizedText "LogErrorUserCreation" $login $_) "ERREUR" # Affiche une erreur en cas de problème.
            }
        }
    }
    Write-Log (Get-LocalizedText "LogImportFinished") # Informe que l'importation est terminée.
    Populate-ADDropdowns # Rafraîchit les listes déroulantes après l'importation.
}

# Fonction pour ajouter un utilisateur unique manuellement.
function Add-SingleUser {
    $prenom = $addUserPrenomTextBox.Text
    $nom = $addUserNomTextBox.Text
    $fonction = $addUserFonctionTextBox.Text

    # Vérifie que tous les champs requis sont remplis.
    if (-not $prenom -or -not $nom -or -not $fonction) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorAddUserFields"), "Erreur", "OK", "Error")
        return
    }

    $login = "$($prenom.Substring(0,1).ToLower()).$($nom.ToLower())"
    $email = "$login@$domainSuffix"
    $password = "Azerty123456!"
    $ouFonctionPath = "OU=$fonction,OU=Utilisateurs,$domainPath"

    # Vérifie et crée l'OU de fonction si nécessaire.
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouFonctionPath'" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $fonction -Path "OU=Utilisateurs,$domainPath"
        Write-Log (Get-LocalizedText "LogFunctionOuCreated" $fonction)
    }

    # Vérifie si l'utilisateur existe déjà.
    if (Get-ADUser -Filter { SamAccountName -eq $login }) {
        Write-Log (Get-LocalizedText "LogUserExists" $login) "AVERTISSEMENT"
    } else {
        # Tente de créer l'utilisateur.
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
            Write-Log (Get-LocalizedText "LogUserCreatedManual" $login) # Informe de la création.
            # Efface les champs après la création.
            $addUserPrenomTextBox.Clear()
            $addUserNomTextBox.Clear()
            $addUserFonctionTextBox.Clear()
            Populate-ADDropdowns # Rafraîchit les listes déroulantes.
        } catch {
            Write-Log (Get-LocalizedText "LogErrorUserCreationManual" $login $_) "ERREUR" # Affiche une erreur.
        }
    }
}

# Fonction pour supprimer un utilisateur sélectionné.
function Remove-SpecificUser {
    $login = $deleteUserLoginComboBox.SelectedItem # Récupère l'utilisateur sélectionné dans la ComboBox.
    # Vérifie si un utilisateur a été sélectionné.
    if (-not $login) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectUserToDelete"), "Erreur", "OK", "Error")
        return
    }

    # Demande confirmation avant de supprimer l'utilisateur.
    $confirm = [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgConfirmDeleteUser" $login), (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        # Tente de supprimer l'utilisateur.
        try {
            $user = Get-ADUser -Identity $login -ErrorAction Stop # Récupère l'objet utilisateur.
            Remove-ADUser -Identity $user -Confirm:$false -ErrorAction Stop # Supprime l'utilisateur sans confirmation supplémentaire.
            Write-Log (Get-LocalizedText "LogUserDeleted" $login) # Informe de la suppression réussie.
            $deleteUserLoginComboBox.SelectedIndex = -1 # Désélectionne l'élément dans la ComboBox.
            Populate-ADDropdowns # Rafraîchit les listes déroulantes.
        } catch {
            Write-Log (Get-LocalizedText "LogErrorUserDeletion" $login $_) "ERREUR" # Affiche une erreur.
        }
    }
}

# Fonction pour créer un nouveau groupe.
function Add-SpecificGroup {
    $groupName = $createGroupNameTextBox.Text # Récupère le nom du groupe.
    # Vérifie si le nom du groupe est renseigné.
    if (-not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorCreateGroupField"), "Erreur", "OK", "Error")
        return
    }

    # Tente de créer le groupe.
    try {
        # Vérifie si le groupe existe déjà.
        if(Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue) {
             Write-Log (Get-LocalizedText "LogGroupExists" $groupName) "AVERTISSEMENT" # Informe que le groupe existe déjà.
             return
        }
        New-ADGroup -Name $groupName -GroupScope Global -Path $domainPath -ErrorAction Stop # Crée le groupe à la racine du domaine détecté.
        Write-Log (Get-LocalizedText "LogGroupCreated" $groupName) # Informe de la création réussie.
        $createGroupNameTextBox.Clear() # Efface le champ.
        Populate-ADDropdowns # Rafraîchit les listes déroulantes.
    } catch {
        Write-Log (Get-LocalizedText "LogErrorGroupCreation" $groupName $_) "ERREUR" # Affiche une erreur.
    }
}

# Fonction pour supprimer un groupe sélectionné.
function Remove-SpecificGroup {
    $groupName = $deleteGroupNameComboBox.SelectedItem # Récupère le nom du groupe sélectionné.
    # Vérifie si un groupe a été sélectionné.
    if (-not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectGroupToDelete"), "Erreur", "OK", "Error")
        return
    }

    # Demande confirmation avant de supprimer le groupe.
    $confirm = [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgConfirmDeleteGroup" $groupName), (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        # Tente de supprimer le groupe.
        try {
            Get-ADGroup -Identity $groupName -ErrorAction Stop | Remove-ADGroup -Confirm:$false -ErrorAction Stop # Supprime le groupe.
            Write-Log (Get-LocalizedText "LogGroupDeleted" $groupName) # Informe de la suppression réussie.
            $deleteGroupNameComboBox.SelectedIndex = -1 # Désélectionne l'élément.
            Populate-ADDropdowns # Rafraîchit les listes déroulantes.
        } catch {
            Write-Log (Get-LocalizedText "LogErrorGroupDeletion" $groupName $_) "ERREUR" # Affiche une erreur.
        }
    }
}

# Fonction pour ajouter un utilisateur à un groupe.
function Add-UserToSpecificGroup {
    $login = $addUserToGroupLoginComboBox.SelectedItem # Récupère le login de l'utilisateur.
    $groupName = $addUserToGroupGroupNameComboBox.SelectedItem # Récupère le nom du groupe.

    # Vérifie que l'utilisateur et le groupe sont sélectionnés.
    if (-not $login -or -not $groupName) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorAddUserToGroupFields"), "Erreur", "OK", "Error")
        return
    }

    # Tente d'ajouter l'utilisateur au groupe.
    try {
        Add-ADGroupMember -Identity $groupName -Members $login -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogUserAddedToGroup" $login $groupName) # Informe de l'ajout réussi.
        $addUserToGroupLoginComboBox.SelectedIndex = -1 # Désélectionne l'utilisateur.
        $addUserToGroupGroupNameComboBox.SelectedIndex = -1 # Désélectionne le groupe.
        Populate-ADDropdowns # Rafraîchit les listes déroulantes.
    }# Inside Add-UserToSpecificGroup or related error handling
    catch {
    $errorMessage = $_.Exception.Message # Get the actual error message
    Write-Log (Get-LocalizedText "LogErrorAddUserToGroup" $login $groupName $errorMessage) # Add $errorMessage here
    }
}


# --- NOUVEAU ---
# Fonction pour remplir les ComboBox des OUs.
function Populate-OUComboBoxes {
    if (-not $script:OUList) {
        try {
            # Stocke la liste dans une variable de scope script pour ne pas la re-requêter à chaque fois
            $script:OUList = Get-ADOrganizationalUnit -Filter * | Select-Object DistinguishedName, Name | Sort-Object Name
            Write-Log (Get-LocalizedText "LogOURefreshed")
        } catch {
            Write-Log (Get-LocalizedText "LogErrorOURefresh" $_.Exception.Message) "ERREUR"
            return
        }
    }

    # Nettoyage des listes
    $createOUParentPathComboBox.Items.Clear()
    $deleteOUComboBox.Items.Clear()

    # Remplissage de la liste de création
    $createOUParentPathComboBox.Items.Add($domainPath) # Ajoute la racine du domaine
    foreach ($ou in $script:OUList) {
        [void]$createOUParentPathComboBox.Items.Add($ou.DistinguishedName)
    }

    # Remplissage de la liste de suppression
    foreach ($ou in $script:OUList) {
        [void]$deleteOUComboBox.Items.Add($ou.DistinguishedName)
    }
}

# Fonction globale pour remplir TOUTES les listes déroulantes de l'interface.
function Refresh-GuiLists {
    Write-Log (Get-LocalizedText "LogRefreshLists")
    $script:OUList = $null # Forcer le rechargement des OUs

    # --- Utilisateurs ---
    $deleteUserLoginComboBox.Items.Clear()
    $addUserToGroupLoginComboBox.Items.Clear()
    try {
        $adUsers = Get-ADUser -Filter * -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName | Sort-Object
        foreach ($userLogin in $adUsers) {
            [void]$deleteUserLoginComboBox.Items.Add($userLogin)
            [void]$addUserToGroupLoginComboBox.Items.Add($userLogin)
        }
        Write-Log (Get-LocalizedText "LogUsersRefreshed")
    } catch { Write-Log (Get-LocalizedText "LogErrorUsersRefresh" $_.Exception.Message) "ERREUR" }

    # --- Groupes ---
    $deleteGroupNameComboBox.Items.Clear()
    $addUserToGroupGroupNameComboBox.Items.Clear()
    try {
        $adGroups = Get-ADGroup -Filter * -Properties Name | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($groupName in $adGroups) {
            [void]$deleteGroupNameComboBox.Items.Add($groupName)
            [void]$addUserToGroupGroupNameComboBox.Items.Add($groupName)
        }
        Write-Log (Get-LocalizedText "LogGroupsRefreshed")
    } catch { Write-Log (Get-LocalizedText "LogErrorGroupsRefresh" $_.Exception.Message) "ERREUR" }

    # --- OUs ---
    Populate-OUComboBoxes
}


# --- NOUVEAU ---
# Fonction pour supprimer une Unité d'Organisation.
function Remove-OrganizationalUnit {
    $ouDN = $deleteOUComboBox.SelectedItem
    if (-not $ouDN) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorSelectOUToDelete"), "Erreur", "OK", "Error")
        return
    }

    $isRecursive = $deleteOURecursiveCheckBox.Checked
    $ouName = ($ouDN -split ",*..=")[1] # Extrait le nom simple pour le message
    
    $confirmMessage = if ($isRecursive) {
        (Get-LocalizedText "MsgConfirmDeleteOURecursive" $ouName)
    } else {
        (Get-LocalizedText "MsgConfirmDeleteOU" $ouName)
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMessage, (Get-LocalizedText "ConfirmDeletionTitle"), "YesNo", "Warning")
    if ($confirm -eq 'Yes') {
        try {
            if ($isRecursive) {
                # Suppression Récursive
                Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
            } else {
                # Suppression simple
                Remove-ADOrganizationalUnit -Identity $ouDN -Confirm:$false -ErrorAction Stop
            }
            Write-Log (Get-LocalizedText "LogOUDeleted" $ouName)
            $deleteOUComboBox.SelectedIndex = -1
            Refresh-GuiLists # Rafraîchir toutes les listes
        }
        catch {
            Write-Log (Get-LocalizedText "LogErrorOUDeletion" $ouName $_.Exception.Message) "ERREUR"
        }
    }
}



# Fonction pour mettre à jour les textes de l'interface graphique en fonction de la langue sélectionnée.
function Update-GuiLanguage {
    # Met à jour le texte de la fenêtre principale.
    $mainForm.Text = (Get-LocalizedText "MainWindowTitle")
    # Met à jour les titres des onglets.
    $tabPageImport.Text = (Get-LocalizedText "ImportTabTitle")
    $tabPageUsers.Text = (Get-LocalizedText "UsersTabTitle")
    $tabPageGroups.Text = (Get-LocalizedText "GroupsTabTitle")
    $tabPageOU.Text = (Get-LocalizedText "OUTabTitle") # --- NOUVEAU ---
    $tabPageInfo.Text = (Get-LocalizedText "InformationTabTitle")
    
    # Met à jour les textes des contrôles de l'onglet Import CSV.
    $importLabel.Text = (Get-LocalizedText "ImportLabelPath")
    $browseButton.Text = (Get-LocalizedText "BrowseButton")
    $importButton.Text = (Get-LocalizedText "ImportButton")

    # Met à jour les textes des contrôles de l'onglet Gestion Utilisateurs.
    $addUserGroupBox.Text = (Get-LocalizedText "AddUserGroupBox")
    $addUserPrenomLabel.Text = (Get-LocalizedText "AddUserFirstNameLabel")
    $addUserNomLabel.Text = (Get-LocalizedText "AddUserLastNameLabel")
    $addUserFonctionLabel.Text = (Get-LocalizedText "AddUserFunctionLabel")
    $addUserButton.Text = (Get-LocalizedText "AddUserButton")
    $deleteUserGroupBox.Text = (Get-LocalizedText "DeleteUserGroupBox")
    $deleteUserLoginLabel.Text = (Get-LocalizedText "DeleteUserSelectUserLabel")
    $deleteUserButton.Text = (Get-LocalizedText "DeleteUserButton")

    # Met à jour les textes des contrôles de l'onglet Gestion Groupes.
    $createGroupGroupBox.Text = (Get-LocalizedText "CreateGroupGroupBox")
    $createGroupNameLabel.Text = (Get-LocalizedText "CreateGroupNameLabel")
    $createGroupButton.Text = (Get-LocalizedText "CreateGroupButton")
    $deleteGroupGroupBox.Text = (Get-LocalizedText "DeleteGroupGroupBox")
    $deleteGroupNameLabel.Text = (Get-LocalizedText "DeleteGroupSelectGroupLabel")
    $deleteGroupButton.Text = (Get-LocalizedText "DeleteGroupButton")
    $addUserToGroupBox.Text = (Get-LocalizedText "AddUserToGroupGroupBox")
    $addUserToGroupLoginLabel.Text = (Get-LocalizedText "AddUserToGroupSelectUserLabel")
    $addUserToGroupGroupNameLabel.Text = (Get-LocalizedText "AddUserToGroupSelectGroupLabel")
    $addUserToGroupButton.Text = (Get-LocalizedText "AddUserToGroupButton")

    # Met à jour les textes des contrôles des logs et de la langue.
    $outputLabel.Text = (Get-LocalizedText "LogsLabel")
    $languageLabel.Text = (Get-LocalizedText "LanguageLabel")

    # --- NOUVEAU : Onglet OU ---
    $createOUGroupBox.Text = (Get-LocalizedText "CreateOUGroupBox")
    $createOUNameLabel.Text = (Get-LocalizedText "CreateOUNameLabel")
    $createOUParentLabel.Text = (Get-LocalizedText "CreateOUParentLabel")
    $createOUButton.Text = (Get-LocalizedText "CreateOUButton")
    $deleteOUGroupBox.Text = (Get-LocalizedText "DeleteOUGroupBox")
    $deleteOUSelectLabel.Text = (Get-LocalizedText "DeleteOUSelectLabel")
    $deleteOURecursiveCheckBox.Text = (Get-LocalizedText "DeleteOURecursiveLabel")
    $deleteOUButton.Text = (Get-LocalizedText "DeleteOUButton")    

    # Mise à jour du texte dans l'onglet Information
    $infoTextBox.Text = @"
### Active Directory Administration Tool ###

Version : 2.9
Auteur : Kiraly Geddon aide par Gemini

Description :
Ce script fournit une interface graphique (GUI) pour effectuer des operations
courantes dans Active Directory : import en masse, ajout/suppression d'utilisateurs,
gestion de groupes et gestion des Unités d'Organisation (OU).

Nouveautes de la v2.0 :
  - Detection automatique du domaine Active Directory.
  - Ajout de commentaires detailles pour chaque partie du code.

Nouveautes de la v2.5 :
  - Verification des droits d'administrateur et auto-elevation si necessaire.
  - Verification de la presence des outils RSAT pour AD et proposition
    d'installation automatique si manquants.

Nouveautes pour v2.6 :
  - Remplacement des TextBox par des ComboBox pour la selection d'utilisateurs/groupes existants.
  - Correction de l'erreur "Multiple ambiguous overloads found for "Font".
  - Ajout de la fonction Populate-ADDropdowns pour remplir les ComboBox.
  - Rafraîchissement automatique des ComboBox lors de la selection des onglets pertinents.

Nouveautes pour v2.7 :
  - Ajout d'une gestion multilingue (Français/Anglais) avec selection via ComboBox.
  - Toutes les chaînes de l'interface et des logs sont externalisees pour la traduction.

Nouveautés v2.8 :
  - Corrections de quelques erreurs à l'import CSV.
  - Correction de l'erreur de formatage des chaînes de caractères.

Nouveautés v2.9 :
  - Ajout d'un onglet pour la création et la suppression des Unités d'Organisation (OU).
  - Amélioration du rafraîchissement des listes.
  - Ajout d'une option de suppression récursive pour les OUs.
"@
}
#endregion

#region Definition de l'Interface Graphique (GUI)

# Crée la fenêtre principale.
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = (Get-LocalizedText "MainWindowTitle" -Args $domainSuffix) # Titre de la fenêtre, incluant le suffixe du domaine.
$mainForm.Size = New-Object System.Drawing.Size(650, 680) # Taille de la fenêtre ajustée pour le nouvel onglet.
$mainForm.StartPosition = 'CenterScreen' # Centre la fenêtre à l'écran.
$mainForm.FormBorderStyle = 'FixedDialog' # Empêche le redimensionnement de la fenêtre.
$mainForm.MaximizeBox = $false # Désactive le bouton de maximisation.

# Crée le contrôle TabControl (les onglets).
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(615, 350)
$mainForm.Controls.Add($tabControl)

# --- Onglet 1: Import CSV ---
$tabPageImport = New-Object System.Windows.Forms.TabPage
$tabPageImport.Text = (Get-LocalizedText "ImportTabTitle")
$tabControl.TabPages.Add($tabPageImport)

$importLabel = New-Object System.Windows.Forms.Label
$importLabel.Text = (Get-LocalizedText "ImportLabelPath")
$importLabel.Location = New-Object System.Drawing.Point(20, 30)
$importLabel.AutoSize = $true
$tabPageImport.Controls.Add($importLabel)

$csvPathTextBox = New-Object System.Windows.Forms.TextBox
$csvPathTextBox.Location = New-Object System.Drawing.Point(20, 55)
$csvPathTextBox.Size = New-Object System.Drawing.Size(450, 20)
$tabPageImport.Controls.Add($csvPathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = (Get-LocalizedText "BrowseButton")
$browseButton.Location = New-Object System.Drawing.Point(480, 53)
$browseButton.Size = New-Object System.Drawing.Size(100, 25)
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Fichiers CSV (*.csv)|*.csv" # Filtre pour n'afficher que les fichiers CSV.
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $csvPathTextBox.Text = $openFileDialog.FileName # Affiche le chemin du fichier sélectionné.
    }
})
$tabPageImport.Controls.Add($browseButton)

$importButton = New-Object System.Windows.Forms.Button
$importButton.Text = (Get-LocalizedText "ImportButton")
$importButton.Location = New-Object System.Drawing.Point(230, 120)
$importButton.Size = New-Object System.Drawing.Size(150, 40)
$importButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) # Définit la police et le style.
$importButton.Add_Click({ Start-CsvUserImport }) # Associe le clic du bouton à la fonction d'importation.
$tabPageImport.Controls.Add($importButton)

# --- Onglet 2: Gestion Utilisateurs ---
$tabPageUsers = New-Object System.Windows.Forms.TabPage
$tabPageUsers.Text = (Get-LocalizedText "UsersTabTitle")
$tabControl.TabPages.Add($tabPageUsers)

$addUserGroupBox = New-Object System.Windows.Forms.GroupBox
$addUserGroupBox.Text = (Get-LocalizedText "AddUserGroupBox")
$addUserGroupBox.Location = New-Object System.Drawing.Point(20, 20)
$addUserGroupBox.Size = New-Object System.Drawing.Size(270, 280)
$tabPageUsers.Controls.Add($addUserGroupBox)

$addUserPrenomLabel = New-Object System.Windows.Forms.Label; $addUserPrenomLabel.Text = (Get-LocalizedText "AddUserFirstNameLabel"); $addUserPrenomLabel.Location = New-Object System.Drawing.Point(15, 40); $addUserPrenomLabel.AutoSize = $true; $addUserGroupBox.Controls.Add($addUserPrenomLabel)
$addUserPrenomTextBox = New-Object System.Windows.Forms.TextBox; $addUserPrenomTextBox.Location = New-Object System.Drawing.Point(100, 37); $addUserPrenomTextBox.Size = New-Object System.Drawing.Size(150, 20); $addUserGroupBox.Controls.Add($addUserPrenomTextBox)

$addUserNomLabel = New-Object System.Windows.Forms.Label; $addUserNomLabel.Text = (Get-LocalizedText "AddUserLastNameLabel"); $addUserNomLabel.Location = New-Object System.Drawing.Point(15, 80); $addUserNomLabel.AutoSize = $true; $addUserGroupBox.Controls.Add($addUserNomLabel)
$addUserNomTextBox = New-Object System.Windows.Forms.TextBox; $addUserNomTextBox.Location = New-Object System.Drawing.Point(100, 77); $addUserNomTextBox.Size = New-Object System.Drawing.Size(150, 20); $addUserGroupBox.Controls.Add($addUserNomTextBox)

$addUserFonctionLabel = New-Object System.Windows.Forms.Label; $addUserFonctionLabel.Text = (Get-LocalizedText "AddUserFunctionLabel"); $addUserFonctionLabel.Location = New-Object System.Drawing.Point(15, 120); $addUserFonctionLabel.AutoSize = $true; $addUserGroupBox.Controls.Add($addUserFonctionLabel)
$addUserFonctionTextBox = New-Object System.Windows.Forms.TextBox; $addUserFonctionTextBox.Location = New-Object System.Drawing.Point(100, 117); $addUserFonctionTextBox.Size = New-Object System.Drawing.Size(150, 20); $addUserGroupBox.Controls.Add($addUserFonctionTextBox)

$addUserButton = New-Object System.Windows.Forms.Button; $addUserButton.Text = (Get-LocalizedText "AddUserButton"); $addUserButton.Location = New-Object System.Drawing.Point(60, 180); $addUserButton.Size = New-Object System.Drawing.Size(150, 30); $addUserButton.Add_Click({ Add-SingleUser }); $addUserGroupBox.Controls.Add($addUserButton)

$deleteUserGroupBox = New-Object System.Windows.Forms.GroupBox
$deleteUserGroupBox.Text = (Get-LocalizedText "DeleteUserGroupBox")
$deleteUserGroupBox.Location = New-Object System.Drawing.Point(310, 20)
$deleteUserGroupBox.Size = New-Object System.Drawing.Size(285, 280)
$tabPageUsers.Controls.Add($deleteUserGroupBox)

$deleteUserLoginLabel = New-Object System.Windows.Forms.Label; $deleteUserLoginLabel.Text = (Get-LocalizedText "DeleteUserSelectUserLabel"); $deleteUserLoginLabel.Location = New-Object System.Drawing.Point(15, 40); $deleteUserLoginLabel.AutoSize = $true; $deleteUserGroupBox.Controls.Add($deleteUserLoginLabel)

$deleteUserLoginComboBox = New-Object System.Windows.Forms.ComboBox
$deleteUserLoginComboBox.Location = New-Object System.Drawing.Point(15, 65)
$deleteUserLoginComboBox.Size = New-Object System.Drawing.Size(250, 20)
$deleteUserLoginComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList # Empêche l'utilisateur de taper, il doit choisir.
$deleteUserGroupBox.Controls.Add($deleteUserLoginComboBox)

$deleteUserButton = New-Object System.Windows.Forms.Button; $deleteUserButton.Text = (Get-LocalizedText "DeleteUserButton"); $deleteUserButton.Location = New-Object System.Drawing.Point(65, 180); $deleteUserButton.Size = New-Object System.Drawing.Size(150, 30); $deleteUserButton.BackColor = "MistyRose"; $deleteUserButton.Add_Click({ Remove-SpecificUser }); $deleteUserGroupBox.Controls.Add($deleteUserButton)

# --- Onglet 3: Gestion Groupes ---
$tabPageGroups = New-Object System.Windows.Forms.TabPage
$tabPageGroups.Text = (Get-LocalizedText "GroupsTabTitle")
$tabControl.TabPages.Add($tabPageGroups)

$createGroupGroupBox = New-Object System.Windows.Forms.GroupBox; $createGroupGroupBox.Text = (Get-LocalizedText "CreateGroupGroupBox"); $createGroupGroupBox.Location = New-Object System.Drawing.Point(20, 20); $createGroupGroupBox.Size = New-Object System.Drawing.Size(280, 130); $tabPageGroups.Controls.Add($createGroupGroupBox)
$createGroupNameLabel = New-Object System.Windows.Forms.Label; $createGroupNameLabel.Text = (Get-LocalizedText "CreateGroupNameLabel"); $createGroupNameLabel.Location = New-Object System.Drawing.Point(15, 30); $createGroupNameLabel.AutoSize = $true; $createGroupGroupBox.Controls.Add($createGroupNameLabel)
$createGroupNameTextBox = New-Object System.Windows.Forms.TextBox; $createGroupNameTextBox.Location = New-Object System.Drawing.Point(15, 55); $createGroupNameTextBox.Size = New-Object System.Drawing.Size(250, 20); $createGroupGroupBox.Controls.Add($createGroupNameTextBox)
$createGroupButton = New-Object System.Windows.Forms.Button; $createGroupButton.Text = (Get-LocalizedText "CreateGroupButton"); $createGroupButton.Location = New-Object System.Drawing.Point(90, 85); $createGroupButton.Size = New-Object System.Drawing.Size(100, 25); $createGroupButton.Add_Click({ Add-SpecificGroup }); $createGroupGroupBox.Controls.Add($createGroupButton)

$deleteGroupGroupBox = New-Object System.Windows.Forms.GroupBox; $deleteGroupGroupBox.Text = (Get-LocalizedText "DeleteGroupGroupBox"); $deleteGroupGroupBox.Location = New-Object System.Drawing.Point(315, 20); $deleteGroupGroupBox.Size = New-Object System.Drawing.Size(280, 130); $tabPageGroups.Controls.Add($deleteGroupGroupBox)
$deleteGroupNameLabel = New-Object System.Windows.Forms.Label; $deleteGroupNameLabel.Text = (Get-LocalizedText "DeleteGroupSelectGroupLabel"); $deleteGroupNameLabel.Location = New-Object System.Drawing.Point(15, 30); $deleteGroupNameLabel.AutoSize = $true; $deleteGroupGroupBox.Controls.Add($deleteGroupNameLabel)

$deleteGroupNameComboBox = New-Object System.Windows.Forms.ComboBox
$deleteGroupNameComboBox.Location = New-Object System.Drawing.Point(15, 55)
$deleteGroupNameComboBox.Size = New-Object System.Drawing.Size(250, 20)
$deleteGroupNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$deleteGroupGroupBox.Controls.Add($deleteGroupNameComboBox)

$deleteGroupButton = New-Object System.Windows.Forms.Button; $deleteGroupButton.Text = (Get-LocalizedText "DeleteGroupButton"); $deleteGroupButton.Location = New-Object System.Drawing.Point(90, 85); $deleteGroupButton.Size = New-Object System.Drawing.Size(100, 25); $deleteGroupButton.BackColor = "MistyRose"; $deleteGroupButton.Add_Click({ Remove-SpecificGroup }); $deleteGroupGroupBox.Controls.Add($deleteGroupButton)

$addUserToGroupBox = New-Object System.Windows.Forms.GroupBox; $addUserToGroupBox.Text = (Get-LocalizedText "AddUserToGroupGroupBox"); $addUserToGroupBox.Location = New-Object System.Drawing.Point(20, 170); $addUserToGroupBox.Size = New-Object System.Drawing.Size(575, 140); $tabPageGroups.Controls.Add($addUserToGroupBox)
$addUserToGroupLoginLabel = New-Object System.Windows.Forms.Label; $addUserToGroupLoginLabel.Text = (Get-LocalizedText "AddUserToGroupSelectUserLabel"); $addUserToGroupLoginLabel.Location = New-Object System.Drawing.Point(15, 30); $addUserToGroupLoginLabel.AutoSize = $true; $addUserToGroupBox.Controls.Add($addUserToGroupLoginLabel)

$addUserToGroupLoginComboBox = New-Object System.Windows.Forms.ComboBox
$addUserToGroupLoginComboBox.Location = New-Object System.Drawing.Point(15, 55)
$addUserToGroupLoginComboBox.Size = New-Object System.Drawing.Size(250, 20)
$addUserToGroupLoginComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$addUserToGroupBox.Controls.Add($addUserToGroupLoginComboBox)

$addUserToGroupGroupNameLabel = New-Object System.Windows.Forms.Label; $addUserToGroupGroupNameLabel.Text = (Get-LocalizedText "AddUserToGroupSelectGroupLabel"); $addUserToGroupGroupNameLabel.Location = New-Object System.Drawing.Point(300, 30); $addUserToGroupGroupNameLabel.AutoSize = $true; $addUserToGroupBox.Controls.Add($addUserToGroupGroupNameLabel)

$addUserToGroupGroupNameComboBox = New-Object System.Windows.Forms.ComboBox
$addUserToGroupGroupNameComboBox.Location = New-Object System.Drawing.Point(300, 55)
$addUserToGroupGroupNameComboBox.Size = New-Object System.Drawing.Size(250, 20)
$addUserToGroupGroupNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$addUserToGroupBox.Controls.Add($addUserToGroupGroupNameComboBox)

$addUserToGroupButton = New-Object System.Windows.Forms.Button; $addUserToGroupButton.Text = (Get-LocalizedText "AddUserToGroupButton"); $addUserToGroupButton.Location = New-Object System.Drawing.Point(212, 95); $addUserToGroupButton.Size = New-Object System.Drawing.Size(150, 30); $addUserToGroupButton.Add_Click({ Add-UserToSpecificGroup }); $addUserToGroupBox.Controls.Add($addUserToGroupButton)

# --- NOUVEAU : Onglet 4: Gestion des OUs ---
$tabPageOU = New-Object System.Windows.Forms.TabPage
$tabPageOU.Text = (Get-LocalizedText "OUTabTitle")
$tabControl.TabPages.Add($tabPageOU)

# GroupBox pour la création d'OU
$createOUGroupBox = New-Object System.Windows.Forms.GroupBox; $createOUGroupBox.Text = (Get-LocalizedText "CreateOUGroupBox"); $createOUGroupBox.Location = New-Object System.Drawing.Point(20, 20); $createOUGroupBox.Size = New-Object System.Drawing.Size(575, 130); $tabPageOU.Controls.Add($createOUGroupBox)
$createOUNameLabel = New-Object System.Windows.Forms.Label; $createOUNameLabel.Text = (Get-LocalizedText "CreateOUNameLabel"); $createOUNameLabel.Location = New-Object System.Drawing.Point(15, 30); $createOUNameLabel.AutoSize = $true; $createOUGroupBox.Controls.Add($createOUNameLabel)
$createOUNameTextBox = New-Object System.Windows.Forms.TextBox; $createOUNameTextBox.Location = New-Object System.Drawing.Point(180, 27); $createOUNameTextBox.Size = New-Object System.Drawing.Size(250, 20); $createOUGroupBox.Controls.Add($createOUNameTextBox)

$createOUParentLabel = New-Object System.Windows.Forms.Label; $createOUParentLabel.Text = (Get-LocalizedText "CreateOUParentLabel"); $createOUParentLabel.Location = New-Object System.Drawing.Point(15, 60); $createOUParentLabel.AutoSize = $true; $createOUGroupBox.Controls.Add($createOUParentLabel)
$createOUParentPathComboBox = New-Object System.Windows.Forms.ComboBox; $createOUParentPathComboBox.Location = New-Object System.Drawing.Point(180, 57); $createOUParentPathComboBox.Size = New-Object System.Drawing.Size(380, 20); $createOUParentPathComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $createOUGroupBox.Controls.Add($createOUParentPathComboBox)

$createOUButton = New-Object System.Windows.Forms.Button; $createOUButton.Text = (Get-LocalizedText "CreateOUButton"); $createOUButton.Location = New-Object System.Drawing.Point(220, 90); $createOUButton.Size = New-Object System.Drawing.Size(150, 30); $createOUButton.Add_Click({ Add-OrganizationalUnit }); $createOUGroupBox.Controls.Add($createOUButton)


# GroupBox pour la suppression d'OU
$deleteOUGroupBox = New-Object System.Windows.Forms.GroupBox; $deleteOUGroupBox.Text = (Get-LocalizedText "DeleteOUGroupBox"); $deleteOUGroupBox.Location = New-Object System.Drawing.Point(20, 160); $deleteOUGroupBox.Size = New-Object System.Drawing.Size(575, 150); $tabPageOU.Controls.Add($deleteOUGroupBox)
$deleteOUSelectLabel = New-Object System.Windows.Forms.Label; $deleteOUSelectLabel.Text = (Get-LocalizedText "DeleteOUSelectLabel"); $deleteOUSelectLabel.Location = New-Object System.Drawing.Point(15, 30); $deleteOUSelectLabel.AutoSize = $true; $deleteOUGroupBox.Controls.Add($deleteOUSelectLabel)
$deleteOUComboBox = New-Object System.Windows.Forms.ComboBox; $deleteOUComboBox.Location = New-Object System.Drawing.Point(15, 50); $deleteOUComboBox.Size = New-Object System.Drawing.Size(545, 20); $deleteOUComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $deleteOUGroupBox.Controls.Add($deleteOUComboBox)

$deleteOURecursiveCheckBox = New-Object System.Windows.Forms.CheckBox; $deleteOURecursiveCheckBox.Text = (Get-LocalizedText "DeleteOURecursiveLabel"); $deleteOURecursiveCheckBox.Location = New-Object System.Drawing.Point(15, 80); $deleteOURecursiveCheckBox.AutoSize = $true; $deleteOUGroupBox.Controls.Add($deleteOURecursiveCheckBox)

$deleteOUButton = New-Object System.Windows.Forms.Button; $deleteOUButton.Text = (Get-LocalizedText "DeleteOUButton"); $deleteOUButton.Location = New-Object System.Drawing.Point(220, 105); $deleteOUButton.Size = New-Object System.Drawing.Size(150, 30); $deleteOUButton.BackColor = "MistyRose"; $deleteOUButton.Add_Click({ Remove-OrganizationalUnit }); $deleteOUGroupBox.Controls.Add($deleteOUButton)


# --- Onglet 5: Information --- 
$tabPageInfo = New-Object System.Windows.Forms.TabPage
$tabPageInfo.Text = (Get-LocalizedText "InformationTabTitle") # Utilise la chaîne localisée pour le titre de l'onglet
$tabControl.TabPages.Add($tabPageInfo)

# Zone de texte pour afficher les informations dans l'onglet Information
$infoTextBox = New-Object System.Windows.Forms.TextBox
$infoTextBox.Location = New-Object System.Drawing.Point(10, 10)
$infoTextBox.Size = New-Object System.Drawing.Size(585, 300) # Taille ajustée pour le contenu
$infoTextBox.Multiline = $true
$infoTextBox.ScrollBars = "Vertical" # Ajoute des barres de défilement si le texte est trop long
$infoTextBox.ReadOnly = $true # Rend la zone de texte non éditable
$infoTextBox.Font = New-Object System.Drawing.Font("Consolas", 9) # Police pour une meilleure lisibilité du code/texte
$tabPageInfo.Controls.Add($infoTextBox)


# Zone de logs (en bas de la fenêtre principale)
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = (Get-LocalizedText "LogsLabel")
$outputLabel.Location = New-Object System.Drawing.Point(10, 370)
$outputLabel.AutoSize = $true
$mainForm.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(10, 390)
$outputTextBox.Size = New-Object System.Drawing.Size(615, 170) # Taille ajustée pour accueillir le nouvel onglet sans trop réduire les logs
$outputTextBox.Multiline = $true
$outputTextBox.ScrollBars = "Vertical"
$outputTextBox.ReadOnly = $true
$outputTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
$mainForm.Controls.Add($outputTextBox)

# Contrôle de selection de la langue (en bas de la fenêtre principale)
$languageLabel = New-Object System.Windows.Forms.Label
$languageLabel.Text = (Get-LocalizedText "LanguageLabel")
$languageLabel.Location = New-Object System.Drawing.Point(10, 570) # Position ajustée
$languageLabel.AutoSize = $true
$mainForm.Controls.Add($languageLabel)

$languageComboBox = New-Object System.Windows.Forms.ComboBox
$languageComboBox.Location = New-Object System.Drawing.Point(80, 567) # Position ajustée
$languageComboBox.Size = New-Object System.Drawing.Size(100, 20)
$languageComboBox.Items.AddRange(@("fr", "en")) # Ajouter les langues disponibles
$languageComboBox.SelectedItem = $CurrentLanguage # Sélectionner la langue par défaut
$languageComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList # Empêche l'édition manuelle
$languageComboBox.Add_SelectedIndexChanged({
    $CurrentLanguage = $languageComboBox.SelectedItem # Mettre à jour la variable de langue globale
    Update-GuiLanguage # Appeler la fonction pour rafraîchir tous les textes de l'interface
})
$mainForm.Controls.Add($languageComboBox)


#endregion

#region Affichage de la fenêtre

# Événement déclenché lorsque la fenêtre principale est affichée pour la première fois.
$mainForm.Add_Shown({
    $mainForm.Activate()
    Update-GuiLanguage
    Refresh-GuiLists # Remplir toutes les listes au démarrage.
})

# Événement déclenché lorsque l'onglet sélectionné change.
$tabControl.Add_SelectedIndexChanged({
    # Rafraîchit les listes uniquement pour l'onglet pertinent pour éviter les appels inutiles.
    if ($tabControl.SelectedTab -eq $tabPageUsers -or $tabControl.SelectedTab -eq $tabPageGroups) {
        Refresh-GuiLists
    }
    if ($tabControl.SelectedTab -eq $tabPageOU) {
        Populate-OUComboBoxes # Un rafraîchissement plus léger suffit ici.
    }
})

# Affiche la fenêtre principale.
[void]$mainForm.ShowDialog()
#endregion