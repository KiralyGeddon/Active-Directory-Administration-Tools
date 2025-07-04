#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Onglet Import CSV ###
#
# Description :
# Ce script definit l'interface et la logique de l'onglet "Import en Masse (CSV)".
# Il est charge par main.ps1.
#-----------------------------------------------------------------------------------

# Assurez-vous que les variables globales sont accessibles
# $script:tabControl, $script:CurrentLanguage, $script:Localization, Get-LocalizedText
# $script:domainPath, $script:domainSuffix, Write-Log, Populate-ADDropdowns

#region Fonctions Logiques de l'onglet Import CSV

function Start-CsvUserImport {
    $csvPath = $script:csvPathTextBox.Text
    if (-not (Test-Path $csvPath) -or ($csvPath -notlike "*.csv")) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedText "MsgErrorInvalidCsv"), "Erreur", "OK", "Error")
        return
    }

    try {
        $CSVData = Import-CSV -Path $csvPath -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Log (Get-LocalizedText "LogCsvImportStarted" $csvPath)
    }
    catch {
        Write-Log (Get-LocalizedText "MsgErrorCsvRead") "ERREUR"
        return
    }

    if (-not(Get-ADOrganizationalUnit -Filter "Name -eq 'Utilisateurs'" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name "Utilisateurs" -Path $script:domainPath
        Write-Log (Get-LocalizedText "LogOuCreated" "Utilisateurs")
    }

    foreach ($Utilisateur in $CSVData) {
        $prenom = $Utilisateur.Prenom
        $nom = $Utilisateur.Nom
        $fonction = $Utilisateur.Fonction.Trim()
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
        }
        else {
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
                Write-Log (Get-LocalizedText "LogUserCreated" $login $nom $prenom)
            }
            catch {
                Write-Log (Get-LocalizedText "LogErrorUserCreation" $login $_) "ERREUR"
            }
        }
    }
    Write-Log (Get-LocalizedText "LogImportFinished")
    Populate-ADDropdowns
}

#endregion

#region Définition de l'interface graphique de l'onglet Import CSV

$script:tabPageImport = New-Object System.Windows.Forms.TabPage
$script:tabPageImport.Text = (Get-LocalizedText "ImportTabTitle")
$script:tabControl.TabPages.Add($script:tabPageImport)

$script:importLabel = New-Object System.Windows.Forms.Label
$script:importLabel.Text = (Get-LocalizedText "ImportLabelPath")
$script:importLabel.Location = New-Object System.Drawing.Point(20, 30)
$script:importLabel.AutoSize = $true
$script:tabPageImport.Controls.Add($script:importLabel)

$script:csvPathTextBox = New-Object System.Windows.Forms.TextBox
$script:csvPathTextBox.Location = New-Object System.Drawing.Point(20, 55)
$script:csvPathTextBox.Size = New-Object System.Drawing.Size(450, 20)
$script:tabPageImport.Controls.Add($script:csvPathTextBox)

$script:browseButton = New-Object System.Windows.Forms.Button
$script:browseButton.Text = (Get-LocalizedText "BrowseButton")
$script:browseButton.Location = New-Object System.Drawing.Point(480, 53)
$script:browseButton.Size = New-Object System.Drawing.Size(100, 25)
$script:browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $script:csvPathTextBox.Text = $openFileDialog.FileName
    }
})
$script:tabPageImport.Controls.Add($script:browseButton)

$script:importButton = New-Object System.Windows.Forms.Button
$script:importButton.Text = (Get-LocalizedText "ImportButton")
$script:importButton.Location = New-Object System.Drawing.Point(230, 120)
$script:importButton.Size = New-Object System.Drawing.Size(150, 40)
$script:importButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:importButton.Add_Click({ Start-CsvUserImport })
$script:tabPageImport.Controls.Add($script:importButton)

#endregion