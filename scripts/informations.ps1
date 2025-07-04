#-----------------------------------------------------------------------------------
#               ### Active Directory Management Tool - Onglet Informations ###
#
# Description :
# Ce script definit l'interface et la logique de l'onglet "Informations".
# Il est charge par main.ps1.
#-----------------------------------------------------------------------------------

# Assurez-vous que les variables globales sont accessibles
# $script:tabControl, $script:CurrentLanguage, $script:Localization, Get-LocalizedText

#region Définition de l'onglet Informations

$script:tabPageInfo = New-Object System.Windows.Forms.TabPage
$script:tabPageInfo.Text = (Get-LocalizedText "InfoTabTitle")
$script:tabControl.TabPages.Add($script:tabPageInfo)

# Contenu de l'onglet Information
$script:infoLabelTitle = New-Object System.Windows.Forms.Label
$script:infoLabelTitle.Text = (Get-LocalizedText "InfoLabelTitle")
$script:infoLabelTitle.Location = New-Object System.Drawing.Point(50, 40)
$script:infoLabelTitle.AutoSize = $true
$script:infoLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$script:tabPageInfo.Controls.Add($script:infoLabelTitle)

$script:infoLabelVersion = New-Object System.Windows.Forms.Label
$script:infoLabelVersion.Text = (Get-LocalizedText "InfoLabelVersion")
$script:infoLabelVersion.Location = New-Object System.Drawing.Point(50, 90)
$script:infoLabelVersion.AutoSize = $true
$script:tabPageInfo.Controls.Add($script:infoLabelVersion)

$script:infoLabelAuthor = New-Object System.Windows.Forms.Label
$script:infoLabelAuthor.Text = (Get-LocalizedText "InfoLabelAuthor")
$script:infoLabelAuthor.Location = New-Object System.Drawing.Point(50, 120)
$script:infoLabelAuthor.AutoSize = $true
$script:tabPageInfo.Controls.Add($script:infoLabelAuthor)

$script:infoLabelDescription = New-Object System.Windows.Forms.Label
$script:infoLabelDescription.Text = (Get-LocalizedText "InfoLabelDescription")
$script:infoLabelDescription.Location = New-Object System.Drawing.Point(50, 150)
$script:infoLabelDescription.Size = New-Object System.Drawing.Size(500, 80)
$script:tabPageInfo.Controls.Add($script:infoLabelDescription)

#endregion