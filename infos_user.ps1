# ============================================================================
# Script : UserInfoRunspace.ps1
# Objectif : Interface graphique PowerShell avec Runspace pour logs AD
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globales
$global:isJobRunning = $false
$global:Runspace = $null
$global:PowerShellAsync = $null
$global:AsyncResult = $null

# --- Formulaire principal ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Recherche AD - Informations Utilisateur et Connexions"
$form.Width = 600
$form.Height = 850
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Zone de log
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = '20,710'
$logBox.Size = '540,100'
$logBox.ReadOnly = $true
$logBox.BackColor = 'WhiteSmoke'
$form.Controls.Add($logBox)

function Write-LogToBox {
    param([string]$msg)
    $form.Invoke([System.Action]{
        $logBox.AppendText("$(Get-Date -Format 'HH:mm:ss') - $msg`r`n")
        $logBox.ScrollToCaret()
    })
}

# Groupe de saisie
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Paramètres de recherche"
$groupBox.Size = '540,280'
$groupBox.Location = '20,20'
$form.Controls.Add($groupBox)

# Domaine
$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Text = "Domaine :"
$labelDomain.Location = '10,25'
$groupBox.Controls.Add($labelDomain)

$textBoxDomain = New-Object System.Windows.Forms.TextBox
$textBoxDomain.Location = '110,20'
$textBoxDomain.Width = 150
$groupBox.Controls.Add($textBoxDomain)
$textBoxDomain.Text = $env:USERDNSDOMAIN

# Utilisateur
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Utilisateur :"
$labelUser.Location = '10,55'
$groupBox.Controls.Add($labelUser)

$textBoxUser = New-Object System.Windows.Forms.TextBox
$textBoxUser.Location = '110,50'
$textBoxUser.Width = 150
$groupBox.Controls.Add($textBoxUser)

# Nombre de jours
$labelDays = New-Object System.Windows.Forms.Label
$labelDays.Text = "Nb jours :"
$labelDays.Location = '10,85'
$groupBox.Controls.Add($labelDays)

$numericUpDown = New-Object System.Windows.Forms.NumericUpDown
$numericUpDown.Location = '110,80'
$numericUpDown.Minimum = 1
$numericUpDown.Maximum = 30
$numericUpDown.Value = 7
$groupBox.Controls.Add($numericUpDown)

# DC
$labelDC = New-Object System.Windows.Forms.Label
$labelDC.Text = "Contrôleur de domaine :"
$labelDC.Location = '10,115'
$labelDC.Width = 150
$labelDC.AutoSize = $true
$groupBox.Controls.Add($labelDC)

$comboBoxPDC = New-Object System.Windows.Forms.ComboBox
$comboBoxPDC.Location = '10,135'
$comboBoxPDC.Width = 250
$comboBoxPDC.DropDownStyle = 'DropDownList'
$groupBox.Controls.Add($comboBoxPDC)

# Champs AD
$labelFields = New-Object System.Windows.Forms.Label
$labelFields.Text = "Champs AD :"
$labelFields.Location = '280,20'
$groupBox.Controls.Add($labelFields)

$checkListBox = New-Object System.Windows.Forms.CheckedListBox
$checkListBox.Location = '280,45'
$checkListBox.Size = '240,140'
$groupBox.Controls.Add($checkListBox)

# Boutons sélection champs
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Tout sélectionner"
$btnSelectAll.Location = '280,195'
$btnSelectAll.AutoSize = $true
$btnSelectAll.Width = 120
$btnSelectAll.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBox.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Tout désélectionner"
$btnDeselectAll.Location = '400,195'
$btnDeselectAll.AutoSize = $true
$btnDeselectAll.Width = 140
$btnDeselectAll.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBox.Controls.Add($btnDeselectAll)

# Bouton de recherche
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "Rechercher"
$buttonSearch.Location = '430,230'
$buttonSearch.Size = '90,25'
$groupBox.Controls.Add($buttonSearch)
# Barre de progression
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = '20,310'
$progressBar.Width = 540
$progressBar.Style = 'Marquee'
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Tableau de résultats
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = '20,340'
$dataGridView.Size = '540,360'
$dataGridView.ReadOnly = $true
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.RowHeadersVisible = $false
$dataGridView.SelectionMode = 'FullRowSelect'
$form.Controls.Add($dataGridView)

$col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col1.HeaderText = "Propriété"
$col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col2.HeaderText = "Valeur"
[void]$dataGridView.Columns.Add($col1)
[void]$dataGridView.Columns.Add($col2)

# === Ajout du menu contextuel pour copier et exporter ===
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$menuCopy = New-Object System.Windows.Forms.ToolStripMenuItem
$menuCopy.Text = "Copier"

$menuExport = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExport.Text = "Exporter données"

[void]$contextMenu.Items.Add($menuCopy)
[void]$contextMenu.Items.Add($menuExport)
$dataGridView.SelectionMode = 'CellSelect'
$dataGridView.ContextMenuStrip = $contextMenu

# Gestion du clic sur "Copier"
$menuCopy.Add_Click({
    if ($dataGridView.SelectedCells.Count -gt 0) {
        # On prend la première cellule sélectionnée
        $selectedCell = $dataGridView.SelectedCells[0]
        $cellText = if ($selectedCell.Value) { $selectedCell.Value.ToString() } else { "" }
        [System.Windows.Forms.Clipboard]::SetText($cellText)
        Write-LogToBox "Cellule copiée dans le presse-papier : $cellText"
    }
})

# Gestion du clic sur "Exporter"
$menuExport.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $sfd.Title = "Exporter les données vers un fichier CSV"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $sb = New-Object System.Text.StringBuilder
            # Ajout de l'en-tête (les noms de colonnes)
            $header = ($dataGridView.Columns | ForEach-Object { $_.HeaderText }) -join ","
            $sb.AppendLine($header) | Out-Null

            foreach ($row in $dataGridView.Rows) {
                if (-not $row.IsNewRow) {
                    $rowData = @()
                    foreach ($cell in $row.Cells) {
                        $value = $cell.Value
                        if ($value -eq $null) {
                            $value = ""
                        }
                        else {
                            $value = $value.ToString().Replace('"','""')
                            $value = '"{0}"' -f $value
                        }
                        $rowData += $value
                    }
                    $sb.AppendLine(($rowData -join ",")) | Out-Null
                }
            }
            [System.IO.File]::WriteAllText($sfd.FileName, $sb.ToString())
            Write-LogToBox "Données exportées vers $($sfd.FileName)"
        }
        catch {
            Write-LogToBox "Erreur lors de l'export : $_"
        }
    }
})

# === MAJ DC et champs ===
function UpdateDCAndADFields {
    param($Domain, $User)

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $dcs = Get-ADDomainController -Filter * -Server $Domain | Select-Object -ExpandProperty HostName
        $comboBoxPDC.Items.Clear()
        $dcs | ForEach-Object { [void]$comboBoxPDC.Items.Add($_) }
        if ($dcs.Count -gt 0) { $comboBoxPDC.SelectedIndex = 0 }

        # Champs utilisateur
        if (-not [string]::IsNullOrWhiteSpace($User)) {
            $userObj = Get-ADUser -Server $Domain -Identity $User -Properties * -ErrorAction Stop
            $checkListBox.Items.Clear()
            $props = $userObj.PSObject.Properties | Select-Object -ExpandProperty Name | Sort-Object
            $defaultFields = @("DisplayName","SamAccountName","mail","GivenName","sn")
            foreach ($p in $defaultFields + ($props | Where-Object { $_ -notin $defaultFields })) {
                $index = $checkListBox.Items.Add($p)
                if ($defaultFields -contains $p) { $checkListBox.SetItemChecked($index, $true) }
            }
        }
    } catch {
        Write-LogToBox "Erreur dans UpdateDC: $_"
        $comboBoxPDC.Items.Clear()
        $checkListBox.Items.Clear()
    }
}

# === Sélection / désélection de tous les champs AD ===
$btnSelectAll.Add_Click({
    $checkListBox.BeginUpdate()
    for ($i = 0; $i -lt $checkListBox.Items.Count; $i++) {
        $checkListBox.SetItemChecked($i, $true)
    }
    $checkListBox.EndUpdate()
})

$btnDeselectAll.Add_Click({
    $checkListBox.BeginUpdate()
    for ($i = 0; $i -lt $checkListBox.Items.Count; $i++) {
        $checkListBox.SetItemChecked($i, $false)
    }
    $checkListBox.EndUpdate()
})

# === MAJ automatique des DCs quand on appuie sur Entrée ===
$textBoxDomain.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        UpdateDCAndADFields -Domain $textBoxDomain.Text -User $textBoxUser.Text
    }
})
$textBoxUser.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        UpdateDCAndADFields -Domain $textBoxDomain.Text -User $textBoxUser.Text
    }
})

# === Clic sur "Rechercher" ===
$buttonSearch.Add_Click({
    $dataGridView.Rows.Clear()
    $logBox.Clear()

    $domain = $textBoxDomain.Text
    $user = $textBoxUser.Text
    $nb = [int]$numericUpDown.Value
    $dc = $comboBoxPDC.SelectedItem

    if (-not $domain -or -not $user) {
        [System.Windows.Forms.MessageBox]::Show("Renseignez domaine et utilisateur.","Attention",0,48)
        return
    }
    if (-not $dc) {
        [System.Windows.Forms.MessageBox]::Show("Sélectionnez un DC.","Information manquante",0,48)
        return
    }

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $userObj = Get-ADUser -Server $domain -Identity $user -Properties * -ErrorAction Stop

        # Affichage des champs AD sélectionnés
        foreach ($field in $checkListBox.CheckedItems) {
            $value = if ($userObj.PSObject.Properties[$field]) { $userObj.$field } else { "N/A" }
            $dataGridView.Rows.Add($field, $value)
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur : $($_.Exception.Message)", "Erreur", 0, 16)
        $progressBar.Visible = $false
        $form.Cursor = 'Default'
    }
})

# === Nettoyage à la fermeture ===
$form.Add_FormClosing({
    if ($global:PowerShellAsync) { $global:PowerShellAsync.Dispose() }
    if ($global:Runspace) { $global:Runspace.Close() }
})

# === Initialisation au démarrage ===
UpdateDCAndADFields -Domain $textBoxDomain.Text -User $textBoxUser.Text

# === Affichage du formulaire ===
[void]$form.ShowDialog()
