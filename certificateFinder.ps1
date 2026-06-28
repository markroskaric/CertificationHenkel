Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Set-Win11Style {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Control,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Form','Label','TextBox','Button','Log')]
        [string]$Type
    )

    switch ($Type) {
        'Form' {
            $Control.BackColor = [System.Drawing.Color]::FromArgb(248,250,252)
            $Control.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $Control.Padding = New-Object System.Windows.Forms.Padding(12)
        }
        'Label' {
            $Control.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $Control.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
            $Control.AutoSize = $true
        }
        'TextBox' {
            $Control.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $Control.BackColor = [System.Drawing.Color]::White
            $Control.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
        }
        'Button' {
            $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $Control.FlatAppearance.BorderSize = 0
            $Control.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $Control.BackColor = [System.Drawing.Color]::FromArgb(0,120,212)
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.UseVisualStyleBackColor = $false
            $Control.Cursor = [System.Windows.Forms.Cursors]::Hand
        }
        'Log' {
            $Control.Font = New-Object System.Drawing.Font('Consolas', 9)
            $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $Control.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
            $Control.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
            $Control.ReadOnly = $true
        }
    }
}

# =========================
# FORM
# =========================

$form = New-Object System.Windows.Forms.Form
$form.Text = "COA Certificate Finder"
$form.Size = New-Object System.Drawing.Size(750,480)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(750,480)
Set-Win11Style -Control $form -Type 'Form'

function Create-Field {
    param ($labelText, $yPos)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(10,$yPos)
    $label.Size = New-Object System.Drawing.Size(160,22)
    Set-Win11Style -Control $label -Type 'Label'

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(180,$yPos)
    $textbox.Size = New-Object System.Drawing.Size(400,28)
    Set-Win11Style -Control $textbox -Type 'TextBox'

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Browse"
    $button.Location = New-Object System.Drawing.Point(600,$yPos)
    $button.Size = New-Object System.Drawing.Size(100,30)
    Set-Win11Style -Control $button -Type 'Button'

    $form.Controls.Add($label)
    $form.Controls.Add($textbox)
    $form.Controls.Add($button)

    return @{ TextBox=$textbox; Button=$button }
}

$txtField = Create-Field "TXT file:" 20
$sourceField = Create-Field "Source folder:" 60
$targetField = Create-Field "Target folder:" 100

# START BUTTON
$startBtn = New-Object System.Windows.Forms.Button
$startBtn.Text = "START"
$startBtn.Location = New-Object System.Drawing.Point(300,140)
$startBtn.Size = New-Object System.Drawing.Size(140,38)
Set-Win11Style -Control $startBtn -Type 'Button'
$form.Controls.Add($startBtn)

# LOG
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Location = New-Object System.Drawing.Point(10,190)
$logBox.Size = New-Object System.Drawing.Size(710,240)
Set-Win11Style -Control $logBox -Type 'Log'
$form.Controls.Add($logBox)

function Log($msg) {
    $logBox.AppendText($msg + "`r`n")
    $logBox.Refresh()
}

# =========================
# BROWSE
# =========================

$txtField.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    if ($dialog.ShowDialog() -eq "OK") {
        $txtField.TextBox.Text = $dialog.FileName
    }
})

$sourceField.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq "OK") {
        $sourceField.TextBox.Text = $dialog.SelectedPath
    }
})

$targetField.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq "OK") {
        $targetField.TextBox.Text = $dialog.SelectedPath
    }
})

# =========================
# MAIN
# =========================

$startBtn.Add_Click({

try {

    $startBtn.Enabled = $false
    $logBox.Clear()

    $idhFile = $txtField.TextBox.Text
    $sourceFolder = $sourceField.TextBox.Text
    $targetFolder = $targetField.TextBox.Text

    if ([string]::IsNullOrWhiteSpace($idhFile) -eq $true -or !(Test-Path $idhFile)) {
        Log "ERROR: TXT file missing"
        return
    }

    if (!(Test-Path $sourceFolder)) {
        Log "ERROR: Source folder missing"
        return
    }

    if (!(Test-Path $targetFolder)) {
        New-Item -ItemType Directory -Path $targetFolder | Out-Null
    }

    $lines = Get-Content $idhFile

    Log "Indexing PDFs..."
    $allPdf = Get-ChildItem $sourceFolder -Recurse -Filter *.pdf -File

    $results = @()

    foreach ($line in $lines) {

        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # separator
        if ($line -match "`t") {
            $parts = $line -split "`t"
        } elseif ($line -match ";") {
            $parts = $line -split ";"
        } else {
            $parts = $line -split "\s+"
        }

        if ($parts.Count -lt 2) {
            Log "Invalid: $line"
            continue
        }

        $idh = $parts[0].Trim()
        $batch = $parts[1].Trim()

        if ($parts.Count -ge 3) {
            if ($line -match "`t" -or $line -match ";") {
                $product = $parts[2].Trim()
            } else {
                $product = ($parts[2..($parts.Count-1)] -join " ").Trim()
            }
        } else {
            $product = ""
        }

        if ($idh -match "IDH") { continue }

        Log "Processing: $idh / $batch"

        $files = $allPdf | Where-Object {
            $_.Name -like "*$idh*" -and $_.Name -like "*$batch*"
        }

        if ($files.Count -gt 0) {

            $selected = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1

            Copy-Item $selected.FullName -Destination $targetFolder -Force

            $results += [PSCustomObject]@{
                IDH=$idh
                Batch=$batch
                Product=$product
                CopiedFile=$selected.Name
                CopiedPath=$selected.FullName
                CopiedDate=$selected.LastWriteTime
                Count=$files.Count
                AllFiles=($files.Name -join ", ")
                AllPaths=($files.FullName -join " | ")
                Missing=""
            }

            Log "FOUND ($($files.Count)) -> copied newest"
        }
        else {
            $results += [PSCustomObject]@{
                IDH=$idh
                Batch=$batch
                Product=$product
                CopiedFile=""
                CopiedPath=""
                CopiedDate=""
                Count=0
                AllFiles=""
                AllPaths=""
                Missing="YES"
            }

            Log "MISSING"
        }
    }

    Log "Creating Excel..."

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $wb = $excel.Workbooks.Add()

    $wsF = $wb.Worksheets.Item(1)
    $wsF.Name = "NAJDENE"

    $wsM = $wb.Worksheets.Add()
    $wsM.Name = "TREBA NAREDIT"

    $headersF = "IDH","Batch","Product","Copied file","Copied path","Copied date","Found count","All files","All paths"
    $headersM = "IDH","Batch","Product","SUD","Sudcharge"

    for ($i=0;$i -lt $headersF.Count;$i++){
        $wsF.Cells.Item(1,$i+1)=$headersF[$i]
    }

    for ($i=0;$i -lt $headersM.Count;$i++){
        $wsM.Cells.Item(1,$i+1)=$headersM[$i]
    }

    $rF=2
    $rM=2

    foreach ($r in $results){

        if ($r.Missing -eq ""){

            $wsF.Cells.Item($rF,1)=$r.IDH
            $wsF.Cells.Item($rF,2)=$r.Batch
            $wsF.Cells.Item($rF,3)=$r.Product
            $wsF.Cells.Item($rF,4)=$r.CopiedFile
            $wsF.Cells.Item($rF,5)=$r.CopiedPath
            $wsF.Cells.Item($rF,6)=$r.CopiedDate
            $wsF.Cells.Item($rF,7)=$r.Count
            $wsF.Cells.Item($rF,8)=$r.AllFiles
            $wsF.Cells.Item($rF,9)=$r.AllPaths
            $rF++
        }
        else {
            $wsM.Cells.Item($rM,1)=$r.IDH
            $wsM.Cells.Item($rM,2)=$r.Batch
            $wsM.Cells.Item($rM,3)=$r.Product
            $rM++
        }
    }

    $wsF.Columns.AutoFit()
    $wsM.Columns.AutoFit()

    $excelPath="$targetFolder\report.xlsx"
    $wb.SaveAs($excelPath)
    $wb.Close()
    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Log "DONE"
    Log $excelPath

    [System.Windows.Forms.MessageBox]::Show("Finished!","Done")

}
catch {
    Log $_
}
finally {
    $startBtn.Enabled = $true
}

})

$form.ShowDialog()
