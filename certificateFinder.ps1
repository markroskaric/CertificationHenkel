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

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Location = New-Object System.Drawing.Point(180,$yPos)
    $pathLabel.Size = New-Object System.Drawing.Size(400,28)
    $pathLabel.Text = "Not selected"
    $pathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $pathLabel.AutoEllipsis = $true
    $pathLabel.AutoSize = $false
    $pathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pathLabel.BackColor = [System.Drawing.Color]::White
    $pathLabel.ForeColor = [System.Drawing.Color]::FromArgb(64,64,64)
    $pathLabel.Padding = New-Object System.Windows.Forms.Padding(4,0,4,0)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Browse"
    $button.Location = New-Object System.Drawing.Point(600,$yPos)
    $button.Size = New-Object System.Drawing.Size(100,30)
    Set-Win11Style -Control $button -Type 'Button'

    $form.Controls.Add($label)
    $form.Controls.Add($pathLabel)
    $form.Controls.Add($button)

    return [PSCustomObject]@{
        DisplayLabel = $pathLabel
        Button = $button
        Value = $null
    }
}

$certificationField = Create-Field "Certificates folder:" 60

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

$certificationField.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq "OK") {
        $certificationField.Value = $dialog.SelectedPath
        $certificationField.DisplayLabel.Text = $dialog.SelectedPath
    }
})

# =========================
# MAIN
# =========================

$startBtn.Add_Click({

try {
    $IdhBatchFile = Join-Path -Path $PSScriptRoot -ChildPath "BatchIdh.xlsx"
    $startBtn.Enabled = $false
    $logBox.Clear()

    $idhFile = $IdhBatchFile
    $certificationFolder = $certificationField.Value

    if ([string]::IsNullOrWhiteSpace($idhFile) -or !(Test-Path $idhFile)) {
        Log "ERROR: BatchIdh.xlsx not found"
        return
    }

    if ([string]::IsNullOrWhiteSpace($certificationFolder) -or !(Test-Path $certificationFolder)) {
        Log "ERROR: Certificates folder missing"
        return
    }

    Log "Indexing PDFs..."
    $allPdf = Get-ChildItem $certificationFolder -Recurse -Filter *.pdf -File
    Log "Found $($allPdf.Count) PDF files in $certificationFolder"

    # Build a PDF candidate lookup keyed by IDH to speed matching.
    $uniqueIdhs = @{}
    $pdfCandidatesByIdh = @{}
    # we build the lookup after the workbook is open and table columns are known

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    try {
        $wb = $excel.Workbooks.Open($idhFile)
        Log "Opened workbook: $idhFile"
    } catch {
        Log "ERROR opening workbook: $($_.ToString())"
        throw
    }
    try {
        $ws = $wb.Worksheets.Item("Table_1")
    } catch {
        $ws = $wb.Worksheets.Item(1)
        Log "WARNING: Worksheet 'Table_1' not found, using worksheet '$($ws.Name)'"
    }
    Log "Using worksheet: $($ws.Name)"

    try {
        $table = $ws.ListObjects.Item("Table_1")
    } catch {
        if ($ws.ListObjects.Count -gt 0) {
            $table = $ws.ListObjects.Item(1)
            Log "WARNING: Table 'Table_1' not found, using table '$($table.Name)'"
        } else {
            Log "ERROR: No table found on worksheet '$($ws.Name)'"
            throw "No table found on worksheet '$($ws.Name)'"
        }
    }
    Log "Using table: $($table.Name) with $($table.DataBodyRange.Rows.Count) rows"

    $idCol = $table.ListColumns["IDH"].Index
    $batchCol = $table.ListColumns["Batch"].Index

    # Build a PDF candidate lookup keyed by IDH to speed matching.
    $pdfCandidatesByIdh = @{}
    if ($table.DataBodyRange -ne $null) {
        $idhSet = @{}
        foreach ($row in $table.DataBodyRange.Rows) {
            $idh = [string]$row.Cells.Item(1, $idCol).Value2
            if (-not [string]::IsNullOrWhiteSpace($idh)) {
                $idhLower = $idh.ToLower().Trim()
                if (-not $idhSet.ContainsKey($idhLower)) { $idhSet[$idhLower] = $true }
            }
        }
        foreach ($pdf in $allPdf) {
            $pdfNameLower = $pdf.Name.ToLower()
            foreach ($idhLower in $idhSet.Keys) {
                if ($pdfNameLower.Contains($idhLower)) {
                    if (-not $pdfCandidatesByIdh.ContainsKey($idhLower)) { $pdfCandidatesByIdh[$idhLower] = @() }
                    $pdfCandidatesByIdh[$idhLower] += $pdf
                }
            }
        }
    }

    # Prepare counters
    $redCount = 0
    $orangeCount = 0
    $blueCount = 0
    $greenCount = 0

    $fileColIndex = $null

    $reportRows = @()

    # First pass: count keys for rows where batch starts with 08
    $keyCounts = @{}
    if ($table.DataBodyRange -ne $null) {
        $seen = @{}
        foreach ($row in $table.DataBodyRange.Rows) {
            $idh = [string]$row.Cells.Item(1, $idCol).Value2
            $batchCell = $row.Cells.Item(1, $batchCol)
            $batch = $batchCell.Text
            if ([string]::IsNullOrWhiteSpace($batch)) { $batch = [string]$batchCell.Value2 }
            $batch = $batch.Trim()
            if ($batch.Length -lt 2) { $batch = $batch.PadLeft(2,'0') }
            if ($batch -and $batch.Length -ge 2 -and $batch.Substring(0,2) -eq '08') {
                $key = "$idh|$batch"
                $keyCounts[$key] = ($keyCounts[$key] + 0) + 1
            }
        }
    }

    # Second pass: apply coloring and write FoundPath
    if ($table.DataBodyRange -ne $null) {
        foreach ($row in $table.DataBodyRange.Rows) {
            $idh = [string]$row.Cells.Item(1, $idCol).Value2
            $batchCell = $row.Cells.Item(1, $batchCol)
            $batch = $batchCell.Text
            if ([string]::IsNullOrWhiteSpace($batch)) { $batch = [string]$batchCell.Value2 }
            $batch = $batch.Trim()
            if ($batch.Length -lt 2) { $batch = $batch.PadLeft(2,'0') }
            $key = "$idh|$batch"
            if (-not $seen.ContainsKey($key)) { $seen[$key] = 0 }
            $seen[$key] = $seen[$key] + 1

            if (-not ($batch -and $batch.Length -ge 2 -and $batch.Substring(0,2) -eq '08')) {
                # Non-08: do not modify original, just record
                $redCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Non-08'; FoundPath = ''; ExcelRow = $row.Row }
                continue
            }

            $count = ($keyCounts[$key] + 0)
            if ($count -gt 1 -and $seen[$key] -gt 1) {
                # Duplicate (2nd+): do not modify original, just record
                $orangeCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Duplicate'; FoundPath = ''; ExcelRow = $row.Row }
                continue
            }

            # Unique 08 row: search for PDF among candidate files for this IDH
            $idhLower = $idh.ToLower().Trim()
            $batchLower = $batch.ToLower()
            $match = $null
            $candidates = @()
            if ($pdfCandidatesByIdh.ContainsKey($idhLower)) { $candidates = $pdfCandidatesByIdh[$idhLower] }
            foreach ($pdf in $candidates) {
                $pdfName = $pdf.Name.ToLower()
                if ($pdfName.Contains($batchLower)) {
                    $match = $pdf
                    break
                }
            }

            if ($match) {
                # Unique & found: record only
                $blueCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Found'; FoundPath = $match.FullName; ExcelRow = $row.Row }
            } else {
                # Unique & missing: record only
                $greenCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Missing'; FoundPath = ''; ExcelRow = $row.Row }
            }
        }
    }
    $wb.Close($false)
    $excel.Quit()

    # Export an Excel workbook with color formatting for easy viewing
    try {
        $reportXlsx = Join-Path $PSScriptRoot ("CertificateReport_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $repExcel = New-Object -ComObject Excel.Application
        $repExcel.Visible = $false
        $repWb = $repExcel.Workbooks.Add()
        $repWs = $repWb.Worksheets.Item(1)

        # Headers
        $repWs.Cells.Item(1,1).Value2 = 'IDH'
        $repWs.Cells.Item(1,2).Value2 = 'Batch'
        $repWs.Cells.Item(1,3).Value2 = 'Status'
        $repWs.Cells.Item(1,4).Value2 = 'FoundPath'
        $repWs.Cells.Item(1,5).Value2 = 'ExcelRow'

        $rowIndex = 2
        foreach ($r in $reportRows) {
            $repWs.Cells.Item($rowIndex,1).Value2 = [string]($r.IDH -ne $null ? $r.IDH : '')
            $repWs.Cells.Item($rowIndex,2).Value2 = [string]($r.Batch -ne $null ? $r.Batch : '')
            $repWs.Cells.Item($rowIndex,3).Value2 = [string]($r.Status -ne $null ? $r.Status : '')
            $foundPath = [string]($r.FoundPath -ne $null ? $r.FoundPath : '')
            if ($foundPath) {
                $cell = $repWs.Cells.Item($rowIndex,4)
                try {
                    $repWs.Hyperlinks.Add($cell, $foundPath, "", "", $foundPath)
                } catch {
                    $repWs.Cells.Item($rowIndex,4).Value2 = $foundPath
                }
            } else {
                $repWs.Cells.Item($rowIndex,4).Value2 = ""
            }
            $repWs.Cells.Item($rowIndex,5).Value2 = [string]($r.ExcelRow -ne $null ? $r.ExcelRow : '')

            # Apply color based on status
            $color = 0xFFFFFF
            switch ($r.Status) {
                'Non-08' { $color = 0xC0C0C0 }      # red
                'Duplicate' { $color = 0xFFA500 }   # orange
                'Found' { $color = 0x99CCFF }       # light blue
                'Missing' { $color = 0xCCFFCC }     # light green
            }
            try {
                $repWs.Range(("A{0}:E{0}" -f $rowIndex)).Interior.Color = $color
            } catch { }

            $rowIndex++
        }

        # Autofit columns
        $repWs.Columns.Item("A:E").AutoFit()

        $repWb.SaveAs($reportXlsx, 51)
        $repWb.Close($false)
        $repExcel.Quit()
        Log "Excel report saved: $reportXlsx"
    } catch {
        Log "ERROR saving Excel report: $($_.ToString())"
    }

    Log "Red (non-08): $redCount, Orange (duplicates): $orangeCount, Blue (unique+found): $blueCount, Green (unique+missing): $greenCount"


    Log "Processing completed."

}
catch {
    Log ($_.ToString())
}
finally {
    $startBtn.Enabled = $true
}

})

$form.ShowDialog()
