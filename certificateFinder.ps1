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

    $pathLabel = New-Object System.Windows.Forms.TextBox
    $pathLabel.Location = New-Object System.Drawing.Point(180,$yPos)
    $pathLabel.Size = New-Object System.Drawing.Size(400,28)
    $pathLabel.Text = "Not selected"
    $pathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pathLabel.BackColor = [System.Drawing.Color]::White
    $pathLabel.ForeColor = [System.Drawing.Color]::FromArgb(64,64,64)
    $pathLabel.Margin = New-Object System.Windows.Forms.Padding(0)
    $pathLabel.Multiline = $false

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

$spinnerIndex = 0
$spinnerTimer = New-Object System.Windows.Forms.Timer
$spinnerTimer.Interval = 200
$spinnerTimer.Add_Tick({
    if (-not $startBtn.Enabled) {
        $spinnerIndex = ($spinnerIndex + 1) % 4
        switch ($spinnerIndex) {
            0 { $startBtn.Text = 'START' }
            1 { $startBtn.Text = 'START.' }
            2 { $startBtn.Text = 'START..' }
            3 { $startBtn.Text = 'START...' }
        }
    }
})

function Start-Spinner {
    $spinnerIndex = 0
    $startBtn.Text = 'START'
    $spinnerTimer.Start()
}

function Stop-Spinner {
    $spinnerTimer.Stop()
    $startBtn.Text = 'START'
}

# LOG
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Location = New-Object System.Drawing.Point(10,190)
$logBox.Size = New-Object System.Drawing.Size(710,240)
$logBox.ReadOnly = $true
$logBox.HideSelection = $false
Set-Win11Style -Control $logBox -Type 'Log'
$form.Controls.Add($logBox)
function Log($msg, $color = $null) {
    if (-not $color) {
        if ($msg.StartsWith('ERROR')) { $color = [System.Drawing.Color]::FromArgb(220,20,60) }
        elseif ($msg.StartsWith('WARNING')) { $color = [System.Drawing.Color]::FromArgb(218,165,32) }
        elseif ($msg.StartsWith('Missing')) { $color = [System.Drawing.Color]::FromArgb(255,0,0) }
        elseif ($msg.StartsWith('Found')) { $color = [System.Drawing.Color]::FromArgb(0,128,0) }
        elseif ($msg.StartsWith('Processing') -or $msg.StartsWith('Indexing') -or $msg.StartsWith('Creating') -or $msg.StartsWith('Using')) { $color = [System.Drawing.Color]::FromArgb(0,0,0) }
        else { $color = [System.Drawing.Color]::Black }
    }
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor = $color
    $logBox.AppendText($msg + "`r`n")
    $logBox.SelectionColor = [System.Drawing.Color]::Black
    $logBox.Refresh()
}

# =========================
# BROWSE
# =========================

$certificationField.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the certificates folder"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq "OK") {
        $certificationField.Value = $dialog.SelectedPath
        $certificationField.DisplayLabel.Text = $dialog.SelectedPath
    }
})

# =========================
# COPY VERIFICATION FUNCTION
# =========================

function Copy-FileWithVerification {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 500
    )
    
    $sourceFile = Get-Item $SourcePath -ErrorAction SilentlyContinue
    if (-not $sourceFile) {
        return [PSCustomObject]@{
            Success = $false
            Error = "Source file not found or inaccessible"
            SourceSize = 0
            DestinationSize = 0
        }
    }
    
    $sourceSize = $sourceFile.Length
    $retryCount = 0
    $copySuccess = $false
    $lastError = ""
    
    while ($retryCount -lt $MaxRetries -and -not $copySuccess) {
        try {
            Copy-Item -Path $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop
            $copySuccess = $true
        } catch {
            $lastError = $_.Exception.Message
            $retryCount++
            if ($retryCount -lt $MaxRetries) {
                Start-Sleep -Milliseconds $RetryDelayMs
            }
        }
    }
    
    if (-not $copySuccess) {
        return [PSCustomObject]@{
            Success = $false
            Error = "Copy failed after $MaxRetries attempts: $lastError"
            SourceSize = $sourceSize
            DestinationSize = 0
        }
    }
    
    # Verify the copy
    $destFile = Get-Item $DestinationPath -ErrorAction SilentlyContinue
    if (-not $destFile) {
        return [PSCustomObject]@{
            Success = $false
            Error = "Destination file not found after copy"
            SourceSize = $sourceSize
            DestinationSize = 0
        }
    }
    
    $destSize = $destFile.Length
    
    if ($sourceSize -ne $destSize) {
        return [PSCustomObject]@{
            Success = $false
            Error = "File size mismatch: source=$sourceSize, destination=$destSize bytes"
            SourceSize = $sourceSize
            DestinationSize = $destSize
        }
    }
    
    return [PSCustomObject]@{
        Success = $true
        Error = ""
        SourceSize = $sourceSize
        DestinationSize = $destSize
    }
}

# =========================
# MAIN
# =========================

$startBtn.Add_Click({

try {

    #Set variables
    $IdhBatchFile = Join-Path -Path $PSScriptRoot -ChildPath "BatchIdh.xlsx"
    $startBtn.Enabled = $false
    Start-Spinner
    $logBox.Clear()
    $certificationFolder = $certificationField.DisplayLabel.Text
    $certificationField.Value = $certificationFolder

    #Check variables
    if ([string]::IsNullOrWhiteSpace($IdhBatchFile) -or !(Test-Path $IdhBatchFile)) {
        Log "ERROR: BatchIdh.xlsx not found"
        return
    }

    if ([string]::IsNullOrWhiteSpace($certificationFolder) -or !(Test-Path $certificationFolder)) {
        Log "ERROR: Certificates folder missing"
        return
    }

    #Start the process
    Log "Indexing PDFs..."
    $allPdf = Get-ChildItem $certificationFolder -Recurse -Filter *.pdf -File
    Log "Found $($allPdf.Count) PDF files in $certificationFolder"

    $outputFolder = Join-Path $PSScriptRoot ("FoundCertificates_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    $pdfOutputFolder = Join-Path $outputFolder "PDFs"
    New-Item -ItemType Directory -Path $pdfOutputFolder -Force | Out-Null
    Log "Creating output folders..."

    #Excel setup

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    try {
        $wb = $excel.Workbooks.Open($IdhBatchFile, $null, $true)
        Log "Opened workbook: $IdhBatchFile"
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

    # Prepare counters
    $redCount = 0
    $orangeCount = 0
    $blueCount = 0
    $greenCount = 0

    $reportRows = @()

    # First pass: count keys for rows where batch starts with 08
    $idCol = $table.ListColumns["IDH"].Index
    $batchCol = $table.ListColumns["Batch"].Index
    $keyCounts = @{}
    if ($table.DataBodyRange -ne $null) {
        foreach ($row in $table.DataBodyRange.Rows) {
            
           try {
                $idh = [string]$row.Cells.Item(1, $idCol).Value2
                $batch = [string]$row.Cells.Item(1, $batchCol).Value2
            } catch {
                Log "Excel COM error row $($row.Row) IDHcol=$idCol Batchcol= $batchCol $($_.Exception.Message)"
                
            }
            $key = "$idh-*-$batch"
            $keyCounts[$key] = ($keyCounts[$key] + 0) + 1
        }
    }

    # Second pass: apply coloring and write FoundPath
    if ($table.DataBodyRange -ne $null) {
        foreach ($row in $table.DataBodyRange.Rows) {
            try {
                $idh = [string]$row.Cells.Item(1, $idCol).Value2
                $batch = [string]$row.Cells.Item(1, $batchCol).Value2   
            } catch {
                Log "Excel COM error row $($row.Row) IDHcol=$idCol Batchcol= $batchCol $($_.Exception.Message)"
                
            }   
            $key = "$idh-*-$batch"
            Log "Processing IDH: $idh, Batch: $batch"
 

            if ($batch -and $batch -notlike '08*') {
                # Non-08: do not modify original, just record
                $redCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Non-08'; FoundPath = ''; ExcelRow = $row.Row }
                Log "Non-08 batch"
            }
            elseif ($keyCounts.ContainsKey($key) -and $keyCounts[$key] -gt 1) { 
                # Duplicate (2nd+): do not modify original, just record
                $keyCounts[$key] = $keyCounts[$key] -1 
                $orangeCount++
                $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Duplicate'; FoundPath = ''; ExcelRow = $row.Row }
                Log "Duplicate entry"
            }
            elseif ($batch -like '08*') {
                $match = $null
                foreach ($pdf in $allPdf) {
                    if ($pdf.Name -like "*$idh*" -and $pdf.Name -like "*$batch*") {
                        $match = $pdf
                        break
                    }
                }

                if ($match) {
                    $destinationName = $match.Name
                    $destinationPath = Join-Path $pdfOutputFolder $destinationName
                    $copyIndex = 1
                    while (Test-Path $destinationPath) {
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($match.Name)
                        $extension = [System.IO.Path]::GetExtension($match.Name)
                        $destinationName = "{0}_{1}{2}" -f $baseName, $copyIndex, $extension
                        $destinationPath = Join-Path $pdfOutputFolder $destinationName
                        $copyIndex++
                    }
                    
                    # Copy with verification
                    $copyResult = Copy-FileWithVerification -SourcePath $match.FullName -DestinationPath $destinationPath
                    
                    if ($copyResult.Success) {
                        # Unique & found: record only
                        $blueCount++
                        $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Found'; FoundPath = $match.FullName; ExcelRow = $row.Row }
                        Log "Found and copied: $($match.FullName) to $destinationPath (Size: $($copyResult.SourceSize) bytes)"
                    } else {
                        # Copy failed - treat as missing
                        $greenCount++
                        $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Missing copy failed'; FoundPath = $match.FullName; ExcelRow = $row.Row }
                        Log "Missing: $idh-$batch - ERROR: $($copyResult.Error)"
                    }
                } else {
                    # Unique & missing: record only
                    $greenCount++
                    $reportRows += [PSCustomObject]@{ IDH = $idh; Batch = $batch; Status = 'Missing'; FoundPath = ''; ExcelRow = $row.Row }
                    Log "Missing: $idh-$batch"
                }
            }
                     
        }
            
        }
    
    $wb.Close($false)
    $excel.Quit()

    # Export an Excel workbook with color formatting for easy viewing
    Log "Generating Excel report..."
    try {
        $reportXlsx = Join-Path $outputFolder ("CertificateReport_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $repExcel = New-Object -ComObject Excel.Application
        $repExcel.Visible = $false
        $repWb = $repExcel.Workbooks.Add()
        $repWs = $repWb.Worksheets.Item(1)

        # Headers
        $repWs.Cells.Item(1,1).Value2 = 'IDH'
        $repWs.Cells.Item(1,1).NumberFormat = '@'
        $repWs.Cells.Item(1,2).Value2 = 'Batch'
        $repWs.Cells.Item(1,2).NumberFormat = '@'
        $repWs.Cells.Item(1,3).Value2 = 'Status'
        $repWs.Cells.Item(1,4).Value2 = 'FoundPath'
        $repWs.Cells.Item(1,5).Value2 = 'ExcelRow'

        $rowIndex = 2
        foreach ($r in $reportRows) {
            $idhCell = $repWs.Cells.Item($rowIndex,1)
            $idhCell.Value2 = if ($null -ne $r.IDH) { [string]$r.IDH } else { '' }
            $idhCell.NumberFormat = '@'
            $batchCell = $repWs.Cells.Item($rowIndex,2)
            $batchCell.NumberFormat = '@'
            $batchCell.Value2 = if ($null -ne $r.Batch) { [string]$r.Batch } else { '' }
           
            $repWs.Cells.Item($rowIndex,3).Value2 = if ($null -ne $r.Status) { [string]$r.Status } else { '' }
            $foundPath = if ($null -ne $r.FoundPath) { [string]$r.FoundPath } else { '' }
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
            $repWs.Cells.Item($rowIndex,5).Value2 = if ($null -ne $r.ExcelRow) { [string]$r.ExcelRow } else { '' }

            # Apply color based on status
            $color = 0xFFFFFF
            switch ($r.Status) {
                'Non-08' { $color = 0x0000FF }      # red
                'Duplicate' { $color = 0x00A5FF }   # orange
                'Found' { $color = 0xFFCC99 }       # light blue
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
    Stop-Spinner
    $startBtn.Enabled = $true
}

})

$form.ShowDialog()
