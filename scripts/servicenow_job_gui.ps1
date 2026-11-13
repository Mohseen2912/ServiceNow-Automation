Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$workspaceRoot = Split-Path -Parent $PSScriptRoot
$agentScript = Join-Path $PSScriptRoot "servicenow_job_agent.py"
$sampleWorkbookDirectory = Join-Path $workspaceRoot "outputs\servicenow_excel_sample"
$sampleWorkbookPath = Join-Path $sampleWorkbookDirectory "servicenow_changes_sample.xlsx"

function Quote-ProcessArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Set-DefaultOutputPath {
    param([string]$PipelinePath)

    if ([string]::IsNullOrWhiteSpace($PipelinePath)) {
        return
    }

    $directory = [System.IO.Path]::GetDirectoryName($PipelinePath)
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($PipelinePath)
    $extension = [System.IO.Path]::GetExtension($PipelinePath)
    $outputTextBox.Text = Join-Path $directory ($filename + "_servicenow" + $extension)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "ServiceNow Pipeline Generator"
$form.StartPosition = "CenterScreen"
$form.ClientSize = New-Object System.Drawing.Size(760, 455)
$form.MinimumSize = New-Object System.Drawing.Size(790, 500)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Create a ServiceNow-enabled pipeline"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 17)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$titleLabel.Location = New-Object System.Drawing.Point(28, 22)
$titleLabel.AutoSize = $true
$form.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Select the pipeline and Excel row. A new YAML file will be created; the original remains unchanged."
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$subtitleLabel.Location = New-Object System.Drawing.Point(31, 61)
$subtitleLabel.Size = New-Object System.Drawing.Size(690, 24)
$form.Controls.Add($subtitleLabel)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(28, 96)
$panel.Size = New-Object System.Drawing.Size(704, 274)
$panel.BackColor = [System.Drawing.Color]::White
$panel.BorderStyle = "FixedSingle"
$form.Controls.Add($panel)

function Add-FieldLabel {
    param([string]$Text, [int]$Top)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point(18, $Top)
    $label.Size = New-Object System.Drawing.Size(135, 24)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $panel.Controls.Add($label)
}

function New-TextBox {
    param([int]$Top, [int]$Width = 440)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(158, $Top)
    $textBox.Size = New-Object System.Drawing.Size($Width, 28)
    $panel.Controls.Add($textBox)
    return $textBox
}

function New-BrowseButton {
    param([int]$Top, [scriptblock]$OnClick)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Browse..."
    $button.Location = New-Object System.Drawing.Point(608, $Top - 1)
    $button.Size = New-Object System.Drawing.Size(78, 30)
    $button.FlatStyle = "System"
    $button.Add_Click($OnClick)
    $panel.Controls.Add($button)
    return $button
}

Add-FieldLabel -Text "Pipeline YAML" -Top 21
$pipelineTextBox = New-TextBox -Top 18
$pipelineBrowse = New-BrowseButton -Top 18 -OnClick {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select an Azure pipeline"
    $dialog.Filter = "YAML pipeline (*.yml;*.yaml)|*.yml;*.yaml|All files (*.*)|*.*"
    $dialog.InitialDirectory = $workspaceRoot
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    $dialog.RestoreDirectory = $true
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $pipelineTextBox.Text = $dialog.FileName
        Set-DefaultOutputPath -PipelinePath $dialog.FileName
    }
}

Add-FieldLabel -Text "Excel / CSV" -Top 73
$excelTextBox = New-TextBox -Top 70
if ([System.IO.File]::Exists($sampleWorkbookPath)) {
    $excelTextBox.Text = $sampleWorkbookPath
}
$excelBrowse = New-BrowseButton -Top 70 -OnClick {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select the ServiceNow input file"
    $dialog.Filter = "Excel or CSV (*.xlsx;*.xlsm;*.csv)|*.xlsx;*.xlsm;*.csv|Excel workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|CSV file (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.InitialDirectory = if ([System.IO.Directory]::Exists($sampleWorkbookDirectory)) {
        $sampleWorkbookDirectory
    }
    else {
        $workspaceRoot
    }
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    $dialog.RestoreDirectory = $true
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $excelTextBox.Text = $dialog.FileName
    }
}

Add-FieldLabel -Text "Worksheet" -Top 125
$sheetTextBox = New-TextBox -Top 122 -Width 205
$sheetTextBox.Text = "Changes"

$rowLabel = New-Object System.Windows.Forms.Label
$rowLabel.Text = "Excel row"
$rowLabel.Location = New-Object System.Drawing.Point(393, 125)
$rowLabel.Size = New-Object System.Drawing.Size(82, 24)
$rowLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$panel.Controls.Add($rowLabel)

$rowPicker = New-Object System.Windows.Forms.NumericUpDown
$rowPicker.Location = New-Object System.Drawing.Point(481, 122)
$rowPicker.Size = New-Object System.Drawing.Size(117, 28)
$rowPicker.Minimum = 2
$rowPicker.Maximum = 1000000
$rowPicker.Value = 2
$panel.Controls.Add($rowPicker)

Add-FieldLabel -Text "Output YAML" -Top 177
$outputTextBox = New-TextBox -Top 174
$outputBrowse = New-BrowseButton -Top 174 -OnClick {
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Save the generated pipeline"
    $dialog.Filter = "YAML pipeline (*.yml)|*.yml|YAML pipeline (*.yaml)|*.yaml"
    $dialog.DefaultExt = "yml"
    $dialog.RestoreDirectory = $true
    if (-not [string]::IsNullOrWhiteSpace($outputTextBox.Text)) {
        $dialog.FileName = $outputTextBox.Text
    }
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputTextBox.Text = $dialog.FileName
    }
}

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = "Row 1 contains headers. The default sample values are in row 2."
$hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$hintLabel.Location = New-Object System.Drawing.Point(158, 218)
$hintLabel.Size = New-Object System.Drawing.Size(510, 24)
$panel.Controls.Add($hintLabel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$statusLabel.Location = New-Object System.Drawing.Point(31, 392)
$statusLabel.Size = New-Object System.Drawing.Size(520, 30)
$form.Controls.Add($statusLabel)

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "Generate Pipeline"
$generateButton.Location = New-Object System.Drawing.Point(565, 386)
$generateButton.Size = New-Object System.Drawing.Size(167, 42)
$generateButton.BackColor = [System.Drawing.Color]::FromArgb(3, 105, 161)
$generateButton.ForeColor = [System.Drawing.Color]::White
$generateButton.FlatStyle = "Flat"
$generateButton.FlatAppearance.BorderSize = 0
$generateButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$form.Controls.Add($generateButton)
$form.AcceptButton = $generateButton

$generateButton.Add_Click({
    $pipelinePath = $pipelineTextBox.Text.Trim()
    $excelPath = $excelTextBox.Text.Trim()
    $sheetName = $sheetTextBox.Text.Trim()
    $outputPath = $outputTextBox.Text.Trim()

    if (-not [System.IO.File]::Exists($pipelinePath)) {
        [System.Windows.Forms.MessageBox]::Show("Select an existing pipeline YAML file.", "Pipeline required", "OK", "Warning") | Out-Null
        return
    }
    if (-not [System.IO.File]::Exists($excelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Select an existing Excel or CSV file.", "Excel file required", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($outputPath)) {
        [System.Windows.Forms.MessageBox]::Show("Choose where to save the generated YAML file.", "Output required", "OK", "Warning") | Out-Null
        return
    }

    $pipelineFullPath = [System.IO.Path]::GetFullPath($pipelinePath)
    $outputFullPath = [System.IO.Path]::GetFullPath($outputPath)
    if ($pipelineFullPath.Equals($outputFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        [System.Windows.Forms.MessageBox]::Show("Choose a different output name so the original pipeline remains unchanged.", "Different output required", "OK", "Warning") | Out-Null
        return
    }

    if ([System.IO.File]::Exists($outputFullPath)) {
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "The output file already exists. Replace it?",
            "Confirm replacement",
            "YesNo",
            "Question"
        )
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    $pythonCommand = Get-Command "python.exe" -ErrorAction SilentlyContinue
    $pythonPrefixArguments = @()
    if ($null -eq $pythonCommand) {
        $pythonCommand = Get-Command "py.exe" -ErrorAction SilentlyContinue
        $pythonPrefixArguments = @("-3")
    }
    if ($null -eq $pythonCommand) {
        [System.Windows.Forms.MessageBox]::Show(
            "Python was not found. Ask the automation owner to install Python or provide a packaged application.",
            "Python not found",
            "OK",
            "Error"
        ) | Out-Null
        return
    }

    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputFullPath)
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
    $temporaryPath = Join-Path $outputDirectory (([System.IO.Path]::GetRandomFileName()) + ".yml")

    $generateButton.Enabled = $false
    $statusLabel.Text = "Generating pipeline..."
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(3, 105, 161)
    $form.Refresh()

    try {
        [System.IO.File]::Copy($pipelineFullPath, $temporaryPath, $true)

        $arguments = @($pythonPrefixArguments) + @(
            $agentScript,
            $temporaryPath,
            "--excel",
            [System.IO.Path]::GetFullPath($excelPath),
            "--row",
            [string][int]$rowPicker.Value
        )
        if (-not [string]::IsNullOrWhiteSpace($sheetName)) {
            $arguments += @("--sheet", $sheetName)
        }

        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $pythonCommand.Source
        $processInfo.Arguments = ($arguments | ForEach-Object { Quote-ProcessArgument -Value $_ }) -join " "
        $processInfo.WorkingDirectory = $workspaceRoot
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        [void]$process.Start()
        $standardOutput = $process.StandardOutput.ReadToEnd()
        $standardError = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        if ($process.ExitCode -ne 0) {
            $details = ($standardError + [Environment]::NewLine + $standardOutput).Trim()
            throw "The generator could not update the pipeline.`n`n$details"
        }

        [System.IO.File]::Copy($temporaryPath, $outputFullPath, $true)
        $statusLabel.Text = "Created: " + [System.IO.Path]::GetFileName($outputFullPath)
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
        [System.Windows.Forms.MessageBox]::Show(
            "Pipeline created successfully:`n$outputFullPath",
            "Generation complete",
            "OK",
            "Information"
        ) | Out-Null
    }
    catch {
        $statusLabel.Text = "Generation failed"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Generation failed", "OK", "Error") | Out-Null
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) {
            [System.IO.File]::Delete($temporaryPath)
        }
        $generateButton.Enabled = $true
    }
})

[void]$form.ShowDialog()
