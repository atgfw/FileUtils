Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set DPI awareness (Windows 10+)
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DpiFix {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@

[DpiFix]::SetProcessDPIAware()


function ScannerGui {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Access Report'
    $form.Size = '600,400'
    $form.StartPosition = 'CenterScreen'

    # Main Panel Path Entry
    $pathEntryPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $pathEntryPanel.AutoSizeMode = 'GrowAndShrink'
    $pathEntryPanel.autoSize = $true
    $pathEntryPanel.Dock = 'Fill'
    $pathsLabel = New-Object System.Windows.Forms.Label
    $pathsLabel.Text = "Paths to Scan"
    $pathsLabel.AutoSize = $true

    $pathsBox = New-Object System.Windows.Forms.ListBox
    $pathsBox.Dock = 'Fill'

    # Paths Add/Remove Buttons
    $pathsAddRemovePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $pathsAddRemovePanel.FlowDirection = 'LeftToRight'
    $pathsAddRemovePanel.AutoSizeMode = 'GrowAndShrink'
    $pathsAddRemovePanel.autoSize = $true
    $addPathButton = New-Object System.Windows.Forms.Button
    $addPathButton.Text = 'Add Directory'
    $addPathButton.AutoSize = $true
    $addPathButton.Add_Click({
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select a directory to add"
            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $folderDialog.SelectedPath
                if (-not $pathsBox.Items.Contains($selectedPath)) {
                    $pathsBox.Items.Add($selectedPath)
                }
            }
        })

    $pathsAddRemovePanel.Controls.Add($addPathButton)
    $removePathButton = New-Object System.Windows.Forms.Button
    $removePathButton.Text = 'Remove Dirctory'
    $removePathButton.AutoSize = $true
    $removePathButton.Add_Click({
            # Copy selected items to a list to avoid collection modification errors
            $selectedItems = @()
            foreach ($item in $pathsBox.SelectedItems) {
                $selectedItems += $item
            }

            foreach ($item in $selectedItems) {
                $pathsBox.Items.Remove($item)
            }
        })
    $pathsAddRemovePanel.Controls.Add($removePathButton)
    $pathEntryPanel.Controls.AddRange(@($pathsLabel, $pathsBox, $pathsAddRemovePanel))

    $form.Controls.Add($pathEntryPanel)

    # Bottom panel
    $bottomPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottomPanel.Dock = 'Bottom'
    $bottomPanel.AutoSize = $true
    $bottomPanel.AutoSizeMode = 'GrowAndShrink'
    $bottomPanel.FlowDirection = 'RightToLeft'
    $form.Controls.Add($bottomPanel)

    # Scan Button
    $scanButton = New-Object System.Windows.Forms.Button
    $scanButton.Text = 'Scan'
    $scanButton.Add_click({
        $result = Get-UserAccessChk $pathsBox.Items
        Save-GUI $result
    })
    $bottomPanel.Controls.Add($scanButton)

    [System.Windows.Forms.Application]::Run($form)
}

function Save-GUI {
    param (
        [Parameter(Mandatory)]
        $ItemToSave
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)


    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = 'Fill'
    $table.RowCount = 2
    $form.Controls.Add($table)

    $label = New-Object System.Windows.Forms.Label
    $label.text = 'Scan Completed!'
    $label.TextAlign = 'TopCenter'
    $label.Dock = 'Fill'
    $label.autosize = $true
    $table.Controls.Add($label,0,0)

    $buttons = New-Object System.Windows.Forms.TableLayoutPanel
    $buttons.Dock = 'None'
    $buttons.Anchor = 'None'
    $buttons.AutoSize = $true
    $buttons.ColumnCount = 3
    $buttons.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle 'AutoSize'))
    $buttons.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle 'AutoSize'))
    $buttons.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle 'AutoSize'))
    $table.Controls.add($buttons,0,1)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Autosize = $true
    $saveButton.Add_click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $saveDialog.FilterIndex = 1
        if ($saveDialog.ShowDialog() -eq "OK") {
            $filePath = $saveDialog.FileName
            try {
                $ItemToSave | Export-CSV -Path $filePath
                [System.Windows.Forms.MessageBox]::Show("File saved to:`n$filePath", "Success", "OK", "Information")
                $form.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to save file:`n$($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
    })
    $previewButton = New-Object System.Windows.Forms.Button
    $previewButton.Text = "Preview"
    $previewButton.Autosize = $true
    $previewButton.Add_click({
        $ItemToSave | Out-GridView
    })
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Autosize = $true
    $form.CancelButton = $cancelButton
    $buttons.Controls.AddRange(@($saveButton,$previewButton,$cancelButton))

    $form.PerformAutoScale()
    $form.ShowDialog()
}

function Get-UserAccessChk {
    param (
        [Parameter(Mandatory)]
        [string[]]$Directories,

        [Parameter()]
        [string]$SearchBase = "OU=Users,OU=Sperry Van Ness,DC=parkegroup,DC=local"
    )

    # Validate directories
    $ValidDirs = $Directories | Where-Object { Test-Path $_ }
    if ($ValidDirs.Count -eq 0) {
        throw "No valid directories provided."
    }

    # Retrieve enabled AD users
    Write-Host "Fetching enabled AD users..."
    $UserNames = Get-ADUser -Filter {Enabled -eq $true} -SearchBase $SearchBase | 
        Select-Object Name, SAMAccountName

    if (-not $UserNames) {
        throw "No enabled users found in $SearchBase"
    }

    $Results = @()
    $total = $UserNames.Count
    $i = 0

    foreach ($user in $UserNames) {
        $i++
        Write-Progress -Activity "Scanning Permissions" -Status "Scanning access for $($user.Name)" -PercentComplete (($i / $total) * 100)

        $result = [ordered]@{ Name = $user.Name }

        foreach ($dir in $ValidDirs) {
            try {
                $output = accesschk64.exe $user.SamAccountName $dir -nobanner -d 2>&1

                if ($output -match "No matching objects found.") {
                    $result[$dir] = "Error"
                    continue
                }

                $read  = ($output[0] -eq "R")
                $write = ($output[1] -eq "W")
                $result[$dir] = if ($read -or $write) { "Yes" } else { "No" }
            }
            catch {
                $result[$dir] = "Error"
            }
        }

        $Results += New-Object PSObject -Property $result
    }
    Write-Progress -Activity "Scanning Permissions" -Completed
    
    return $Results
}

ScannerGui