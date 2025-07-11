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
            $folderDialog.MultiSelect = $true
            $folderDialog.Description = "Select a directory to add"
            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPaths = $folderDialog.SelectedPaths
                foreach ($selectedPath in $selectedPaths) {
                    if (-not $pathsBox.Items.Contains($selectedPath)) {
                        $pathsBox.Items.Add($selectedPath)
                    }
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
    $pathEntryPanel.Controls.AddRange(@($pathsLabel, $pathsBox))

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
    $scanButton.AutoSize = $true
    $scanButton.Add_click({
        $OUResult = Select-OU
        if ($OUResult.DialogResult -ne [System.Windows.Forms.DialogResult]) {
            $result = Get-UserAccessChk $pathsBox.Items -SearchBase $OUResult[-1]
            Save-GUI $result
        }
    })
    $bottomPanel.Controls.Add($scanButton)
    $bottomPanel.Controls.Add($removePathButton)
    $bottomPanel.Controls.Add($addPathButton)

    [System.Windows.Forms.Application]::Run($form)
}

function Select-OU {
    Import-Module ActiveDirectory
    $form = New-Object System.Windows.Forms.Form

    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = 'Fill'

    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.AutoSize = $true
    $treeView.Dock = 'Fill'
    $domainRoot = Get-ADDomain
    # Recursive Function to populate Tree
    function Add-OUsToTree {
        param(
            [string]$baseDN,
            [System.Windows.Forms.TreeNode]$parentNode
        )
        $childOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $baseDN -SearchScope OneLevel

        foreach ($ou in $childOUs) {
            $node = New-Object System.Windows.Forms.TreeNode
            $node.Text = $ou.Name
            $node.Tag = $ou.DistinguishedName
            $parentNode.Nodes.Add($node)
            Add-OUsToTree -baseDN $ou.DistinguishedName -parentNode $node
        }
    }
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Text = $domainRoot.Name
    $rootNode.Tag = $domainRoot.DistinguishedName
    $treeView.Nodes.Add($rootNode)
    Add-OUsToTree $domainRoot.DistinguishedName $rootNode
    $treeView.ExpandAll()
    $table.Controls.Add($treeView)
    
    $bottomPanel = New-Object System.Windows.Forms.FlowLayoutPanel

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.AutoSize = 'True'
    $okButton.Text = 'Select'
    $okButton.Add_Click({
        Write-Host $treeView.SelectedNode.Tag
        if ($treeView.SelectedNode.Tag) {
            $form.DialogResult = 'Ok'
            $form.Tag = $treeView.SelectedNode.Tag
            $form.Close()
        }
    })
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.AutoSize = 'True'
    $cancelButton.Text = 'Cancel'
    $form.CancelButton = $cancelButton

    $bottomPanel.Controls.Add($cancelButton)
    $bottomPanel.Controls.Add($okButton)
    $bottomPanel.Dock = 'Bottom'
    $bottomPanel.AutoSize = 'True'
    $bottomPanel.FlowDirection = 'RightToLeft'
    $form.Controls.Add($bottomPanel)

    $form.Controls.add($table)

    $form.ShowDialog()
    return $form.Tag
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
                $output = $null
                if (Test-Path $dir -PathType Leaf) {
                    $output = accesschk64.exe $user.SamAccountName $dir -nobanner 2>&1
                }
                else {
                    $output = accesschk64.exe $user.SamAccountName $dir -nobanner -d 2>&1
                }

                if ($output -match "No matching objects found.") {
                    $result[$dir] = "Error"
                    continue
                }

                $read  = ($output[0] -eq "R")
                $write = ($output[1] -eq "W")
                $result[$dir] = if ($read -or $write) { $true } else { $false }
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