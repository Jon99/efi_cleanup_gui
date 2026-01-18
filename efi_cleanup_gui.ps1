# Self-elevation mechanism (Disabled for EXE conversion - handled by -requireAdmin)
# if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
#     $arguments = "& '" + $myinvocation.mycommand.definition + "'"
#     Start-Process powershell -Verb runAs -ArgumentList $arguments
#     Break
# }

<#
.SYNOPSIS
    EFI/System Reserved Partition Cleanup Tool for Windows Update Issues
.DESCRIPTION
    Professional GUI tool to safely clean up the System Reserved/EFI partition
    Removes fonts and language folders to free space for Windows Updates
.NOTES
    Author: Elite System Administrator
    Version: 2.1
    Requires: Administrator privileges (Auto-elevates)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:selectedDrive = $null
$script:originalSpace = 0
$script:currentSpace = 0
$script:deletedItems = @()

# Main form setup
$form = New-Object System.Windows.Forms.Form
$form.Text = 'EFI/System Reserved Partition Cleanup Tool'
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(660, 30)
$titleLabel.Text = 'System Reserved / EFI Partition Cleanup Tool'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$form.Controls.Add($titleLabel)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 45)
$statusLabel.Size = New-Object System.Drawing.Size(660, 20)
$statusLabel.Text = 'Ready. Click "Scan Partitions" to begin.'
$statusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($statusLabel)

# Partition Selection GroupBox
$partitionGroup = New-Object System.Windows.Forms.GroupBox
$partitionGroup.Location = New-Object System.Drawing.Point(10, 70)
$partitionGroup.Size = New-Object System.Drawing.Size(660, 120)
$partitionGroup.Text = 'Step 1: Select Partition'
$partitionGroup.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($partitionGroup)

# Scan Button
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(20, 25)
$scanButton.Size = New-Object System.Drawing.Size(150, 35)
$scanButton.Text = 'Scan Partitions'
$scanButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$scanButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$scanButton.ForeColor = [System.Drawing.Color]::White
$scanButton.FlatStyle = 'Flat'
$partitionGroup.Controls.Add($scanButton)

# Drive ComboBox
$driveCombo = New-Object System.Windows.Forms.ComboBox
$driveCombo.Location = New-Object System.Drawing.Point(20, 70)
$driveCombo.Size = New-Object System.Drawing.Size(300, 25)
$driveCombo.DropDownStyle = 'DropDownList'
$driveCombo.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$driveCombo.Enabled = $false
$partitionGroup.Controls.Add($driveCombo)

# Drive Letter Selection ComboBox
$letterCombo = New-Object System.Windows.Forms.ComboBox
$letterCombo.Location = New-Object System.Drawing.Point(330, 70)
$letterCombo.Size = New-Object System.Drawing.Size(50, 25)
$letterCombo.DropDownStyle = 'DropDownList'
$letterCombo.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$letterCombo.Enabled = $false
$partitionGroup.Controls.Add($letterCombo)

# Assign Drive Letter Button
$assignButton = New-Object System.Windows.Forms.Button
$assignButton.Location = New-Object System.Drawing.Point(390, 68)
$assignButton.Size = New-Object System.Drawing.Size(120, 30)
$assignButton.Text = 'Assign Letter'
$assignButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$assignButton.Enabled = $false
$partitionGroup.Controls.Add($assignButton)

# Unassign Drive Letter Button
$unassignButton = New-Object System.Windows.Forms.Button
$unassignButton.Location = New-Object System.Drawing.Point(520, 68)
$unassignButton.Size = New-Object System.Drawing.Size(120, 30)
$unassignButton.Text = 'Unassign Letter'
$unassignButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$unassignButton.Enabled = $false
$partitionGroup.Controls.Add($unassignButton)

# Space Info Label
$spaceLabel = New-Object System.Windows.Forms.Label
$spaceLabel.Location = New-Object System.Drawing.Point(190, 30)
$spaceLabel.Size = New-Object System.Drawing.Size(450, 30)
$spaceLabel.Text = ''
$spaceLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$partitionGroup.Controls.Add($spaceLabel)

# Cleanup Options GroupBox
$cleanupGroup = New-Object System.Windows.Forms.GroupBox
$cleanupGroup.Location = New-Object System.Drawing.Point(10, 200)
$cleanupGroup.Size = New-Object System.Drawing.Size(660, 180)
$cleanupGroup.Text = 'Step 2: Select Cleanup Options'
$cleanupGroup.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($cleanupGroup)

# Font Cleanup Checkbox
$fontCheckbox = New-Object System.Windows.Forms.CheckBox
$fontCheckbox.Location = New-Object System.Drawing.Point(20, 30)
$fontCheckbox.Size = New-Object System.Drawing.Size(600, 25)
$fontCheckbox.Text = 'Remove EFI Boot Fonts (EFI\Microsoft\Boot\Fonts\*.*)'
$fontCheckbox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$fontCheckbox.Checked = $true
$fontCheckbox.Enabled = $false
$cleanupGroup.Controls.Add($fontCheckbox)

# Language Folders Checkbox
$langCheckbox = New-Object System.Windows.Forms.CheckBox
$langCheckbox.Location = New-Object System.Drawing.Point(20, 60)
$langCheckbox.Size = New-Object System.Drawing.Size(600, 25)
$langCheckbox.Text = 'Remove Unused Language Folders (jp, cz, dk, de, gr, es, hr, it, kr, lt, lv, no, nl, pl, ro, ru, sk, se, br, tw, pt)'
$langCheckbox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$langCheckbox.Checked = $true
$langCheckbox.Enabled = $false
$cleanupGroup.Controls.Add($langCheckbox)

# Backup Checkbox
$backupCheckbox = New-Object System.Windows.Forms.CheckBox
$backupCheckbox.Location = New-Object System.Drawing.Point(20, 90)
$backupCheckbox.Size = New-Object System.Drawing.Size(600, 25)
$backupCheckbox.Text = 'Create backup before deletion (requires additional space on C:)'
$backupCheckbox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$backupCheckbox.Checked = $false
$backupCheckbox.Enabled = $false
$cleanupGroup.Controls.Add($backupCheckbox)

# Execute Cleanup Button
$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Location = New-Object System.Drawing.Point(20, 130)
$cleanupButton.Size = New-Object System.Drawing.Size(200, 40)
$cleanupButton.Text = 'Execute Cleanup'
$cleanupButton.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$cleanupButton.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
$cleanupButton.ForeColor = [System.Drawing.Color]::White
$cleanupButton.FlatStyle = 'Flat'
$cleanupButton.Enabled = $false
$cleanupGroup.Controls.Add($cleanupButton)

# Preview Space Gain Button
$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Location = New-Object System.Drawing.Point(240, 130)
$previewButton.Size = New-Object System.Drawing.Size(200, 40)
$previewButton.Text = 'Preview Space Gain'
$previewButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$previewButton.Enabled = $false
$cleanupGroup.Controls.Add($previewButton)

# Progress GroupBox
$progressGroup = New-Object System.Windows.Forms.GroupBox
$progressGroup.Location = New-Object System.Drawing.Point(10, 390)
$progressGroup.Size = New-Object System.Drawing.Size(660, 120)
$progressGroup.Text = 'Progress'
$progressGroup.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($progressGroup)

# Progress TextBox
$progressBox = New-Object System.Windows.Forms.TextBox
$progressBox.Location = New-Object System.Drawing.Point(10, 25)
$progressBox.Size = New-Object System.Drawing.Size(640, 85)
$progressBox.Multiline = $true
$progressBox.ScrollBars = 'Vertical'
$progressBox.ReadOnly = $true
$progressBox.Font = New-Object System.Drawing.Font('Consolas', 8)
$progressBox.BackColor = [System.Drawing.Color]::Black
$progressBox.ForeColor = [System.Drawing.Color]::Lime
$progressGroup.Controls.Add($progressBox)

# Close Button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(560, 520)
$closeButton.Size = New-Object System.Drawing.Size(110, 35)
$closeButton.Text = 'Close'
$closeButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($closeButton)
$form.CancelButton = $closeButton

# Functions
function Write-Progress {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $progressBox.AppendText("[$timestamp] $Message`r`n")
    $progressBox.SelectionStart = $progressBox.TextLength
    $progressBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-EFIPartitions {
    Write-Progress "Scanning for System Reserved/EFI partitions..."
    $partitions = @()
    
    # Get all partitions
    $allPartitions = Get-Partition | Where-Object { 
        $_.Type -eq 'System' -or 
        $_.Type -eq 'Recovery' -or 
        ($_.Size -lt 1GB -and $_.DriveLetter -eq $null)
    }
    
    foreach ($partition in $allPartitions) {
        $disk = Get-Disk -Number $partition.DiskNumber
        $volume = Get-Volume -Partition $partition -ErrorAction SilentlyContinue
        
        $info = [PSCustomObject]@{
            DiskNumber      = $partition.DiskNumber
            PartitionNumber = $partition.PartitionNumber
            Size            = [math]::Round($partition.Size / 1MB, 2)
            Type            = $partition.Type
            DriveLetter     = $partition.DriveLetter
            HasLetter       = $partition.DriveLetter -ne $null
            FriendlyName    = if ($volume) { $volume.FileSystemLabel } else { "Unlabeled" }
        }
        
        $partitions += $info
    }
    
    Write-Progress "Found $($partitions.Count) potential partition(s)"
    return $partitions
}

function Assign-DriveLetter {
    param(
        [object]$Partition,
        [string]$SpecificLetter
    )
    
    Write-Progress "Assigning drive letter to partition..."
    
    # Get available drive letters if none specified
    $usedLetters = (Get-Volume).DriveLetter
    $availableLetters = 65..90 | ForEach-Object { [char]$_ } | Where-Object { $_ -notin $usedLetters }
    
    if ($availableLetters.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No available drive letters!", "Error", 'OK', 'Error')
        return $null
    }
    
    $letter = if ($SpecificLetter) { $SpecificLetter } else { $availableLetters[0] }
    
    # Verify the specific letter is actually available (race condition check)
    if ($SpecificLetter -and $SpecificLetter -in $usedLetters) {
        [System.Windows.Forms.MessageBox]::Show("Drive letter $SpecificLetter is already in use!", "Error", 'OK', 'Error')
        return $null
    }
    
    try {
        # Method 1: PowerShell Set-Partition
        Write-Progress "Attempting assignment via Set-Partition..."
        Get-Partition -DiskNumber $Partition.DiskNumber -PartitionNumber $Partition.PartitionNumber | 
        Set-Partition -NewDriveLetter $letter -ErrorAction Stop
        
        Write-Progress "Assigned drive letter: $letter`:"
        Start-Sleep -Seconds 2
        return $letter
    }
    catch {
        Write-Progress "PowerShell assignment failed: $($_.Exception.Message). Trying DiskPart..."
        
        # Method 2: DiskPart Fallback (More reliable for hidden/system partitions)
        try {
            $diskPartScript = "select disk $($Partition.DiskNumber)`nselect partition $($Partition.PartitionNumber)`nassign letter=$letter`nexit"
            $diskPartScript | Out-File -FilePath "$env:TEMP\diskpart_script.txt" -Encoding ASCII -Force
            
            $process = Start-Process -FilePath "diskpart.exe" -ArgumentList "/s `"$env:TEMP\diskpart_script.txt`"" -Wait -PassThru -WindowStyle Hidden
            
            if ($process.ExitCode -eq 0) {
                Write-Progress "DiskPart assignment successful: $letter`:"
                Start-Sleep -Seconds 2
                return $letter
            }
            else {
                throw "DiskPart exited with code $($process.ExitCode)"
            }
        }
        catch {
            Write-Progress "ERROR: Failed to assign drive letter - $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Failed to assign drive letter via both methods.`nError: $($_.Exception.Message)", "Error", 'OK', 'Error')
            return $null
        }
        finally {
            if (Test-Path "$env:TEMP\diskpart_script.txt") { Remove-Item "$env:TEMP\diskpart_script.txt" -Force }
        }
    }
}

function Unassign-DriveLetter {
    param(
        [object]$Partition
    )
    
    Write-Progress "Removing drive letter from partition..."
    
    if (-not $Partition.HasLetter) {
        [System.Windows.Forms.MessageBox]::Show("This partition does not have a drive letter assigned.", "Info", 'OK', 'Information')
        return $false
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will remove the drive letter $($Partition.DriveLetter): from this partition.`n`nThe partition will return to its normal hidden state.`n`nContinue?",
        "Confirm Unassign",
        'YesNo',
        'Question'
    )
    
    if ($result -ne 'Yes') {
        Write-Progress "Unassign operation cancelled by user"
        return $false
    }
    
    try {
        # Method 1: PowerShell Remove-PartitionAccessPath
        Write-Progress "Attempting removal via Remove-PartitionAccessPath..."
        $accessPath = "$($Partition.DriveLetter):"
        Get-Partition -DiskNumber $Partition.DiskNumber -PartitionNumber $Partition.PartitionNumber | 
        Remove-PartitionAccessPath -AccessPath $accessPath -ErrorAction Stop
        
        Write-Progress "Removed drive letter: $($Partition.DriveLetter):"
        Start-Sleep -Seconds 2
        return $true
    }
    catch {
        Write-Progress "PowerShell removal failed: $($_.Exception.Message). Trying DiskPart..."
        
        # Method 2: DiskPart Fallback (More reliable for hidden/system partitions)
        try {
            $diskPartScript = "select disk $($Partition.DiskNumber)`nselect partition $($Partition.PartitionNumber)`nremove letter=$($Partition.DriveLetter)`nexit"
            $diskPartScript | Out-File -FilePath "$env:TEMP\diskpart_script.txt" -Encoding ASCII -Force
            
            $process = Start-Process -FilePath "diskpart.exe" -ArgumentList "/s `"$env:TEMP\diskpart_script.txt`"" -Wait -PassThru -WindowStyle Hidden
            
            if ($process.ExitCode -eq 0) {
                Write-Progress "DiskPart removal successful"
                Start-Sleep -Seconds 2
                return $true
            }
            else {
                throw "DiskPart exited with code $($process.ExitCode)"
            }
        }
        catch {
            Write-Progress "ERROR: Failed to remove drive letter - $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Failed to remove drive letter via both methods.`nError: $($_.Exception.Message)", "Error", 'OK', 'Error')
            return $false
        }
        finally {
            if (Test-Path "$env:TEMP\diskpart_script.txt") { Remove-Item "$env:TEMP\diskpart_script.txt" -Force }
        }
    }
}

function Get-PartitionSpace {
    param([string]$DriveLetter)
    
    try {
        $drive = Get-PSDrive -Name $DriveLetter -ErrorAction Stop
        return @{
            TotalSize = [math]::Round($drive.Used + $drive.Free, 2)
            FreeSpace = [math]::Round($drive.Free, 2)
            UsedSpace = [math]::Round($drive.Used, 2)
        }
    }
    catch {
        return $null
    }
}

function Get-FolderSize {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) { return 0 }
    
    try {
        $size = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        return if ($size) { [math]::Round($size / 1MB, 2) } else { 0 }
    }
    catch {
        return 0
    }
}

function Preview-SpaceGain {
    if (-not $script:selectedDrive) { return }
    
    $driveLetter = $script:selectedDrive
    $totalGain = 0
    $details = @()
    
    Write-Progress "Calculating potential space gain..."
    
    if ($fontCheckbox.Checked) {
        $fontPath = "${driveLetter}:\EFI\Microsoft\Boot\Fonts"
        $fontSize = Get-FolderSize -Path $fontPath
        $totalGain += $fontSize
        $details += "Fonts: $fontSize MB"
        Write-Progress "  Fonts folder: $fontSize MB"
    }
    
    if ($langCheckbox.Checked) {
        $langFolders = @('jp', 'cz', 'dk', 'de', 'gr', 'es', 'hr', 'it', 'kr', 'lt', 'lv', 'no', 'nl', 'pl', 'ro', 'ru', 'sk', 'se', 'br', 'tw', 'pt')
        $basePath = "${driveLetter}:\EFI\Microsoft\Boot"
        
        $langSize = 0
        foreach ($lang in $langFolders) {
            $langPath = Join-Path $basePath $lang
            $size = Get-FolderSize -Path $langPath
            $langSize += $size
        }
        
        $totalGain += $langSize
        $details += "Language folders: $langSize MB"
        Write-Progress "  Language folders: $langSize MB"
    }
    
    Write-Progress "TOTAL ESTIMATED GAIN: $totalGain MB"
    
    $message = "Estimated space to be freed:`n`n" + ($details -join "`n") + "`n`nTotal: $totalGain MB"
    [System.Windows.Forms.MessageBox]::Show($message, "Space Gain Preview", 'OK', 'Information')
}

function Execute-Cleanup {
    $driveLetter = $script:selectedDrive
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will permanently delete selected files from the EFI partition.`n`nAre you sure you want to continue?",
        "Confirm Cleanup",
        'YesNo',
        'Warning'
    )
    
    if ($result -ne 'Yes') {
        Write-Progress "Cleanup cancelled by user"
        return
    }
    
    Write-Progress "=== STARTING CLEANUP OPERATION ==="
    $script:deletedItems = @()
    $spaceBefore = Get-PartitionSpace -DriveLetter $driveLetter
    
    # Backup if requested
    if ($backupCheckbox.Checked) {
        Write-Progress "Creating backup..."
        $backupPath = "C:\EFI_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        try {
            New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
            
            if ($fontCheckbox.Checked) {
                $fontPath = "${driveLetter}:\EFI\Microsoft\Boot\Fonts"
                if (Test-Path $fontPath) {
                    Copy-Item -Path $fontPath -Destination "$backupPath\Fonts" -Recurse -Force
                    Write-Progress "  Backed up Fonts folder"
                }
            }
            
            if ($langCheckbox.Checked) {
                $langFolders = @('jp', 'cz', 'dk', 'de', 'gr', 'es', 'hr', 'it', 'kr', 'lt', 'lv', 'no', 'nl', 'pl', 'ro', 'ru', 'sk', 'se', 'br', 'tw', 'pt')
                $basePath = "${driveLetter}:\EFI\Microsoft\Boot"
                
                foreach ($lang in $langFolders) {
                    $langPath = Join-Path $basePath $lang
                    if (Test-Path $langPath) {
                        Copy-Item -Path $langPath -Destination "$backupPath\$lang" -Recurse -Force
                    }
                }
                Write-Progress "  Backed up language folders"
            }
            
            Write-Progress "Backup created at: $backupPath"
        }
        catch {
            Write-Progress "WARNING: Backup failed - $($_.Exception.Message)"
            $continueResult = [System.Windows.Forms.MessageBox]::Show(
                "Backup failed. Continue with cleanup anyway?",
                "Backup Warning",
                'YesNo',
                'Warning'
            )
            if ($continueResult -ne 'Yes') {
                Write-Progress "Cleanup aborted"
                return
            }
        }
    }
    
    # Delete fonts
    if ($fontCheckbox.Checked) {
        Write-Progress "Removing boot fonts..."
        $fontPath = "${driveLetter}:\EFI\Microsoft\Boot\Fonts"
        
        if (Test-Path $fontPath) {
            try {
                $files = Get-ChildItem -Path $fontPath -File -Force
                foreach ($file in $files) {
                    Remove-Item -Path $file.FullName -Force
                    $script:deletedItems += $file.FullName
                }
                Write-Progress "  Deleted $($files.Count) font file(s)"
            }
            catch {
                Write-Progress "  ERROR: $($_.Exception.Message)"
            }
        }
        else {
            Write-Progress "  Font path not found"
        }
    }
    
    # Delete language folders
    if ($langCheckbox.Checked) {
        Write-Progress "Removing unused language folders..."
        $langFolders = @('jp', 'cz', 'dk', 'de', 'gr', 'es', 'hr', 'it', 'kr', 'lt', 'lv', 'no', 'nl', 'pl', 'ro', 'ru', 'sk', 'se', 'br', 'tw', 'pt')
        $basePath = "${driveLetter}:\EFI\Microsoft\Boot"
        $deletedCount = 0
        
        foreach ($lang in $langFolders) {
            $langPath = Join-Path $basePath $lang
            
            if (Test-Path $langPath) {
                try {
                    Remove-Item -Path $langPath -Recurse -Force
                    $script:deletedItems += $langPath
                    $deletedCount++
                    Write-Progress "  Deleted: $lang"
                }
                catch {
                    Write-Progress "  ERROR deleting $lang - $($_.Exception.Message)"
                }
            }
        }
        
        Write-Progress "  Deleted $deletedCount language folder(s)"
    }
    
    # Show results
    Start-Sleep -Seconds 1
    $spaceAfter = Get-PartitionSpace -DriveLetter $driveLetter
    $spaceFreed = [math]::Round($spaceAfter.FreeSpace - $spaceBefore.FreeSpace, 2)
    
    Write-Progress "=== CLEANUP COMPLETE ==="
    Write-Progress "Space freed: $spaceFreed MB"
    Write-Progress "Free space now: $($spaceAfter.FreeSpace) MB / $($spaceAfter.TotalSize) MB"
    
    $script:currentSpace = $spaceAfter.FreeSpace
    $spaceLabel.Text = "Free: $($spaceAfter.FreeSpace) MB / $($spaceAfter.TotalSize) MB (Freed: $spaceFreed MB)"
    $spaceLabel.ForeColor = [System.Drawing.Color]::Green
    
    [System.Windows.Forms.MessageBox]::Show(
        "Cleanup completed successfully!`n`nSpace freed: $spaceFreed MB`nFree space: $($spaceAfter.FreeSpace) MB",
        "Success",
        'OK',
        'Information'
    )
}

# Event Handlers
$scanButton.Add_Click({
        $statusLabel.Text = 'Scanning partitions...'
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $driveCombo.Items.Clear()
        $driveCombo.Enabled = $false
        $assignButton.Enabled = $false
    
        $partitions = Get-EFIPartitions
    
        if ($partitions.Count -eq 0) {
            $statusLabel.Text = 'No System Reserved/EFI partitions found'
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            Write-Progress "No suitable partitions found"
            return
        }
    
        foreach ($partition in $partitions) {
            $displayText = "Disk $($partition.DiskNumber) Part $($partition.PartitionNumber) - $($partition.Size) MB"
        
            if ($partition.HasLetter) {
                $displayText += " [$($partition.DriveLetter):]"
            }
            else {
                $displayText += " [No Letter]"
            }
        
            $displayText += " - $($partition.Type)"
        
            $item = New-Object System.Windows.Forms.ListViewItem
            $item.Text = $displayText
            $item.Tag = $partition
        
            $driveCombo.Items.Add($item) | Out-Null
        }
    
        $driveCombo.Enabled = $true
        if ($driveCombo.Items.Count -gt 0) {
            $driveCombo.SelectedIndex = 0
        }
    
        # Populate available letters
        $letterCombo.Items.Clear()
        $usedLetters = (Get-Volume).DriveLetter
        $availableLetters = 65..90 | ForEach-Object { [char]$_ } | Where-Object { $_ -notin $usedLetters }
        foreach ($l in $availableLetters) {
            $letterCombo.Items.Add($l) | Out-Null
        }
        if ($letterCombo.Items.Count -gt 0) {
            # Default to 'Z' if available, else first one
            if ($letterCombo.Items.Contains('Z')) { $letterCombo.SelectedItem = 'Z' }
            else { $letterCombo.SelectedIndex = 0 }
        }
    
        $statusLabel.Text = "Found $($partitions.Count) partition(s). Select one to continue."
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    })

$driveCombo.Add_SelectedIndexChanged({
        if ($driveCombo.SelectedItem -eq $null) { return }
    
        $partition = $driveCombo.SelectedItem.Tag
    
        if ($partition.HasLetter) {
            $script:selectedDrive = $partition.DriveLetter
            $letterCombo.Enabled = $false
        
            $space = Get-PartitionSpace -DriveLetter $script:selectedDrive
            if ($space) {
                $assignButton.Enabled = $false
                $assignButton.Text = "Assign Letter"
                $unassignButton.Enabled = $true
            
                $script:originalSpace = $space.FreeSpace
                $script:currentSpace = $space.FreeSpace
                $spaceLabel.Text = "Free: $($space.FreeSpace) MB / $($space.TotalSize) MB"
                $spaceLabel.ForeColor = if ($space.FreeSpace -lt 50) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }
            
                $fontCheckbox.Enabled = $true
                $langCheckbox.Enabled = $true
                $backupCheckbox.Enabled = $true
                $cleanupButton.Enabled = $true
                $previewButton.Enabled = $true
            
                $statusLabel.Text = "Partition ready for cleanup"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
            
                Write-Progress "Selected partition: $($script:selectedDrive): - Free space: $($space.FreeSpace) MB"
            }
            else {
                # Drive has letter but is inaccessible - Allow fixing
                $assignButton.Enabled = $true
                $assignButton.Text = "Fix Mapping"
                $letterCombo.Enabled = $true
                $unassignButton.Enabled = $true
            
                $fontCheckbox.Enabled = $false
                $langCheckbox.Enabled = $false
                $backupCheckbox.Enabled = $false
                $cleanupButton.Enabled = $false
                $previewButton.Enabled = $false
            
                $spaceLabel.Text = "Drive inaccessible - Try fixing mapping"
                $spaceLabel.ForeColor = [System.Drawing.Color]::Red
            
                $statusLabel.Text = "Error accessing drive $($script:selectedDrive):. Click 'Fix Drive Mapping'."
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        else {
            $script:selectedDrive = $null
            $assignButton.Enabled = $true
            $assignButton.Text = "Assign Letter"
            $letterCombo.Enabled = $true
            $unassignButton.Enabled = $false
        
            $fontCheckbox.Enabled = $false
            $langCheckbox.Enabled = $false
            $backupCheckbox.Enabled = $false
            $cleanupButton.Enabled = $false
            $previewButton.Enabled = $false
        
            $spaceLabel.Text = "No drive letter assigned - click 'Assign Drive Letter'"
            $spaceLabel.ForeColor = [System.Drawing.Color]::Orange
        
            $statusLabel.Text = "Partition needs a drive letter"
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        }
    })

$assignButton.Add_Click({
        if ($driveCombo.SelectedItem -eq $null) { return }
    
        $partition = $driveCombo.SelectedItem.Tag
        $selectedLetter = $letterCombo.SelectedItem
    
        if (-not $selectedLetter) {
            [System.Windows.Forms.MessageBox]::Show("Please select a drive letter first.", "Error", 'OK', 'Error')
            return
        }
    
        $letter = Assign-DriveLetter -Partition $partition -SpecificLetter $selectedLetter
    
        if ($letter) {
            # Refresh the partition list
            $scanButton.PerformClick()
        
            # Select the newly assigned partition
            for ($i = 0; $i -lt $driveCombo.Items.Count; $i++) {
                if ($driveCombo.Items[$i].Text -like "*[$letter`:]*") {
                    $driveCombo.SelectedIndex = $i
                    break
                }
            }
        }
    })

$unassignButton.Add_Click({
        if ($driveCombo.SelectedItem -eq $null) { return }
    
        $partition = $driveCombo.SelectedItem.Tag
    
        if (-not $partition.HasLetter) {
            [System.Windows.Forms.MessageBox]::Show("This partition does not have a drive letter assigned.", "Info", 'OK', 'Information')
            return
        }
    
        $success = Unassign-DriveLetter -Partition $partition
    
        if ($success) {
            # Refresh the partition list
            $scanButton.PerformClick()
        
            # Try to select the same partition (now without a letter)
            for ($i = 0; $i -lt $driveCombo.Items.Count; $i++) {
                $item = $driveCombo.Items[$i]
                if ($item.Tag.DiskNumber -eq $partition.DiskNumber -and $item.Tag.PartitionNumber -eq $partition.PartitionNumber) {
                    $driveCombo.SelectedIndex = $i
                    break
                }
            }
        }
    })

$previewButton.Add_Click({
        Preview-SpaceGain
    })

$cleanupButton.Add_Click({
        Execute-Cleanup
    })

$closeButton.Add_Click({
        if ($script:deletedItems.Count -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Cleanup operations were performed. Are you sure you want to exit?",
                "Confirm Exit",
                'YesNo',
                'Question'
            )
        
            if ($result -eq 'Yes') {
                $form.Close()
            }
        }
        else {
            $form.Close()
        }
    })

# Show form
Write-Progress "EFI Cleanup Tool initialized. Ready for operation."
$form.ShowDialog() | Out-Null