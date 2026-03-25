# 1. Load Windows Forms & Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- NOTIFICATION FUNCTION ---
function Show-RigNotification ($Message) {
    $notif = New-Object Windows.Forms.NotifyIcon
    $path = (Get-Process -id $PID).Path
    $notif.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $notif.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $notif.BalloonTipTitle = "RigReport"
    $notif.BalloonTipText = $Message
    $notif.Visible = $true
    $notif.ShowBalloonTip(2000)
    $cleanupTimer = New-Object Windows.Forms.Timer
    $cleanupTimer.Interval = 3000 
    $cleanupTimer.Add_Tick({
        $this.Stop()
        $notif.Visible = $false
        $notif.Dispose()
        $this.Dispose()
    }.GetNewClosure())
    $cleanupTimer.Start()
}

# --- FIXED EXPORT FORMATTER ---
function Get-RigSpecSheet ($Data) {
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("==================================================") | Out-Null
    $sb.AppendLine("RigReport Summary") | Out-Null
    $sb.AppendLine("Generated on: $(Get-Date -Format 'MM-dd-yyyy HH:mm:ss')") | Out-Null
    $sb.AppendLine("==================================================") | Out-Null
    $sb.AppendLine("") | Out-Null
    
    foreach ($prop in $Data.PSObject.Properties) {
        $label = "$($prop.Name) :".PadRight(25)
        $value = if ($prop.Value) { $prop.Value.ToString() } else { "N/A" }
        $sb.AppendLine("$label $value") | Out-Null
    }
    
    $sb.AppendLine("") | Out-Null
    $sb.Append("==================================================") | Out-Null
    return $sb.ToString()
}

# 2. Gather Data
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
$bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
$bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
$ramData = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
$memArray = Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue

$ramCount = ($ramData | Measure-Object).Count
$totalSlots = $memArray.MemoryDevices
$ramTotalGB = [Math]::Round(($ramData | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$ramSpeed = ($ramData.ConfiguredClockSpeed | Select-Object -First 1)
$ramVendors = ($ramData.Manufacturer | Select-Object -Unique) -join " / "

$channelMode = switch ($ramCount) {
    1 { "Single Channel" }
    2 { "Dual Channel" }
    4 { "Quad Channel" }
    default { "$ramCount Sticks" }
}

$gpuNames = (Get-CimInstance Win32_VideoController).Caption -join ", "
$biosDate = if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString("MM-dd-yyyy") } else { "Unknown" }

$script:results = [PSCustomObject]@{
    'OS'                = $os.Caption
    'Rig Name'          = $env:COMPUTERNAME
    'CPU'               = $cpu.Name.Trim()
    'RAM (Total GB)'    = $ramTotalGB
    'RAM Config'        = "$ramCount / $totalSlots Slots ($channelMode)"
    'RAM Details'       = "$($ramVendors) ($($ramSpeed) MHz)"
    'GPU'               = $gpuNames
    'Free Drive Space'  = "$([Math]::Round($disk.FreeSpace / 1GB, 2)) GB"
    'Motherboard'       = "$($bb.Manufacturer) $($bb.Product)"
    'Board Revision'    = $bb.Version
    'BIOS Version'      = $bios.SMBIOSBIOSVersion
    'BIOS Date'         = $biosDate
    'Boot Mode'         = if ($env:Firmware_Type) { $env:Firmware_Type } else { "UEFI" }
}

# 3. Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "RigReport: $env:COMPUTERNAME"
$form.Size = New-Object System.Drawing.Size(550,550) 
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "RigReport v1.0.0"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size = New-Object System.Drawing.Size(400, 40)
$titleLabel.Location = New-Object System.Drawing.Point(25, 15)
$form.Controls.Add($titleLabel)

$listView = New-Object System.Windows.Forms.ListView
$listView.View = "Details"
$listView.HeaderStyle = "None"
$listView.Width = 490
$listView.Height = 350
$listView.Location = New-Object System.Drawing.Point(25, 60)
$listView.FullRowSelect = $true
$listView.BorderStyle = "FixedSingle"
$listView.Columns.Add("P", 160) | Out-Null
$listView.Columns.Add("V", 300) | Out-Null

foreach ($prop in $script:results.PSObject.Properties) {
    $item = New-Object System.Windows.Forms.ListViewItem($prop.Name)
    $valStr = if ($prop.Value) { $prop.Value.ToString() } else { "" }
    $item.SubItems.Add($valStr) | Out-Null
    $listView.Items.Add($item) | Out-Null
}
$listView.AutoResizeColumn(1, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
$form.Controls.Add($listView)

# --- ICON DRAWING LOGIC ---
function Get-SunImage ([System.Drawing.Color]$Color, [bool]$IsHollow) {
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen = New-Object System.Drawing.Pen($Color, 2)
    if ($IsHollow) { $g.DrawEllipse($pen, 10, 10, 12, 12) } 
    else { $brush = New-Object System.Drawing.SolidBrush($Color); $g.FillEllipse($brush, 10, 10, 12, 12) }
    $g.DrawLine($pen, 16, 4, 16, 8); $g.DrawLine($pen, 16, 24, 16, 28)
    $g.DrawLine($pen, 4, 16, 8, 16); $g.DrawLine($pen, 24, 16, 28, 16)
    $g.DrawLine($pen, 8, 8, 11, 11); $g.DrawLine($pen, 21, 21, 24, 24)
    $g.DrawLine($pen, 24, 8, 21, 11); $g.DrawLine($pen, 11, 21, 8, 24)
    $g.Dispose(); return $bmp
}

# --- THEME SWITCHER ---
$script:isDarkMode = $true 

$themeBtn = New-Object System.Windows.Forms.Button
$themeBtn.Size = New-Object System.Drawing.Size(40, 40)
$themeBtn.Location = New-Object System.Drawing.Point(475, 12)
$themeBtn.FlatStyle = "Flat"
$themeBtn.FlatAppearance.BorderSize = 0
$themeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand

function Update-Theme {
    if ($script:isDarkMode) {
        $form.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
        $titleLabel.ForeColor = [System.Drawing.Color]::White
        $listView.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
        $listView.ForeColor = [System.Drawing.Color]::White
        $themeBtn.Image = Get-SunImage -Color ([System.Drawing.Color]::White) -IsHollow $false
    } else {
        $form.BackColor = [System.Drawing.Color]::WhiteSmoke
        $titleLabel.ForeColor = [System.Drawing.Color]::Black
        $listView.BackColor = [System.Drawing.Color]::White
        $listView.ForeColor = [System.Drawing.Color]::Black
        $themeBtn.Image = Get-SunImage -Color ([System.Drawing.Color]::Black) -IsHollow $true
    }
}

$themeBtn.Add_Click({
    $script:isDarkMode = !$script:isDarkMode
    Update-Theme
})
$form.Controls.Add($themeBtn)

# --- ACTION BUTTONS ---
$copyBtn = New-Object System.Windows.Forms.Button
$copyBtn.Text = "Copy All"
$copyBtn.Size = New-Object System.Drawing.Size(235, 45)
$copyBtn.Location = New-Object System.Drawing.Point(25, 430)
$copyBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$copyBtn.BackColor = [System.Drawing.Color]::LightSlateGray
$copyBtn.ForeColor = [System.Drawing.Color]::White
$copyBtn.FlatStyle = "Flat"
$copyBtn.Add_Click({
    $txt = Get-RigSpecSheet -Data $script:results
    [System.Windows.Forms.Clipboard]::SetText($txt)
    Show-RigNotification -Message "Specs copied to clipboard!"
})
$form.Controls.Add($copyBtn)

$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "Export Specs"
$exportBtn.Size = New-Object System.Drawing.Size(235, 45)
$exportBtn.Location = New-Object System.Drawing.Point(280, 430)
$exportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$exportBtn.BackColor = [System.Drawing.Color]::DodgerBlue
$exportBtn.ForeColor = [System.Drawing.Color]::White
$exportBtn.FlatStyle = "Flat"
$exportBtn.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveDialog.FileName = "RigReport_$($env:COMPUTERNAME)"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txt = Get-RigSpecSheet -Data $script:results
        $txt | Set-Content -Path $saveDialog.FileName
        Show-RigNotification -Message "Report saved successfully!"
    }
})
$form.Controls.Add($exportBtn)

Update-Theme
$form.ShowDialog() | Out-Null