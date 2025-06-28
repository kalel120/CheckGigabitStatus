# Function to show a dialog box
Add-Type -AssemblyName System.Windows.Forms

function Show-SpeedDropAlert {
    param (
        [string]$AdapterName,
        [string]$CurrentSpeed
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Network Speed Alert"
    $form.Size = New-Object System.Drawing.Size(400,200)
    $form.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(380,100)
    $label.Text = "Speed drop detected on adapter: $AdapterName`nCurrent speed: $CurrentSpeed`n`nThe connection speed has dropped below 1 Gbps!"
    $form.Controls.Add($label)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150,120)
    $okButton.Size = New-Object System.Drawing.Size(100,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $form.Topmost = $true
    $form.ShowDialog()
}


# Auto-adjust PowerShell window size for better visibility
$host.UI.RawUI.WindowTitle = "Gigabit Status Monitor"
$host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size (120, 300)
$host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size (120, 40)

Write-Host "Starting Network Speed Monitor..."
Write-Host "Press Ctrl+C to stop monitoring."

while ($true) {
    # Get all enabled network adapters
    $activeAdapter = Get-NetAdapter | Where-Object {
        $_.Status -eq "Up" -and 
        $_.MediaType -eq "802.3" -and 
        $_.LinkSpeed -match "bps"
    } | Select-Object -First 1

    if ($activeAdapter) {
        $speed = $activeAdapter.LinkSpeed
        $adapterName = $activeAdapter.Name
        
        # Dynamically adjust window width to fit the output text
        $outputText = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $adapterName - Current Speed: $speed"
        $desiredWidth = [Math]::Max(80, [Math]::Min($outputText.Length + 10, 200))
        if ($host.UI.RawUI.WindowSize.Width -ne $desiredWidth) {
            $host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size ($desiredWidth, $host.UI.RawUI.BufferSize.Height)
            $host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size ($desiredWidth, $host.UI.RawUI.WindowSize.Height)
        }
        Write-Host $outputText
        [System.Console]::SetCursorPosition(0, [System.Console]::BufferHeight - 1)
        
        # Check if speed is below 1 Gbps
        if ($speed -match "(\d+)\s*(Gbps|Mbps)") {
            $value = [int]$matches[1]
            $unit = $matches[2]
            
            if (($unit -eq "Mbps" -and $value -le 100) -or ($unit -eq "Gbps" -and $value -lt 1)) {
                Show-SpeedDropAlert -AdapterName $adapterName -CurrentSpeed $speed
            }
        }
    } else {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): No active Ethernet adapter found."
        [System.Console]::SetCursorPosition(0, [System.Console]::BufferHeight - 1)
    }
    
    # Wait for 5 seconds before next check
    Start-Sleep -Seconds 2.5
}