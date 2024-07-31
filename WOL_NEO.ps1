# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Path to your icon file
$iconPath = "C:\Users\IT\Pictures\hotel_logo_OLED.ico"

# Custom file paths for predefined lists
$neoExcelFilePath = "C:\Users\IT\Downloads\Wake on Lan Script\Downloads\WOL_LIST_custom1.xlsx"
$neoTextFilePath = "C:\Users\IT\Downloads\Wake on Lan Script\Downloads\WOL_LIST_custom1.txt"

# Function to send WOL packet
function Send-WOL {
    param(
        [string]$MacAddress
    )

    $mac = $MacAddress.ToUpper() -replace '[^0-9A-F]',''

    [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
        $_.NetworkInterfaceType -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Wireless80211 -and 
        $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up
    } | ForEach-Object {
        $networkInterface = $_
        $localIpAddress = ($networkInterface.GetIPProperties().UnicastAddresses | Where-Object {
            $_.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork
        })[0].Address
        $targetPhysicalAddress = [System.Net.NetworkInformation.PhysicalAddress]::Parse($mac)
        $targetPhysicalAddressBytes = $targetPhysicalAddress.GetAddressBytes()
        $packet = [byte[]](,0xFF * 102)
        6..101 | ForEach-Object { $packet[$_] = $targetPhysicalAddressBytes[($_ % 6)] }
        $localEndpoint = [System.Net.IPEndPoint]::new($localIpAddress, 0)
        $targetEndpoint = [System.Net.IPEndPoint]::new([System.Net.IPAddress]::Broadcast, 9)
        $client = [System.Net.Sockets.UdpClient]::new($localEndpoint)
        try {
            $client.Send($packet, $packet.Length, $targetEndpoint) | Out-Null
            Write-Host "Sent WOL packet to $MacAddress from $localIpAddress"
        } finally {
            $client.Dispose()
        }
    }
}

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wake-on-LAN Sender"
$form.Size = New-Object System.Drawing.Size(400,350)

# Load the icon and set it for the form
if (Test-Path $iconPath) {
    $icon = New-Object System.Drawing.Icon($iconPath)
    $form.Icon = $icon
} else {
    Write-Host "Icon file not found at $iconPath"
}

# Create Controls
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(380,20)
$label.Text = "Select a file containing device list (Excel or Text):"
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,50)
$textBox.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($textBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(320,50)
$browseButton.Size = New-Object System.Drawing.Size(60,20)
$browseButton.Text = "Browse"
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Text Files (*.txt)|*.txt"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($browseButton)

$sendButton = New-Object System.Windows.Forms.Button
$sendButton.Location = New-Object System.Drawing.Point(10,100)
$sendButton.Size = New-Object System.Drawing.Size(100,30)
$sendButton.Text = "Send WOL"
$sendButton.Add_Click({
    $filePath = $textBox.Text
    if (-not (Test-Path $filePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        if ($filePath.EndsWith(".xlsx")) {
            # Handle Excel file
            $data = Import-Excel -Path $filePath
            if ($data.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No data found in the file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $data | ForEach-Object {
                $mac = $_.'MAC address'
                if ($mac) {
                    Write-Host "Sending WOL to MAC: $mac"
                    Send-WOL -MacAddress $mac
                } else {
                    Write-Host "MAC address missing for one of the entries."
                }
            }
        } elseif ($filePath.EndsWith(".txt")) {
            # Handle Text file
            $macAddresses = Get-Content -Path $filePath
            if ($macAddresses.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No data found in the file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $macAddresses | ForEach-Object {
                $mac = $_.Trim()
                if ($mac) {
                    Write-Host "Sending WOL to MAC: $mac"
                    Send-WOL -MacAddress $mac
                } else {
                    Write-Host "MAC address missing in one of the lines."
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Unsupported file type. Please select an Excel or Text file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        [System.Windows.Forms.MessageBox]::Show("WOL packets sent successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to read file or send WOL packets. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($sendButton)

# Create NEO PC Button
$neoButton = New-Object System.Windows.Forms.Button
$neoButton.Location = New-Object System.Drawing.Point(10,150)
$neoButton.Size = New-Object System.Drawing.Size(100,30)
$neoButton.Text = "NEO PC"
$neoButton.Add_Click({
    if (Test-Path $neoExcelFilePath) {
        $filePath = $neoExcelFilePath
    } elseif (Test-Path $neoTextFilePath) {
        $filePath = $neoTextFilePath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Custom list files not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $textBox.Text = $filePath
    $sendButton.PerformClick()
})
$form.Controls.Add($neoButton)

# Show Form
$form.ShowDialog()
$form.Dispose()
