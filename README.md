
# Wake-on-LAN Sender Script
![image](https://github.com/user-attachments/assets/bd06bf8b-42d0-4649-a6f1-d29a3d693bb0)


## Overview

This PowerShell script allows you to send Wake-on-LAN (WOL) packets to devices based on a list of MAC addresses provided in either an Excel file or a text file. It also includes a feature to send WOL packets to a predefined set of devices using custom lists.

## Features

- **Send WOL packets** to multiple devices using a list of MAC addresses.
- **Support for both Excel and text file formats**.
- **Customizable list** for specific device groups.
- **User-friendly GUI** for selecting files and sending packets.

## Prerequisites

- **PowerShell**: This script requires PowerShell to run.
- **ImportExcel Module**: Install the `ImportExcel` module to handle Excel files.

## Script Components

### 1. **Module and Assembly Imports**

```powershell
# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
```

**Explanation**: 
- **`Import-Module ImportExcel`**: Loads the ImportExcel module to handle Excel files.
- **`Add-Type`**: Adds .NET types for creating the Windows Forms and drawing components.

### 2. **Path to Icon File**

```powershell
# Path to your icon file
$iconPath = "C:\Users\IT\Pictures\hotel_logo_OLED.ico"
```

**Explanation**:
- **`$iconPath`**: Path to the icon file used in the GUI.

### 3. **Function to Send WOL Packet**

```powershell
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
```

**Explanation**:
- **`Send-WOL` Function**: Sends a WOL packet to a specified MAC address.
- **`$mac`**: Formats the MAC address.
- **`$packet`**: Constructs the WOL packet.
- **`$client.Send`**: Sends the packet over UDP.

### 4. **Create and Configure the Form**

```powershell
# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wake-on-LAN Sender"
$form.Size = New-Object System.Drawing.Size(400,300)

# Load the icon and set it for the form
if (Test-Path $iconPath) {
    $icon = New-Object System.Drawing.Icon($iconPath)
    $form.Icon = $icon
} else {
    Write-Host "Icon file not found at $iconPath"
}
```

**Explanation**:
- **`$form`**: Creates the main form for the GUI.
- **`$icon`**: Sets the form icon if the file is found.

### 5. **Create GUI Controls**

```powershell
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
```

**Explanation**:
- **`$label`**: Displays instruction text on the form.
- **`$textBox`**: Allows users to enter or browse for the file path.
- **`$browseButton`**: Opens a file dialog to select a file.
- **`$sendButton`**: Executes the WOL sending operation.

### 6. **Show the Form**

```powershell
# Show Form
$form.ShowDialog()
$form.Dispose()
```

**Explanation**:
- **`$form.ShowDialog()`**: Displays the form and keeps it open until closed.
- **`$form.Dispose()`**: Cleans up resources when the form is closed.

## Customizing the Script

1. **Change Icon**: Update the `$iconPath` variable to point to a new icon file.
2. **Update File Paths**: Modify the paths for the Excel and text files in the script.
3. **Modify GUI Elements**: Adjust the positions and sizes of controls by changing the `Location` and `Size` properties.

## Running the Script

1. **Open PowerShell**: Run PowerShell as Administrator.
2. **Execute Script**: Use the command `PowerShell -ExecutionPolicy Bypass -File "C:\Path\To\YourScript.ps1"` to execute the
3. **Right Click on the Scrip: Run with powershell

 script.

## Scheduling the Script

To run the script automatically at a specific time, use Task Scheduler:

1. Create a batch file (`Run_WOL_Script.bat`) to execute the PowerShell script:

    ```batch
    @echo off
    PowerShell -ExecutionPolicy Bypass -File "C:\Path\To\Scheduled_WOL.ps1"
    ```

2. Open Task Scheduler and create a new task:
   - **Trigger**: Set the desired time (e.g., 6 AM, Monday to Friday).
   - **Action**: Set the action to start the batch file created above.

## Troubleshooting

- **File Not Found**: Ensure the paths to the Excel or text files are correct.
- **Icon Issues**: Verify the path to the icon file is accurate.

For any additional help or modifications, feel free to fork the repository and make your own changes.

---

Feel free to modify or add additional details specific to your environment or requirements.
