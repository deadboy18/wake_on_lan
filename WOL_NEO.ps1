# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

# Custom file paths for predefined lists
$neoExcelFilePath = "C:\Users\IT\Downloads\Wake on Lan Script\Scheduled wake on lan\WOL_LIST_custom1.xlsx"
$neoTextFilePath = "C:\Users\IT\Downloads\Wake on Lan Script\Scheduled wake on lan\WOL_LIST_custom1.txt"

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

# Determine which file to use
if (Test-Path $neoExcelFilePath) {
    $filePath = $neoExcelFilePath
} elseif (Test-Path $neoTextFilePath) {
    $filePath = $neoTextFilePath
} else {
    Write-Host "Custom list files not found."
    exit 1
}

if ($filePath.EndsWith(".xlsx")) {
    # Handle Excel file
    $data = Import-Excel -Path $filePath
    $data | ForEach-Object {
        $mac = $_.'MAC address'
        if ($mac) {
            Send-WOL -MacAddress $mac
        }
    }
} elseif ($filePath.EndsWith(".txt")) {
    # Handle Text file
    $macAddresses = Get-Content -Path $filePath
    $macAddresses | ForEach-Object {
        $mac = $_.Trim()
        if ($mac) {
            Send-WOL -MacAddress $mac
        }
    }
} else {
    Write-Host "Unsupported file type. Please use an Excel or Text file."
    exit 1
}
