# Import the Excel module
Import-Module ImportExcel

# OU Paths
$ou1 = "OU=OU1,DC=domain,DC=com"
$ou2 = "OU=OU2,DC=domain,DC=com"
$ou3 = "OU=OU3,DC=domain,DC=com"
$ou4 = "OU=OU4,DC=domain,DC=com"

# Define IP patterns and middle numbers for each OU
$ou1IPs = @("10.10.2", "10.23.23")
$ou1MiddleNumbers = @("2112", "3101")

$ou2IPs = @("10.10.3", "10.23.45")
$ou2MiddleNumbers = @("3100", "4102")

$ou3IPs = @("10.10.4", "10.23.60")
$ou3MiddleNumbers = @("2113", "5103")

# Report Paths
$excelPath = "C:\Reports\NotMovedComputers.xlsx"
$htmlPath = "C:\Reports\MovedComputers.html"
$failedComputers = @()
$successComputers = @()

# Get all computers from AD
$computers = Get-ADComputer -Filter * -Properties Name, IPv4Address, DistinguishedName, OperatingSystem

# Process each computer
foreach ($computer in $computers) {
    $name = $computer.Name
    $ip = $computer.IPv4Address
    $dn = $computer.DistinguishedName
    $os = $computer.OperatingSystem

    # Skip if name contains restricted keywords or OS contains "Server"
    if ($name -match "super|wow|best" -or $os -match "Server") {
        Write-Output "$name skipped due to name or OS restrictions."
        continue
    }

    # Extract middle number (e.g., "2112" from "san-2112-f232")
    $middleNumber = if ($name -match "^[a-z]+-(\d+)-[a-z0-9]+$") { $matches[1] } else { $null }

    # Get first three octets of IP
    $ipPrefix = if ($ip) { ($ip -split "\.")[0..2] -join "." } else { $null }

    # Initialize OU variable
    $targetOU = $null

    # Determine Target OU
    if ($ipPrefix -in $ou1IPs -or $middleNumber -in $ou1MiddleNumbers) {
        $targetOU = $ou1
    }
    elseif ($ipPrefix -in $ou2IPs -or $middleNumber -in $ou2MiddleNumbers) {
        $targetOU = $ou2
    }
    elseif ($ipPrefix -in $ou3IPs -or $middleNumber -in $ou3MiddleNumbers) {
        $targetOU = $ou3
    }
    if (-not $targetOU) {
        $failedComputers += [PSCustomObject]@{
            ComputerName = $name
            IPAddress    = $ip
            MiddleNumber = $middleNumber
            ErrorMessage = "No matching OU found"
        }
        continue
    }

    # Try to move the computer and catch any errors
    try {
        Move-ADObject -Identity $dn -TargetPath $targetOU -ErrorAction Stop
        $successComputers += [PSCustomObject]@{
            ComputerName = $name
            IPAddress    = $ip
            MiddleNumber = $middleNumber
            MovedToOU    = $targetOU
            Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        Write-Output "$name moved to $targetOU"
    }
    catch {
        $failedComputers += [PSCustomObject]@{
            ComputerName = $name
            IPAddress    = $ip
            MiddleNumber = $middleNumber
            ErrorMessage = $_.Exception.Message
            Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        Write-Output "Failed to move $name : $($_.Exception.Message)"
    }
}

# Export Failed Computers to Excel
if ($failedComputers.Count -gt 0) {
    $failedComputers | Export-Excel -Path $excelPath -WorksheetName "FailedComputers" -AutoSize -Append
    Write-Output "Failed computers report written to $excelPath"
}

# Generate HTML Report for Moved Computers
if ($successComputers.Count -gt 0) {
    $htmlHeader = @"
    <html>
    <head>
        <title>Moved Computers Report</title>
        <style>
            table {border-collapse: collapse; width: 100%;}
            th, td {border: 1px solid #ddd; padding: 8px; text-align: left;}
            th {background-color: #4CAF50; color: white;}
            tr:nth-child(even) {background-color: #f2f2f2;}
        </style>
    </head>
    <body>
        <h2>Moved Computers Report - $(Get-Date -Format "yyyy-MM-dd")</h2>
        <table>
            <tr>
                <th>Computer Name</th>
                <th>IP Address</th>
                <th>Middle Number</th>
                <th>Moved To OU</th>
                <th>Timestamp</th>
            </tr>
"@

    $htmlBody = $successComputers | ForEach-Object {
        "<tr><td>$($_.ComputerName)</td><td>$($_.IPAddress)</td><td>$($_.MiddleNumber)</td><td>$($_.MovedToOU)</td><td>$($_.Timestamp)</td></tr>"
    }

    $htmlFooter = @"
        </table>
    </body>
    </html>
"@

    # Combine HTML sections
    $htmlReport = $htmlHeader + ($htmlBody -join "`n") + $htmlFooter
    $htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Output "HTML report generated at $htmlPath"
} else {
    Write-Output "No computers were moved. HTML report not created."
}
