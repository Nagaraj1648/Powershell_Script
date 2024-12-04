# Define HTML report file path
$formattedDate = Get-Date -Format "ddMMyyyy_HHmm" # No need to modify
$htmlformatdate = Get-Date -Format "dd-MM-yyyy HH:mm" # No need to modify
$reportPath = "C:\Reports\ActiveDirectory_Account_Lockout_Report_$formattedDate.html" # Modify your own report save path


# Define the AD servers
$ADServers = @("Domain controller name")  # Replace with your actual server names

# Define the time frame for the logs
$startTime = (Get-Date).AddDays(-90)  # Adjust as needed
$endTime = Get-Date

# Initialize an array to hold all events
$allLockoutEvents = @()

# Loop through each AD server and collect the logs
foreach ($server in $ADServers) {
    try {
        $lockoutEvents = Get-WinEvent -ComputerName $server -FilterHashtable @{
            LogName   = 'Security'
            Id        = 4740
            StartTime = $startTime
            EndTime   = $endTime
        } | Select-Object @{
            Name       = 'Timestamp'
            Expression = {$_.TimeCreated}
        }, @{
            Name       = 'User ID'
            Expression = {$_.Properties[0].Value}
        }, @{
            Name       = 'Source PC'
            Expression = {$_.Properties[1].Value}
        }, @{
            Name       = 'Domain Controller'
            Expression = {$_.MachineName}
        }

        # Add the events to the master list
        $allLockoutEvents += $lockoutEvents
    } catch {
        Write-Warning "Failed to retrieve logs from $server : $_"
    }
}

# Convert the aggregated lockout events to HTML
$htmlOutput = @"
<!DOCTYPE html>
<html>
<head>
    <title>Account Lockout Events </title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            text-align: center; /* Center-align the entire body content */
        }
        h1 {
            color: #2c3e50;
            text-align: center;
        }
        .table-container {
            display: inline-block; /* Center the table within the page */
            text-align: left; /* Ensure text inside the table is left-aligned */
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 18px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #ddd;
        }
    </style>
</head>
<body>
    <h1>Account Lockout Events $htmlformatdate</h1>
    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th>Timestamp</th>
                    <th>User ID</th>
                    <th>Source PC</th>
                    <th>Domain Controller</th>
                </tr>
            </thead>
            <tbody>
"@

# Append rows to the HTML table
foreach ($event in $allLockoutEvents) {
    $htmlOutput += @"
                <tr>
                    <td>$($event.Timestamp)</td>
                    <td>$($event.'User ID')</td>
                    <td>$($event.'Source PC')</td>
                    <td>$($event.'Domain Controller')</td>
                </tr>
"@
}

# Close the HTML tags
$htmlOutput += @"
            </tbody>
        </table>
    </div>
</body>
</html>
"@


# Save the HTML output to a file
$htmlOutput | Out-File -FilePath $reportPath -Encoding utf8

# Optionally, open the HTML file in a web browser
Start-Process $reportPath
