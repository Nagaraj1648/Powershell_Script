# Get all domain controllers in the domain
$DCs = Get-ADDomainController -Filter * 

# Start background jobs for each DC
$Jobs = @()
foreach ($dc in $DCs) {
    $Jobs += Start-Job -ScriptBlock {
        $Timesource = Invoke-Command -ComputerName $using:dc.Hostname -ScriptBlock { w32tm.exe /Query /source } 
        $Currenttime = Invoke-Command -ComputerName $using:dc.Hostname -ScriptBlock { Get-Date }
        [PSCustomObject]@{
            Domaincontroller = $using:dc.Hostname
            IPAddress        = $using:dc.IPv4Address
            Timesource       = $Timesource
            Currenttime      = $Currenttime
        }
    }
}

# Wait for all jobs to finish
$Jobs | Wait-Job

# Retrieve results
$result = $Jobs | Receive-Job

# Cleanup
$Jobs | Remove-Job

# Convert to HTML and save report (same as above)
$HtmlReport = $result | ConvertTo-Html -Property Domaincontroller, IPAddress, Timesource, Currenttime -Head "<Style>table {width: 100%; border-collapse: collapse; } th, td { border: 1px solid black; padding: text-align: left;} th { background-color: #f2f2f2;}</style>" -Title "Domain Controller Time Report"


#Save the report file 
$Reportpath = "C:\Reports\NTP_Server_Source_details_on_all_domaincontroller.HTML"
$HtmlReport | Out-File -Encoding utf8 $Reportpath
Start-Process $Reportpath
