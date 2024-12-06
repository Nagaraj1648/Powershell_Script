# List of servers to check
$Servers = @("ad-chn-01")  # Replace with your server names

# Function to check pending restart status
function Check-PendingRestart {
    param([string]$ComputerName)
    try {
        # Query the registry for pending restart status
        $Pending = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            $Pending = @{
                ComponentBasedServicing = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
                WindowsUpdate           = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
                PendingFileRename       = Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
            }
            $Pending
        }
        return [PSCustomObject]@{
            Server                  = $ComputerName
            ComponentBasedServicing = $Pending.ComponentBasedServicing
            WindowsUpdate           = $Pending.WindowsUpdate
            PendingFileRename       = $Pending.PendingFileRename
            NeedsRestart            = ($Pending.ComponentBasedServicing -or $Pending.WindowsUpdate -or $Pending.PendingFileRename)
        }
    } catch {
        Write-Warning "Unable to check $ComputerName : $_"
        return $null
    }
}

# Check all servers
$Results = $Servers | ForEach-Object { Check-PendingRestart -ComputerName $_ }

# Export results to CSV
$Results | Export-Csv -Path "C:\PendingRestartReport.csv" -NoTypeInformation -Encoding UTF8

# Display results
$Results | Format-Table -AutoSize
