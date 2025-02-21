#Requires -Module ActiveDirectory

function New-ADUsersFromCSV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, HelpMessage = "Path to export log file for created users (default: CreatedUsers_<date>.csv in script directory)")]
        [string]$ExportPath = "$PSScriptRoot\CreatedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
        
        [Parameter(Mandatory = $false, HelpMessage = "Path to export log file for already existing users (default: AlreadyExistedUsers_<date>.csv in script directory)")]
        [string]$AlreadyExistPath = "$PSScriptRoot\AlreadyExistedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

        [Parameter(Mandatory = $false, HelpMessage = "Default OU if not specified in CSV")]
        [string]$DefaultOU = "OU=Users,DC=example,DC=com",

        [Parameter(Mandatory = $false, HelpMessage = "Default password if not specified in CSV")]
        [string]$DefaultPassword = "P@ssw0rd123!"
    )

    # Define function to keep PowerShell window visible
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    '
    $SW_RESTORE = 9
    $hwnd = [Console.Window]::GetForegroundWindow()

    # Display script heading
    Clear-Host
    Write-Host "=============================================================" -ForegroundColor Cyan
    Write-Host "       New-ADUsersFromCSV Script       " -ForegroundColor Yellow -BackgroundColor DarkBlue
    Write-Host " Purpose: Create Active Directory Users from a CSV File" -ForegroundColor Cyan
    Write-Host " Features: GUI file picker, sample data preview, exports created and existing users" -ForegroundColor Cyan
    Write-Host "=============================================================" -ForegroundColor Cyan
    Write-Host ""

    # Load Windows Forms assembly
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Write-Host "Windows Forms loaded successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to load System.Windows.Forms: $_"
        return
    }

    # Restore PowerShell window
    [Console.Window]::ShowWindowAsync($hwnd, $SW_RESTORE) | Out-Null

    # Configure OpenFileDialog
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Select a CSV File Containing New User Data"
    $OpenFileDialog.Multiselect = $false

    # Show OpenFileDialog
    $dialogResult = $OpenFileDialog.ShowDialog()
    if ($dialogResult -ne 'OK') {
        Write-Warning "No CSV file selected. Exiting script."
        Start-Sleep -Seconds 2
        return
    }
    $CSVPath = $OpenFileDialog.FileName

    # Validate CSV file
    if (-not (Test-Path -Path $CSVPath)) {
        Write-Error "CSV file not found at $CSVPath. Exiting."
        return
    }

    # Import CSV
    try {
        $Users = Import-Csv -Path $CSVPath
        if (-not $Users) {
            Write-Warning "CSV file is empty. Nothing to process."
            return
        }
    } catch {
        Write-Error "Failed to import CSV: $_"
        return
    }

    # Display sample data
    Write-Host "Sample Data from CSV (up to 3 rows):" -ForegroundColor Yellow
    $sampleData = $Users | Select-Object -First 3 | Format-Table -AutoSize | Out-String
    Write-Host $sampleData -ForegroundColor White
    Write-Host "Total rows in CSV: $($Users.Count)" -ForegroundColor Yellow
    Write-Host ""

    # Restore PowerShell window before confirmation
    [Console.Window]::ShowWindowAsync($hwnd, $SW_RESTORE) | Out-Null

    # Form-based confirmation
    Write-Host "Preparing confirmation dialog..." -ForegroundColor Cyan
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Confirm User Creation"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.TopMost = $true
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 60)
    $label.Text = "Review the sample data above. Proceed with creating $($Users.Count) users from '$CSVPath'?"
    $form.Controls.Add($label)

    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(100, 100)
    $yesButton.Size = New-Object System.Drawing.Size(75, 30)
    $yesButton.Text = "Yes"
    $yesButton.Add_Click({ $form.Tag = "Yes"; $form.Close() })
    $form.Controls.Add($yesButton)

    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(200, 100)
    $noButton.Size = New-Object System.Drawing.Size(75, 30)
    $noButton.Text = "No"
    $noButton.Add_Click({ $form.Tag = "No"; $form.Close() })
    $form.Controls.Add($noButton)

    Write-Host "Displaying confirmation dialog..." -ForegroundColor Cyan
    $form.ShowDialog() | Out-Null
    $confirmation = $form.Tag

    if ($confirmation -ne "Yes") {
        Write-Warning "User creation cancelled."
        Start-Sleep -Seconds 2
        return
    }
    Write-Host "Proceeding with user creation..." -ForegroundColor Green

    # Restore PowerShell window after confirmation
    [Console.Window]::ShowWindowAsync($hwnd, $SW_RESTORE) | Out-Null

    # Arrays for tracking
    $CreatedUsers = @()
    $AlreadyExist = @()

    # Process each user
    foreach ($User in $Users) {
        $SamAccountName = $User.SamAccountName
        if ([string]::IsNullOrEmpty($SamAccountName)) {
            Write-Warning "Skipping entry with missing SamAccountName in CSV."
            continue
        }

        try {
            $FirstName = if ([string]::IsNullOrEmpty($User.FirstName)) { "" } else { $User.FirstName }
            $LastName = if ([string]::IsNullOrEmpty($User.LastName)) { "" } else { $User.LastName }
            $OU = if ([string]::IsNullOrEmpty($User.OU)) { $DefaultOU } else { $User.OU }
            $Password = if ([string]::IsNullOrEmpty($User.Password)) { $DefaultPassword } else { $User.Password }
            $UPN = if ([string]::IsNullOrEmpty($User.UserPrincipalName)) { "$SamAccountName@example.com" } else { $User.UserPrincipalName }

            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$OU'" -ErrorAction SilentlyContinue)) {
                Write-Warning "OU '$OU' does not exist for $SamAccountName. Using default OU: $DefaultOU"
                $OU = $DefaultOU
            }

            if (Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue) {
                Write-Warning "User $SamAccountName already exists. Skipping creation."
                $AlreadyExist += [PSCustomObject]@{
                    SamAccountName = $SamAccountName
                    FullName       = "$FirstName $LastName".Trim()
                    UPN            = $UPN
                    OU             = $OU
                    Status         = "Already Exists"
                    CheckDate      = Get-Date
                }
                continue
            }

            $NewUserParams = @{
                Name                  = "$FirstName $LastName".Trim()
                GivenName             = $FirstName
                Surname               = $LastName
                SamAccountName        = $SamAccountName
                UserPrincipalName     = $UPN
                Path                  = $OU
                AccountPassword       = (ConvertTo-SecureString $Password -AsPlainText -Force)
                Enabled               = $true
                ChangePasswordAtLogon = $true
                ErrorAction           = 'Stop'
            }

            New-ADUser @NewUserParams

            $CreatedUsers += [PSCustomObject]@{
                SamAccountName = $SamAccountName
                FullName       = "$FirstName $LastName".Trim()
                UPN            = $UPN
                OU             = $OU
                Password       = $Password
                CreationDate   = Get-Date
                Status         = "Success"
            }

            Write-Host "User $SamAccountName created successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to create user $SamAccountName : $_"
            $CreatedUsers += [PSCustomObject]@{
                SamAccountName = $SamAccountName
                FullName       = "$FirstName $LastName".Trim()
                UPN            = $UPN
                OU             = $OU
                Password       = $Password
                CreationDate   = Get-Date
                Status         = "Failed - $_"
            }
        }
    }

    # Export created users
    if ($CreatedUsers) {
        try {
            $CreatedUsers | Export-Csv -Path $ExportPath -NoTypeInformation -Force
            Write-Host "Created users exported to $ExportPath" -ForegroundColor Cyan
        } catch {
            Write-Error "Failed to export created users to $ExportPath : $_"
        }
    } else {
        Write-Warning "No new users were created to export."
    }

    # Export already existing users
    if ($AlreadyExist) {
        try {
            $AlreadyExist | Export-Csv -Path $AlreadyExistPath -NoTypeInformation -Force
            Write-Host "Already existing users exported to $AlreadyExistPath" -ForegroundColor Cyan
        } catch {
            Write-Error "Failed to export already existing users to $AlreadyExistPath : $_"
        }
    } else {
        Write-Warning "No users were skipped as already existing."
    }

    # Execution Summary with corrected counts
    Write-Host ""
    Write-Host "Execution Summary:" -ForegroundColor Yellow
    $createdCount = ($CreatedUsers | Where-Object { $_.Status -eq 'Success' }).Count
    $failedCount = ($CreatedUsers | Where-Object { $_.Status -ne 'Success' }).Count
    $existingCount = $AlreadyExist.Count
    Write-Host " - Users Created: $createdCount" -ForegroundColor Green
    Write-Host " - Users Failed: $failedCount" -ForegroundColor Red
    Write-Host " - Users Already Existed: $existingCount" -ForegroundColor Magenta
    Start-Sleep -Seconds 3
}

# Example Usage:
New-ADUsersFromCSV