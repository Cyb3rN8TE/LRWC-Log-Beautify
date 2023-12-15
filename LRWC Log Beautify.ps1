
# Banner start
Write-Host ""
Write-Host -ForegroundColor White '888      8888888b.  888       888  .d8888b.       888                            888888b.                              888    d8b  .d888          '
Write-Host -ForegroundColor White '888      888   Y88b 888   o   888 d88P  Y88b      888                            888  "88b                             888    Y8P d88P"           '
Write-Host -ForegroundColor White '888      888    888 888  d8b  888 888    888      888                            888  .88P                             888        888             '
Write-Host -ForegroundColor White '888      888   d88P 888 d888b 888 888             888      .d88b.   .d88b.       8888888K.   .d88b.   8888b.  888  888 888888 888 888888 888  888 '
Write-Host -ForegroundColor White '888      8888888P"  888d88888b888 888             888     d88""88b d88P"88b      888  "Y88b d8P  Y8b     "88b 888  888 888    888 888    888  888 '
Write-Host -ForegroundColor White '888      888 T88b   88888P Y88888 888    888      888     888  888 888  888      888    888 88888888 .d888888 888  888 888    888 888    888  888 '
Write-Host -ForegroundColor White '888      888  T88b  8888P   Y8888 Y88b  d88P      888     Y88..88P Y88b 888      888   d88P Y8b.     888  888 Y88b 888 Y88b.  888 888    Y88b 888 '
Write-Host -ForegroundColor White '88888888 888   T88b 888P     Y888  "Y8888P"       88888888 "Y88P"   "Y88888      8888888P"   "Y8888  "Y888888  "Y88888  "Y888 888 888     "Y88888 '
Write-Host -ForegroundColor White '                                                                        888                                                                   888 '
Write-Host -ForegroundColor White 'Author: Cyb3rN8TE 2023                                             Y8b d88P                                                              Y8b d88P '
Write-Host -ForegroundColor White 'Version: 2.0.1                                                      "Y88P"                                                                "Y88P"  '
#Banner end

# Switch statement that detects which operating system is being used (Windows or Mac OS)

switch ($env:OS) {
    # Windows OS Script
    "Windows_NT" {

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        try {
            Add-Type -AssemblyName Microsoft.Office.Interop.Excel -ErrorAction Stop

            Write-Host 
            Write-Host "Microsoft.Office.Interop.Excel assembly found. Proceeding..."
        }
        catch {
            Write-Error "Microsoft.Office.Interop.Excel assembly not found. Please make sure that Microsoft Office is installed on your system"
            Write-Host "Press ENTER key to exit..."
            $null = Read-Host
            Exit
        }

        # Set dimensions for PowerShell window

        $Width = 170
        $Height = 60
        [Console]::SetWindowSize($Width, $Height)

        # Check if ImportExcel Module is installed for the user and if not, install it

        if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    
        Write-Host ""
        Write-Output "ImportExcel module not found. Installing module..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        } else {
        Write-Host ""
        Write-Output "ImportExcel module found."
        }

        # Import Module
        Import-Module ImportExcel

        Write-Host ""
        Write-Host "Windows OS detected"

        do {

            Write-Host ""
            Write-Host "Please select your file for processing.."
            
            # Define array of values to check for in column headers
            $valuesToDelete = @('Actions','Log Source Host','Host (Impacted) KBytes Rcvd','Host (Impacted) KBytes Sent','Host (Impacted) KBytes Total','Host (Impacted) Packets Rcvd','Host (Impacted) Packets Sent','Host (Impacted) Packets Total','User (Origin) Identity','User (Impacted) Identity','Rule Block','First Log Date','Last Log Date','Log Sequence Number','Log Count','Serial Number','Priority','Severity','Quantity','Amount','Size','Rate','Duration','Version')
            
            # Create open file dialog
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
            $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
            
            # Show the dialog and get the selected file
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) 
            {
                Clear-Host

                $filePath = $openFileDialog.FileName
            
                # Load the CSV file
                $data = Import-Csv -Path $filePath
            
                # Move columns to reflect provided order
                $orderedColumns = @(
                    "Log Source Entity",
                    "Log Date",
                    "Log Source Type",
                    "Log Source",
                    "Session",
                    "User (Origin)",
                    "Host (Origin)",
                    "MAC Address (Origin)",
                    "IP Address (Origin)",
                    "NAT IP Address (Origin)",
                    "Location (Origin)",
                    "Classification",
                    "Common Event",
                    "MPE Rule Name",
                    "Protocol",
                    "Application",
                    "Known Application",
                    "TCP/UDP Port (Origin)",
                    "Direction",
                    "TCP/UDP Port (Impacted)",
                    "User (Impacted)",
                    "Host (Impacted)",
                    "MAC Address (Impacted)",
                    "IP Address (Impacted)",
                    "NAT IP Address (Impacted)",
                    "Location (Impacted)",
                    "User Agent",
                    "Command",
                    "URL",
                    "Action",
                    "Subject",
                    "Reason",
                    "Response Code",
                    "Result",
                    "Status",
                    "Policy",
                    "Vendor Message ID",
                    "Vendor Info",
                    "Object",
                    "Object Name",
                    "Object Type",
                    "Process Name",
                    "Process ID",
                    "Parent Process ID",
                    "Parent Process Name",
                    "Parent Process Path",
                    "Hash",
                    "Threat Name",
                    "Threat ID",
                    "CVE",
                    "Sender",
                    "Recipient",
                    "Sender Identity",
                    "Recipient Identity",
                    "Session Type",
                    "Group",
                    "Country (Origin)",
                    "Country (Impacted)",
                    "Region (Origin)",
                    "Region (Impacted)",
                    "Zone (Origin)",
                    "Zone (Impacted)",
                    "Hostname (Origin)",
                    "Hostname (Impacted)",
                    "Known Host (Origin)",
                    "Known Host (Impacted)",
                    "Interface (Origin)",
                    "Interface (Impacted)",
                    "Network (Origin)",
                    "Network (Impacted)",
                    "NAT TCP/UDP Port (Origin)",
                    "NAT TCP/UDP Port (Impacted)",
                    "Domain (Impacted)",
                    "Domain (Origin)",
                    "Entity (Origin)",
                    "Entity (Impacted)",
                    "Log Message"
                )
            
                # Reorder columns
                $data = $data | Select-Object $orderedColumns
            
                # Define the column names that should be modified
                $columnsToModify = @(
                    "URL",
                    "Subject",
                    "Host (Origin)",
                    "Host (Impacted)",
                    "IP Address (Origin)",
                    "IP Address (Impacted)",
                    "NAT IP Address (Origin)",
                    "NAT IP Address (Impacted)",
                    "Hostname (Origin)",
                    "Hostname (Impacted)",
                    "Known Host (Origin)",
                    "Known Host (Impacted)",
                    "Domain (Impacted)",
                    "Domain (Origin)",
                    "Protocol",
                    "Log Message"
                )

                # Convert Log Date format from MM/DD/YYYY to DD/MM/YYYY
                foreach ($row in $data) {
                    if ($row.PSObject.Properties.Name -contains 'Log Date' -and ![string]::IsNullOrEmpty($row."Log Date")) {
                        $logDate = Get-Date $row."Log Date"
                        $formattedLogDate = $logDate.ToString("dd/MM/yyyy HH:mm:ss")
                        $row."Log Date" = $formattedLogDate
                    }
                }

                # Loop through each row of the data
                foreach ($row in $data) {
                    # Remove columns containing specified values
                    foreach ($value in $valuesToDelete) {
                        if ($row.PSObject.Properties.Name -contains $value) {
                            $row.PSObject.Properties.Remove($value)
                        }
                    }
            
                    # Modify specified columns
                    foreach ($column in $columnsToModify) {
                        $value = $row.$column
                            
                        # Check if the value is not empty
                        if (![string]::IsNullOrEmpty($value)) {
                            # Replace '.' and ':' with '[.]' and '[:]', respectively
                            $newValue = $value.Replace(".", "[.]").Replace(":", "[:]")
                            $row.$column = $newValue
                        }
                    }
                }
        
                # Initialise row count
                $rowCount = $data.Count
        
                # Convert the contents to an Excel workbook
                $excel = New-Object -ComObject Excel.Application
                $workbook = $excel.Workbooks.Add()
                $sheet = $workbook.Sheets.Item(1)
        
                # Write the CSV data to the worksheet
                $rowIndex = 1
                foreach ($row in $data) {
                    $colIndex = 1
                    foreach ($prop in $row.PSObject.Properties) {
                        $value = $prop.Value
                        # Write the header row to the worksheet
                        if ($rowIndex -eq 1) {
                            $sheet.Cells.Item($rowIndex, $colIndex) = $prop.Name
                        }
                        # Write the data rows to the worksheet
                        $sheet.Cells.Item($rowIndex + 1, $colIndex) = $value
                        $colIndex++
                    }
                    $rowIndex++
        
                    # Display progress bar
                    $percentComplete = [int]([Math]::Min($rowIndex / $rowCount * 100, 100))
        
                    Write-Progress -Activity "Processing File and Converting to Excel workbook..." -PercentComplete $percentComplete
                }
        
                    # Resise the cells to fit the contents
                    $range = $sheet.UsedRange
                    $range.EntireColumn.AutoFit() | Out-Null
        
                    # Apply TableStyleMedium13
                    $table = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $range, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
                    $table.TableStyle = "TableStyleMedium13"

                    # Remove progress bar
                    Write-Progress -Activity "Converting to Excel workbook..." -Completed
                    Write-Host""
                    Write-Host "Processing Complete..."
        
                    # Show save dialog
                    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to select where to save the Excel file?" + [Environment]::NewLine + "" + [Environment]::NewLine + "Note: Selecting no will save the file in the same location as the original CSV", "Save As", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                    
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Create save file dialog
                        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                        $saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                        $saveFileDialog.FileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
                        $saveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($filePath)
        
                        # Show the dialog and get the selected filename and location
                        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $excel.DisplayAlerts = $false
                            $workbook.SaveAs($saveFileDialog.FileName, 51)
                            $excel.DisplayAlerts = $true
                            $workbook.Close()
                            $excel.Quit()
        
                            # Release the Excel object from memory to prevent memory leaks
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
                            # Kill excel processes that are not visible
                            Get-Process Excel | Where-Object {$_.MainWindowTitle -eq ''} | Stop-Process
                        }
                    }
        
                    else {
        
                        $excel.DisplayAlerts = $false
                        $workbook.SaveAs($filePath + ".xlsx", 51)
                        $excel.DisplayAlerts = $true
                        $workbook.Close()
                        $excel.Quit()
                        
                        # Release the Excel object from memory to prevent memory leaks
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
                        # Kill excel processes that are not visible
                        Get-Process Excel | Where-Object {$_.MainWindowTitle -eq ''} | Stop-Process
                    }
        
            }

            Clear-Host
        
            # Prompt the user to process another file
            $msgBoxTitle = "Process more files?"
            $msgBoxMessage = "Do you want to process more CSV files?"
            $msgBoxButtons = [System.Windows.Forms.MessageBoxButtons]::YesNo
            $msgBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Question
            $processMoreFiles = [System.Windows.Forms.MessageBox]::Show($msgBoxMessage, $msgBoxTitle, $msgBoxButtons, $msgBoxIcon)

        # Keep prompting user to process more CSV files until the user selects no
} while ($processMoreFiles -eq [System.Windows.Forms.DialogResult]::Yes)

    }
    # Mac OS Script
    default {

        do {
            
            Write-Host ""
            # Prompt the user for the input file path
            do {
                Write-Host "Please specify the input CSV file path..."
                $inputFilePath = Read-Host
            } while (-not (Test-Path $inputFilePath))

            # Prompt the user for the output file path and name
            do {
                Write-Host "Please specify the output CSV file path and filename..."
                $outputFilePath = Read-Host
            } while (-not (Test-Path (Split-Path $outputFilePath)) -or (Split-Path $outputFilePath).Equals($null))

            # If the output file name is not specified, use the input file name with "_modified" appended
            if ($outputFilePath -eq "") {
                $outputFilePath = [IO.Path]::ChangeExtension($inputFilePath, "csv")
                $outputFilePath = [IO.Path]::Combine((Split-Path $inputFilePath), [IO.Path]::GetFileNameWithoutExtension($outputFilePath) + "_modified.csv")
            }
            
            # Define array of values to check for in column headers
            $valuesToDelete = @('Actions','Log Source Host','Host (Impacted) KBytes Rcvd','Host (Impacted) KBytes Sent','Host (Impacted) KBytes Total','Host (Impacted) Packets Rcvd','Host (Impacted) Packets Sent','Host (Impacted) Packets Total','User (Origin) Identity','User (Impacted) Identity','Rule Block','First Log Date','Last Log Date','Log Sequence Number','Log Count','Serial Number','Priority','Severity','Quantity','Amount','Size','Rate','Duration','Version')
            
            # Load the CSV file
            $data = Import-Csv -Path $inputFilePath
            
            # Move columns to reflect provided order
            $orderedColumns = @(
                "Log Source Entity",
                "Log Date",
                "Log Source Type",
                "Log Source",
                "Session",
                "User (Origin)",
                "Host (Origin)",
                "MAC Address (Origin)",
                "IP Address (Origin)",
                "NAT IP Address (Origin)",
                "Location (Origin)",
                "Classification",
                "Common Event",
                "MPE Rule Name",
                "Protocol",
                "Application",
                "Known Application",
                "TCP/UDP Port (Origin)",
                "Direction",
                "TCP/UDP Port (Impacted)",
                "User (Impacted)",
                "Host (Impacted)",
                "MAC Address (Impacted)",
                "IP Address (Impacted)",
                "NAT IP Address (Impacted)",
                "Location (Impacted)",
                "User Agent",
                "Command",
                "URL",
                "Action",
                "Subject",
                "Reason",
                "Response Code",
                "Result",
                "Status",
                "Policy",
                "Vendor Message ID",
                "Vendor Info",
                "Object",
                "Object Name",
                "Object Type",
                "Process Name",
                "Process ID",
                "Parent Process ID",
                "Parent Process Name",
                "Parent Process Path",
                "Hash",
                "Threat Name",
                "Threat ID",
                "CVE",
                "Sender",
                "Recipient",
                "Sender Identity",
                "Recipient Identity",
                "Session Type",
                "Group",
                "Country (Origin)",
                "Country (Impacted)",
                "Region (Origin)",
                "Region (Impacted)",
                "Zone (Origin)",
                "Zone (Impacted)",
                "Hostname (Origin)",
                "Hostname (Impacted)",
                "Known Host (Origin)",
                "Known Host (Impacted)",
                "Interface (Origin)",
                "Interface (Impacted)",
                "Network (Origin)",
                "Network (Impacted)",
                "NAT TCP/UDP Port (Origin)",
                "NAT TCP/UDP Port (Impacted)",
                "Domain (Impacted)",
                "Domain (Origin)",
                "Entity (Origin)",
                "Entity (Impacted)",
                "Log Message"
            )
            
            # Reorder columns
            $data = $data | Select-Object $orderedColumns
            
            # Define the column names that should be modified
            $columnsToModify = @(
                "URL",
                "Subject",
                "Host (Origin)",
                "Host (Impacted)",
                "IP Address (Origin)",
                "IP Address (Impacted)",
                "NAT IP Address (Origin)",
                "NAT IP Address (Impacted)",
                "Hostname (Origin)",
                "Hostname (Impacted)",
                "Known Host (Origin)",
                "Known Host (Impacted)",
                "Domain (Impacted)",
                "Domain (Origin)",
                "Protocol",
                "Log Message"
            )

            # Loop through each row of the data
            foreach ($row in $data) {
                # Remove columns containing specified values
                foreach ($value in $valuesToDelete) {
                    if ($row.PSObject.Properties.Name -contains $value) {
                        $row.PSObject.Properties.Remove($value)
                    }
                }

                # Modify specified columns
                foreach ($column in $columnsToModify) {
                    $value = $row.$column

                    # Check if the value is not empty
                    if (![string]::IsNullOrEmpty($value)) {
                        # Defang values
                        $newValue = $value.Replace(".", "[.]").Replace(":", "[:]")
                        $row.$column = $newValue
                    }
                }
            }

            # Convert Log Date format from MM/DD/YYYY to DD/MM/YYYY
            $logDate = [DateTime]::ParseExact($row.'Log Date', 'MM/dd/yyyy HH:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
            $row.'Log Date' = $logDate.ToString('dd/MM/yyyy HH:mm:ss tt')
            
            # Export the modified data to a CSV file
            $data | Export-Csv -Path $outputFilePath -NoTypeInformation
        
            # Prompt user to process more CSV files
            $answer = Read-Host "Do you want to process more CSV files? (Y/N)"
            if ($answer -eq "Y") {
                continue
            } elseif ($answer -eq "N") {
                break
            } else {
                Write-Host "Invalid input. Please enter Y or N."
            }
        
        } while ($true)
    }
}

# Prompt the user to exit
Write-Host ""
Write-Host "Press ENTER key to exit..."
$null = Read-Host

