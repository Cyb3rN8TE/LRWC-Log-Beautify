Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


Write-Host ""
Write-Host 'LR ' -NoNewline; Write-Host -ForegroundColor White 'WC Log Exporter ' -NoNewline; Write-Host 'V 1.0.0'
Write-Host 'Compiled by NateDeMaster'
Write-Host ""

function Show-LoadingBar {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]$DurationInSeconds
    )

    $interval = 50
    $totalTicks = [Math]::Floor($DurationInSeconds * 1000 / $interval)
    $completedTicks = 0

    Write-Host "Loading..." -NoNewline

    while ($completedTicks -lt $totalTicks) {
        Write-Host -NoNewline "." -ForegroundColor Cyan
        $completedTicks++
        Start-Sleep -Milliseconds $interval
    }

    Write-Host " Done." -ForegroundColor Green
}


Show-LoadingBar -DurationInSeconds 2

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
Write-Host "Please select your file for processing.."

# Define array of values to check for in column headers
$valuesToDelete = @('User Agent', 'Response Code', 'Quantity', 'Amount', 'Rate', 'Duration', 'Host (Impacted) KBytes Rcvd', 'Host (Impacted) KBytes Sent', 'Host (Impacted) Packets Sent', 'Host (Impacted) Packets Total', 'Severity', 'Vendor Info', 'Serial Number', 'Entity (Origin)', 'Entity (Impacted)', 'Region (Origin)', 'Region (Impacted)', 'Log Count', 'Log Source Host', 'Log Sequence Number', 'First Log Date', 'Last Log Date', 'Rule Block', 'User (Origin) Identity', 'User (Impacted) Identity')

# Create open file dialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$openFileDialog.Filter = "CSV files (*.csv)|*.csv"

# Show the dialog and get the selected file
if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePath = $openFileDialog.FileName

    # Load the CSV file
    $data = Import-Csv -Path $filePath


    # Move columns to reflect provided order

    $orderedColumns = @(
    "Log Source Entity",
    "Log Date",
    "User (Origin)",
    "Session",
    "User (Impacted)",
    "Log Source Type",
    "Log Source",
    "Classification",
    "Common Event",
    "Direction",
    "Host (Origin)",
    "Host (Impacted)",
    "Application",
    "Object",
    "Object Name",
    "Object Type",
    "Hash",
    "Policy",
    "Result",
    "URL",
    "Subject",
    "Version",
    "Command",
    "Reason",
    "Action",
    "Status",
    "Session Type",
    "Process Name",
    "Process ID",
    "Parent Process ID",
    "Parent Process Name",
    "Parent Process Path",
    "Size",
    "Known Application",
    "Host (Impacted) KBytes Total",
    "Host (Impacted) Packets Rcvd",
    "Priority",
    "Vendor Message ID",
    "MPE Rule Name",
    "Threat Name",
    "Threat ID",
    "CVE",
    "MAC Address (Origin)",
    "MAC Address (Impacted)",
    "Interface (Origin)",
    "Interface (Impacted)",
    "IP Address (Origin)",
    "IP Address (Impacted)",
    "NAT IP Address (Origin)",
    "NAT IP Address (Impacted)",
    "Hostname (Origin)",
    "Hostname (Impacted)",
    "Known Host (Origin)",
    "Known Host (Impacted)",
    "Network (Origin)",
    "Network (Impacted)",
    "Domain (Impacted)",
    "Domain (Origin)",
    "Protocol",
    "TCP/UDP Port (Origin)",
    "TCP/UDP Port (Impacted)",
    "NAT TCP/UDP Port (Origin)",
    "NAT TCP/UDP Port (Impacted)",
    "Actions",
    "Sender Identity",
    "Recipient Identity",
    "Sender",
    "Recipient",
    "Group",
    "Zone (Origin)",
    "Zone (Impacted)",
    "Location (Origin)",
    "Location (Impacted)",
    "Country (Origin)",
    "Country (Impacted)",
    "Log Message"
)

# Reorder columns
$data = $data | Select-Object $orderedColumns

    # Remove columns containing specified values
    foreach ($value in $valuesToDelete) {
        $columnIndex = 0
        foreach ($header in $data[0].PSObject.Properties.Name) {
            if ($header -eq $value) {
                $data = $data | Select-Object -Property ($data[0].PSObject.Properties.Name | Where-Object {$_ -ne $value})
                break
            }
            $columnIndex++
        }
    }


# Get the local time zone
$localTimeZone = [TimeZoneInfo]::Local

# Loop through each row of the data
foreach ($row in $data) {
    # Parse the UTC date/time value in the Log Date column
    $utcDateTime = [DateTime]::Parse($row.'Log Date')

    # Convert the UTC date/time value to the local time zone
    $localDateTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcDateTime, $localTimeZone)

    # Replace the value in the Log Date column with the converted local time value
    $row.'Log Date' = $localDateTime.ToString('yyyy-MM-dd HH:mm:ss')
    
}
   

# Define the column names that should be modified
$columnsToModify = @(
    "URL",
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

# Loop through each row and modify the specified columns
foreach ($row in $data) {
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

    # Initialize row count
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

    # Resize the cells to fit the contents
    $range = $sheet.UsedRange
    $range.EntireColumn.AutoFit() | Out-Null

   # Remove progress bar
   Write-Progress -Activity "Converting to Excel workbook..." -Completed

   Write-Host "Processing Complete..."

    # Save and close the workbook
    $workbook.SaveAs($filePath + ".xlsx", 51) # 51 is the XlFileFormat for Excel 2007 or later
    $workbook.Close($true)
    $excel.Quit()

    Write-Host

# Prompt user to select an Excel file
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls"
$dialogResult = $fileDialog.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePath = $fileDialog.FileName
    # Prompt user for formatting confirmation
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to format the Excel workbook as a table?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Load the Excel workbook and format as table
        $xlsx = Import-Excel $filePath
        $xlsx | Format-Table
        $xlsx | Export-Excel -Path $filePath -TableStyle 'Medium13' -AutoSize
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
        # Do nothing, end script
    } else {
        Write-Host "Invalid input. Please select Yes or No."
    }
}


}
