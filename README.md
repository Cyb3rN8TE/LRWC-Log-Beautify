# LRWC Log Beautify

**Why do we need this?**

When you export logs from the LogRhythm Web Console, it puts the logs in a CSV file with the fields as column headings. It can be difficult to interpret the logs from the CSV given not only the unformatted nature of the file type but also the 'Log Date' time format being in UTC. This tool aims to format the logs so that they appear presentable to anyone who accesses the log export, should that be a colleague, manager or a client. 

**What does it do?**
1) Imports raw CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from UTC to Local Time
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Formats the spreadsheet as a table
6) Save resulting file as a .XLSX Workbook

**Prerequisites:**

Microsoft Office 2016 installed on system

**Dependencies:**
- Microsoft.Office.Interop.Excel.dll in the root directory of the script (provided)
- NuGet Package Manager (https://www.nuget.org/downloads)
- ImportExcel (https://www.powershellgallery.com/packages/ImportExcel/)

**Virus Scan**

https://www.virustotal.com/gui/file/a8038e6b8580a66952db703427b8cc13902b9545f1794dc95537c134b05dd0eb/detection

**How to run the tool:**
1) Right click LRWC Log Beautify.ps1 and select 'run with PowerShell'
2) If NuGet and ImportExcel are NOT detected on the system, accept any prompts if you want to download and install them automatically. 
3) Alternatively, you can download and install them manually by entering the commands below:

**NuGet**

Open PowerShell as an administrator.
```
Install-PackageProvider -Name NuGet -Force
```
Once the installation is complete, you can verify that NuGet is installed by running the following command:
```
Get-PackageProvider -Name NuGet
```
This should display the version of NuGet that is installed on your system.

**ImportExcel**

Open PowerShell as an administrator.

Run the following command to install the module from the PowerShell Gallery:
```
Install-Module -Name ImportExcel
```
