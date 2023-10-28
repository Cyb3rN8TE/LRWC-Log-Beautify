# LRWC Log Beautify

<img src="https://github.com/Cyb3rN8TE/LRWC-Log-Beautify/blob/Dev/Images/Logo.png" alt="LRWC Log Beautify Logo" width="150" height="150">

**Introduction**

LRWC Log Beautify is a PowerShell script designed to assist security analysts in enhancing the readability and usability of log data from the LogRhythm SIEM. This tool simplifies the process of organising, processing, and presenting log data, making it easier to read, filter, and work with, ultimately improving the workflow for security analysts and other users.

**Disclaimer:** Please ensure that you have the necessary permissions to process the data in your CSV files. Handle sensitive information responsibly. LRWC Log Beautify is an independent project and is not affiliated with or endorsed by LogRhythm.

# Windows

**What does it do?**
1) Imports raw .CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from UTC to Local Time
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Formats the spreadsheet as a styled table
6) Save resulting file as a .XLSX Workbook

**How to run the tool:**
1) Right click LRWC Log Beautify.ps1 and select 'run with PowerShell'
2) If NuGet and ImportExcel are NOT detected on the system, accept any prompts if you want to download and install them automatically. 
3) Alternatively, you can download and install them manually by entering the commands below:

**Prerequisites:**

Microsoft Office 2016 installed on system

Microsoft Windows OS:
- Windows 7 with Service Pack 1 (SP1)
- Windows Vista with Service Pack 2 (SP2)
- Windows 8
- Windows 10
- Windows 11

**Dependencies:**
- PowerShell 7 (https://github.com/PowerShell/PowerShell/releases/tag/v7.3.3)
- NuGet Package Manager (https://www.nuget.org/downloads)
- ImportExcel (https://www.powershellgallery.com/packages/ImportExcel/)

**Installing NuGet**

Open PowerShell as an administrator.
```
Install-PackageProvider -Name NuGet -Force
```
Once the installation is complete, you can verify that NuGet is installed by running the following command:
```
Get-PackageProvider -Name NuGet
```
This should display the version of NuGet that is installed on your system.

**Installing ImportExcel**

Open PowerShell as an administrator.

Run the following command to install the module from the PowerShell Gallery:
```
Install-Module -Name ImportExcel
```

# macOS

**What does it do?**
1) Imports raw .CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from UTC to Local Time
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Save resulting file as a .CSV

**How to run the tool:**
1) Start the macOS native terminal application
2) Enter the command powershell and hit enter

```
pwsh
```

3) Run LRWC Log Beautify.ps1

```
./LRWC Log Beautify.ps1
```

**Prerequisites:**

macOS 10.13 and higher

**Dependencies:**
- PowerShell 7 (https://github.com/PowerShell/PowerShell/releases/tag/v7.3.3)

* * *

**Virus Scan**

- LRWC Log Beautify.ps1 (https://www.virustotal.com/gui/file/8e43def286f2e006894df6eb2e4ceefd4d0ac82fea09ef80a69057acdb97e351/detection)

**Demo**

![alt text](Demo/Demo1.gif)
![alt text](Demo/Demo2.gif)
