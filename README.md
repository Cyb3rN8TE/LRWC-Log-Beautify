# LRWC Log Beautify

<img src="https://github.com/Cyb3rN8TE/LRWC-Log-Beautify/blob/Dev/Images/Logo.png" alt="LRWC Log Beautify Logo" width="150" height="150">

**Introduction**

LRWC Log Beautify is a PowerShell script designed to assist security analysts in enhancing the readability and usability of log data from the LogRhythm SIEM. This tool simplifies the process of organising, processing, and presenting log data, making it easier to read, filter, and work with, ultimately improving the workflow for security analysts and other users.

**Disclaimer:** Please ensure that you have the necessary permissions to process the data in your CSV files. Handle sensitive information responsibly. LRWC Log Beautify is an independent project and is not affiliated with or endorsed by LogRhythm.

# Windows

**Functionality**
1) Imports raw .CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from MM/DD/YYYY format to DD/MM/YYYY format.
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Formats the spreadsheet as a styled table
6) Save resulting file as a .XLSX Workbook

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

**Functionality**
1) Imports raw .CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from MM/DD/YYYY format to DD/MM/YYYY format.
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Save resulting file as a .CSV

**Prerequisites:**

macOS 10.13 and higher

**Dependencies:**
- PowerShell 7 (https://github.com/PowerShell/PowerShell/releases/tag/v7.3.3)

* * *

**Virus Scan**

- LRWC Log Beautify.ps1 (https://www.virustotal.com/gui/file/e3fe0587c4e6aed45e980a7c76400cd9aa3564cd7f07ef4bd1faf390a828a08e?nocache=1)
- LRWC Log Beautify.bas (https://www.virustotal.com/gui/file/213c8d36cc965d795b95e4a76b774bf2602c3d2c891fef70be5fe2fcf6e5787a?nocache=1)

**Demo**

![alt text](Demo/Demo1.gif)
![alt text](Demo/Demo2.gif)
