# LRWC Log Beautify

**Why do we need this?**

It can be difficult to interpret SIEM logs from a CSV file due to the lack of formatting.  

This tool aims to make SIEM logs appear presentable to anyone who accesses the log export, whether that be a colleague, manager or a client.

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
- Microsoft.Office.Interop.Excel.dll in the root directory of the script (provided)
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
- Microsoft.Office.Interop.Excel.dll (https://www.virustotal.com/gui/file/d5fbf3f71c40ca63b27601b5275c1cf5dc0cfd187c972579e2100f9215a375fe)

**Demo**

![alt text](Demo/Demo1.gif)
![alt text](Demo/Demo2.gif)
