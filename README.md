# LR WC Log Exporter

Why do we need this?

When you export logs from the LogRhythm Web Console, it puts the logs in a CSV file with the fields as column headings. It can be difficult to interpret the logs from the CSV given not only the unformatted nature of the file type but also the 'Log Date' time format being in UTC. This tool aims to format the logs so that they appear presentable to anyone who accesses the log export, should that be a colleague, manager or a client. 

What does it do?
1) Imports raw CSV that the user selects
2) Re-arranges field columns for readability
3) Converts 'Log Date' from UTC to Local Time
4) Defangs any IP Addresses, URLs and Domains within any cells.
5) Formats the spreadsheet as a table (selectable option) 
6) Save resulting file as a .XLSX Workbook
