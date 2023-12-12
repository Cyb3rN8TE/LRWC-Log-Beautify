Attribute VB_Name = "LRWCLogBeautify"

Sub LRWC_Log_Beautify()
    Dim MR As Range
    Dim cell As Range
    Dim col As Variant
    Dim i As Integer
    Dim undesiredColumns As Variant
    
    ' Define undesired column headers (fields)
    undesiredColumns = Array("Quantity", "Amount", "Rate", "Duration", "Host (Impacted) KBytes Rcvd", "Host (Impacted) KBytes Sent", _
                             "Host (Impacted) Packets Sent", "Host (Impacted) Packets Total", "Severity", "Vendor Info", "Serial Number", _
                             "Region (Origin)", "Region (Impacted)", "Log Count", "Log Source Host", "Log Sequence Number", "First Log Date", _
                             "Last Log Date", "Rule Block", "User (Origin) Identity", "User (Impacted) Identity")
    
    ' Set the range of columns to search starting from row 1
    Set MR = Range("A1:ED1")
    
    ' Delete undesired columns (fields)
    For Each col In undesiredColumns
        For Each cell In MR
            If cell.Value = col Then
                cell.EntireColumn.Delete
                Exit For
            End If
        Next cell
    Next col
    
    ' Rearrange columns
    Dim desiredColumns As Variant
    desiredColumns = Array("Log Source Entity", "Log Date", "Log Source Type", "Log Source", "Session", "User (Origin)", "Host (Origin)", "MAC Address (Origin)", _
                           "IP Address (Origin)", "Location (Origin)", "Classification", "Common Event", "MPE Rule Name", "Protocol", "TCP/UDP Port (Origin)", _
                           "Direction", "TCP/UDP Port (Impacted)", "User (Impacted)", "Host (Impacted)", "MAC Address (Impacted)", "IP Address (Impacted)", _
                           "Location (Impacted)", "Application", "Object", "Object Name", "Object Type", "Hash", "Policy", "Result", "URL", "User Agent", _
                           "Subject", "Version", "Command", "Response Code", "Reason", "Action", "Status", "Session Type", "Process Name", "Process ID", _
                           "Parent Process ID", "Parent Process Name", "Parent Process Path", "Size", "Known Application", "Priority", "Vendor Message ID", _
                           "Threat Name", "Threat ID", "CVE", "Actions", "Sender Identity", "Recipient Identity", "Sender", "Recipient", "Group", _
                           "NAT TCP/UDP Port (Origin)", "NAT TCP/UDP Port (Impacted)", "Interface (Origin)", "Interface (Impacted)", "NAT IP Address (Origin)", _
                           "NAT IP Address (Impacted)", "Network (Origin)", "Network (Impacted)", "Domain (Impacted)", "Domain (Origin)", "Zone (Origin)", _
                           "Zone (Impacted)", "Country (Origin)", "Country (Impacted)", "Hostname (Origin)", "Hostname (Impacted)", "Known Host (Origin)", _
                           "Known Host (Impacted)", "Entity (Origin)", "Entity (Impacted)", "Log Message")
    
    i = 1
    For Each col In desiredColumns
        For Each cell In MR
            If cell.Value = col Then
                cell.EntireColumn.Cut Destination:=Cells(1, i)
                i = i + 1
                Exit For
            End If
        Next cell
    Next col
    
    ' Autofit columns
    Cells.Columns.AutoFit
    
    ' Change to DD/MM/YYYY format in "Log Date" column
    Dim logDateRange As Range
    Dim cellInLogDate As Range 
    
    ' Define the range for the "Log Date" column
    Set logDateRange = Range("B2:B" & Cells(Rows.Count, "B").End(xlUp).Row)
    
    For Each cellInLogDate In logDateRange
        ' Check if the cell is not empty
        If cellInLogDate.Value <> "" Then
            ' Split the date value by '/' delimiter
            Dim parts() As String
            parts = Split(cellInLogDate.Value, "/")
            
            ' Check if the date has three parts (day/month/year)
            If UBound(parts) = 2 Then
                ' Swap the day and month parts
                Dim temp As String
                temp = parts(1) ' variable to hold the month
                parts(1) = parts(0) ' Assign month to the day part
                parts(0) = temp ' Assign day to the month part
                
                ' Join the parts back together and update the cell value
                cellInLogDate.Value = Join(parts, "/")
            End If
        End If
    Next cellInLogDate
    
    ' Replace characters in specified columns
    Dim replaceColumns As Variant
    replaceColumns = Array("URL", "Subject", "Host (Origin)", "Host (Impacted)", "IP Address (Origin)", "IP Address (Impacted)", _
                           "NAT IP Address (Origin)", "NAT IP Address (Impacted)", "Hostname (Origin)", "Hostname (Impacted)", _
                           "Known Host (Origin)", "Known Host (Impacted)", "Domain (Impacted)", "Domain (Origin)", "Protocol", "Log Message")
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim colName As Variant
    Dim lastRow As Long
    Dim colReplace As Range
    
    ' Defang artefacts

    For Each colName In replaceColumns
        Set colReplace = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not colReplace Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, colReplace.Column).End(xlUp).Row
            Set colReplace = ws.Range(ws.Cells(2, colReplace.Column), ws.Cells(lastRow, colReplace.Column))
            
            For Each cellReplace In colReplace
                If Not IsEmpty(cellReplace.Value) Then
                    cellReplace.Value = Replace(cellReplace.Value, ".", "[.]")
                    cellReplace.Value = Replace(cellReplace.Value, ":", "[:]")
                End If
            Next cellReplace
        End If
    Next colName
    
    ' Convert data to a table
    ActiveSheet.ListObjects.Add(xlSrcRange, _
        Range("A1").CurrentRegion, XlListObjectHasHeaders:=xlYes, _
        TableStyleName:="TableStyleMedium7").Name = "Table1"
End Sub
