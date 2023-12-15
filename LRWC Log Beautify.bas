'Author: Cyb3rN8TE 2023 
'Version: 1.0.0

Attribute VB_Name = "LRWCLogBeautify"

Sub LRWC_Log_Beautify()
    ' Declare variables
    Dim MR As Range
    Dim cell As Range
    Dim col As Variant
    Dim i As Integer
    Dim undesiredColumns As Variant
    
    ' Define undesired column headers (fields)
    undesiredColumns = Array("Actions", "Log Source Host", _
        "Host (Impacted) KBytes Rcvd", "Host (Impacted) KBytes Sent", _
        "Host (Impacted) KBytes Total", "Host (Impacted) Packets Rcvd", _
        "Host (Impacted) Packets Sent", "Host (Impacted) Packets Total", _
        "User (Origin) Identity", "User (Impacted) Identity", _
        "Rule Block", "First Log Date", "Last Log Date", _
        "Log Sequence Number", "Log Count", "Serial Number", _
        "Priority", "Severity", "Quantity", "Amount", _
        "Size", "Rate", "Duration", "Version")
    
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
    Dim columnOrder As Variant
    columnOrder = Array("Log Source Entity", "Log Date", "Log Source Type", "Log Source", "Session", _
        "User (Origin)", "Host (Origin)", "MAC Address (Origin)", "IP Address (Origin)", _
        "NAT IP Address (Origin)", "Location (Origin)", "Classification", "Common Event", _
        "MPE Rule Name", "Protocol", "Application", "Known Application", "TCP/UDP Port (Origin)", _
        "Direction", "TCP/UDP Port (Impacted)", "User (Impacted)", "Host (Impacted)", _
        "MAC Address (Impacted)", "IP Address (Impacted)", "NAT IP Address (Impacted)", _
        "Location (Impacted)", "User Agent", "Command", "URL", "Action", "Subject", _
        "Reason", "Response Code", "Result", "Status", "Policy", "Vendor Message ID", _
        "Vendor Info", "Object", "Object Name", "Object Type", "Process Name", _
        "Process ID", "Parent Process ID", "Parent Process Name", "Parent Process Path", _
        "Hash", "Threat Name", "Threat ID", "CVE", "Sender", "Recipient", _
        "Sender Identity", "Recipient Identity", "Session Type", "Group", "Country (Origin)", _
        "Country (Impacted)", "Region (Origin)", "Region (Impacted)", "Zone (Origin)", _
        "Zone (Impacted)", "Hostname (Origin)", "Hostname (Impacted)", "Known Host (Origin)", _
        "Known Host (Impacted)", "Interface (Origin)", "Interface (Impacted)", "Network (Origin)", _
        "Network (Impacted)", "NAT TCP/UDP Port (Origin)", "NAT TCP/UDP Port (Impacted)", _
        "Domain (Impacted)", "Domain (Origin)", "Entity (Origin)", "Entity (Impacted)", "Log Message")
    
    Dim columnIndex As Variant
    Dim idx As Long

    For idx = LBound(columnOrder) To UBound(columnOrder)
        Set columnIndex = Nothing
        On Error Resume Next
        Set columnIndex = Rows(1).Find(What:=columnOrder(idx), LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If Not columnIndex Is Nothing Then
            If columnIndex.Column <> idx + 1 Then
                columnIndex.EntireColumn.Cut
                Columns(idx + 1).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        End If
    Next idx
  
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
    
    ' Define columns to defang
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
		
    ' Sort by "Log Source Entity" column in alphabetical order
    Dim tbl As ListObject
    Dim sortColumn As Range
    
    ' Define the table
    Set tbl = ActiveSheet.ListObjects("Table1")
    
    ' Find the "Log Source Entity" column
    Set sortColumn = tbl.ListColumns("Log Source Entity").Range
    
    ' Sort the table by the "Log Source Entity" column in ascending order
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortColumn, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
End Sub
