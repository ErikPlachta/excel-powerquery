Sub RefreshQueriesAndPivots()

    ' Define the arrays for Power Queries and Pivot Tables
    Dim queryNames As Variant
    Dim pivotTableNames As Variant
    
    ' Populate the arrays with the names of the queries and corresponding pivot tables
    queryNames = Array("Query1", "Query2", "Query3")  ' Add your query names here
    pivotTableNames = Array("PivotTable1", "PivotTable2", "PivotTable3")  ' Add your pivot table names here

    Dim ws As Worksheet
    Dim i As Integer
    Dim connection As WorkbookConnection
    Dim pivotTable As PivotTable
    Dim isSuccess As Boolean
    
    ' Error handling initialization
    isSuccess = True
    On Error GoTo ErrorHandler ' Start error handling

    ' Loop through each query in the array and refresh it
    For i = LBound(queryNames) To UBound(queryNames)
        For Each connection In ThisWorkbook.Connections
            If connection.Name = queryNames(i) Then
                connection.Refresh
                Exit For
            End If
        Next connection
    Next i

    ' Refresh the related pivot tables
    For i = LBound(pivotTableNames) To UBound(pivotTableNames)
        ' Assuming pivot tables are in specific worksheets, change as necessary
        For Each ws In ThisWorkbook.Worksheets
            For Each pivotTable In ws.PivotTables
                If pivotTable.Name = pivotTableNames(i) Then
                    pivotTable.RefreshTable
                    Exit For
                End If
            Next pivotTable
        Next ws
    Next i

    ' If everything went well, update the named range "INPUTS_LAST_REFRESH_DT"
    If isSuccess Then
        UpdateLastRefreshDate
    End If

    MsgBox "Queries and Pivot Tables have been refreshed."

    Exit Sub ' Exit the subroutine normally

' Error handler section
ErrorHandler:
    isSuccess = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"

End Sub

' Subroutine to update the "INPUTS_LAST_REFRESH_DT" named range with the current date and time
Sub UpdateLastRefreshDate()
    On Error GoTo RangeError ' Start error handling for named range
    
    ' Update the named range "INPUTS_LAST_REFRESH_DT" with the current date and time
    ThisWorkbook.Names("INPUTS_LAST_REFRESH_DT").RefersToRange.Value = Now
    
    Exit Sub ' Exit the subroutine normally

' Error handler section for named range update
RangeError:
    MsgBox "Failed to update the named range 'INPUTS_LAST_REFRESH_DT': " & Err.Description, vbCritical, "Error"
End Sub