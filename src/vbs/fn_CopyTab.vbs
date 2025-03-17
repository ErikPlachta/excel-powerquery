Dim xlApp, srcWorkbook, destWorkbook, srcSheet, destSheet, sheetName, i, namedRange
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True ' Ensure Excel is visible

' Prompt user for the source workbook (active workbook)
Set srcWorkbook = xlApp.ActiveWorkbook
If srcWorkbook Is Nothing Then
    MsgBox "No active workbook found. Open a workbook and try again.", vbExclamation, "Error"
    WScript.Quit
End If

' Ask the user to select a sheet
sheetName = InputBox("Enter the name of the sheet to copy:", "Select Sheet")
If sheetName = "" Then WScript.Quit

On Error Resume Next
Set srcSheet = srcWorkbook.Sheets(sheetName)
On Error GoTo 0

If srcSheet Is Nothing Then
    MsgBox "Sheet not found. Please enter a valid sheet name.", vbExclamation, "Error"
    WScript.Quit
End If

' List open workbooks and prompt user for destination workbook
Dim wbNames
wbNames = ""
For i = 1 To xlApp.Workbooks.Count
    If xlApp.Workbooks(i).Name <> srcWorkbook.Name Then
        wbNames = wbNames & vbCrLf & i & ": " & xlApp.Workbooks(i).Name
    End If
Next

If wbNames = "" Then
    MsgBox "No other open workbooks found. Open a workbook and try again.", vbExclamation, "Error"
    WScript.Quit
End If

Dim destIndex
destIndex = InputBox("Select the destination workbook by entering its number:" & vbCrLf & wbNames, "Select Destination Workbook")
If Not IsNumeric(destIndex) Or destIndex < 1 Or destIndex > xlApp.Workbooks.Count Then WScript.Quit

Set destWorkbook = xlApp.Workbooks(CInt(destIndex))

' Copy sheet to new workbook
srcSheet.Copy , destWorkbook.Sheets(destWorkbook.Sheets.Count)
Set destSheet = destWorkbook.Sheets(destWorkbook.Sheets.Count)

' Remove duplicate named ranges to avoid conflicts
For Each namedRange In destWorkbook.Names
    If InStr(1, namedRange.RefersTo, srcSheet.Name, vbTextCompare) > 0 Then
        On Error Resume Next
        namedRange.Delete
        On Error GoTo 0
    End If
Next

' Notify user
MsgBox "Sheet '" & sheetName & "' copied successfully to '" & destWorkbook.Name & "'. Named ranges have been cleaned up.", vbInformation, "Success"

' Clean up
Set srcWorkbook = Nothing
Set destWorkbook = Nothing
Set xlApp = Nothing