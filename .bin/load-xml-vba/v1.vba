Option Explicit

Private Const ROOT_REG_PATH As String = "HKEY_CURRENT_USER\Software\Microsoft\Office\"
Private Const REG_SUB_KEY As String = "\WEF\Developer\"
Private Const NAMEDRANGE_LAST_PATH As String = "LastManifestPath"

' Entry Point: Load Manifest
Public Sub LoadManifestPrompt()
    Dim path As String
    path = GetManifestPathFromUser()
    If path = "" Then Exit Sub

    If WriteManifestRegistry(path) Then
        SaveLastManifestPath path
        MsgBox "Manifest registered successfully!" & vbCrLf & _
               "Restart Excel to activate the add-in.", vbInformation
    Else
        MsgBox "Error writing manifest to registry.", vbCritical
    End If
End Sub

' Entry Point: Remove Manifest
Public Sub RemoveManifestPrompt()
    Dim path As String
    path = GetManifestPathFromUser()
    If path = "" Then Exit Sub

    If DeleteManifestRegistry(path) Then
        SaveLastManifestPath path
        MsgBox "Manifest removed successfully!" & vbCrLf & _
               "Restart Excel to complete the removal.", vbInformation
    Else
        MsgBox "Manifest not found or removal failed.", vbExclamation
    End If
End Sub

' Entry Point: Show All Registered Manifests
Public Sub ShowRegisteredManifests()
    On Error GoTo Failed

    Dim wsh As Object, regKey As String
    Dim output As String, versionKey As String
    versionKey = GetOfficeRegistryKey()
    regKey = "HKCU\" & Mid(versionKey, Len("HKEY_CURRENT_USER\") + 1) & REG_SUB_KEY

    Set wsh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "reg query """ & regKey & """"

    Dim exec As Object
    Set exec = CreateObject("WScript.Shell").Exec("cmd /c " & cmd)

    Dim line As String
    Do While Not exec.StdOut.AtEndOfStream
        line = exec.StdOut.ReadLine
        If InStr(line, ".xml") > 0 Then output = output & line & vbCrLf
    Loop

    If output = "" Then output = "(No registered manifests found.)"

    MsgBox "Registered Add-in Manifests:" & vbCrLf & vbCrLf & output, vbInformation, "Registry Entries"
    Exit Sub
Failed:
    MsgBox "Failed to read registry keys. Run Excel as Administrator if needed.", vbCritical
End Sub

' Registry Write
Private Function WriteManifestRegistry(manifestPath As String) As Boolean
    On Error GoTo Failed
    Dim wsh As Object
    Dim regPath As String

    Set wsh = CreateObject("WScript.Shell")
    regPath = GetOfficeRegistryKey() & REG_SUB_KEY & manifestPath
    wsh.RegWrite regPath, manifestPath, "REG_SZ"
    WriteManifestRegistry = True
    Exit Function
Failed:
    WriteManifestRegistry = False
End Function

' Registry Delete
Private Function DeleteManifestRegistry(manifestPath As String) As Boolean
    On Error GoTo Failed
    Dim wsh As Object
    Dim regPath As String

    Set wsh = CreateObject("WScript.Shell")
    regPath = GetOfficeRegistryKey() & REG_SUB_KEY & manifestPath
    wsh.RegDelete regPath
    DeleteManifestRegistry = True
    Exit Function
Failed:
    DeleteManifestRegistry = False
End Function

' Prompt for manifest file
Private Function GetManifestPathFromUser() As String
    Dim fd As FileDialog
    Dim defaultPath As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    defaultPath = GetLastManifestPath()

    With fd
        .Title = "Select Manifest XML File"
        .Filters.Clear
        .Filters.Add "XML Files", "*.xml"
        If defaultPath <> "" Then .InitialFileName = defaultPath
        If .Show = -1 Then
            GetManifestPathFromUser = .SelectedItems(1)
        Else
            GetManifestPathFromUser = ""
        End If
    End With
End Function

' Store last manifest path in workbook (hidden range)
Private Sub SaveLastManifestPath(path As String)
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(NAMEDRANGE_LAST_PATH)
    If n Is Nothing Then
        ThisWorkbook.Names.Add Name:=NAMEDRANGE_LAST_PATH, RefersTo:="=""" & path & """"
    Else
        n.RefersTo = "=""" & path & """"
    End If
    On Error GoTo 0
End Sub

' Retrieve last used path
Private Function GetLastManifestPath() As String
    On Error Resume Next
    GetLastManifestPath = ThisWorkbook.Names(NAMEDRANGE_LAST_PATH).RefersTo
    GetLastManifestPath = Replace(GetLastManifestPath, "=", "")
    GetLastManifestPath = Replace(GetLastManifestPath, """", "")
    On Error GoTo 0
End Function

' Office registry root key by version
Private Function GetOfficeRegistryKey() As String
    Dim ver As String
    ver = Application.Version ' e.g. "16.0"
    GetOfficeRegistryKey = ROOT_REG_PATH & ver
End Function

' Create Add-ins tab menu
Public Sub CreateManifestMenu()
    Dim cb As CommandBar
    Dim popUp As CommandBarPopup

    On Error Resume Next
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    If cb Is Nothing Then Exit Sub

    Call RemoveManifestMenu

    Set popUp = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    popUp.Caption = "Manifest Tools"
    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = "Load Manifest"
        .OnAction = "LoadManifestPrompt"
        .FaceId = 71
    End With
    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = "Remove Manifest"
        .OnAction = "RemoveManifestPrompt"
        .FaceId = 74
    End With
    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = "List Registered Manifests"
        .OnAction = "ShowRegisteredManifests"
        .FaceId = 98
    End With
End Sub

' Cleanup menu
Public Sub RemoveManifestMenu()
    Dim ctrl As CommandBarControl
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Caption = "Manifest Tools" Then ctrl.Delete
    Next ctrl
End Sub