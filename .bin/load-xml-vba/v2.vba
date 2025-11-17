Option Explicit

'@module ManifestAddInManager
'@description Module that manages Excel web add-in manifest sideloading and UI commands.
'@usage Call LoadManifestPrompt, RemoveManifestPrompt, ShowRegisteredManifests, and CreateManifestMenu from Excel.
'@resources https://learn.microsoft.com/office/dev/add-ins/publish/host-and-deploy-office-add-ins

Private Type ManifestConfig
    RootRegPath As String
    RegSubKey As String
    LastPathName As String
    MenuCaption As String
    MenuLoadCaption As String
    MenuRemoveCaption As String
    MenuListCaption As String
    MenuLoadFaceId As Long
    MenuRemoveFaceId As Long
    MenuListFaceId As Long
End Type

'@function GetManifestConfig
'@description Provides immutable configuration details for registry and UI behavior.
'@param None
'@returns ManifestConfig – Structure populated with default values.
'@usedBy LoadManifestPrompt, RemoveManifestPrompt, ShowRegisteredManifests, CreateManifestMenu, RemoveManifestMenu
'@example Dim cfg As ManifestConfig: cfg = GetManifestConfig
Private Function GetManifestConfig() As ManifestConfig
    Static cfg As ManifestConfig
    Static isInitialized As Boolean

    If Not isInitialized Then
        cfg.RootRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\"
        cfg.RegSubKey = "\\WEF\\Developer\\"
        cfg.LastPathName = "LastManifestPath"
        cfg.MenuCaption = "Manifest Tools"
        cfg.MenuLoadCaption = "Load Manifest"
        cfg.MenuRemoveCaption = "Remove Manifest"
        cfg.MenuListCaption = "List Registered Manifests"
        cfg.MenuLoadFaceId = 71
        cfg.MenuRemoveFaceId = 74
        cfg.MenuListFaceId = 98
        isInitialized = True
    End If

    GetManifestConfig = cfg
End Function

'@function LoadManifestPrompt
'@description Prompts the user for a manifest and writes the registry entry to sideload it.
'@param None
'@returns Sub – Displays status message boxes instead of returning data.
'@usedBy Menu button created via CreateManifestMenu or manual invocation.
'@example LoadManifestPrompt
Public Sub LoadManifestPrompt()
    Dim manifestPath As String
    manifestPath = PromptForManifestPath()
    If manifestPath = "" Then Exit Sub

    If WriteManifestRegistry(manifestPath) Then
        SaveLastManifestPath manifestPath
        MsgBox "Manifest registered successfully!" & vbCrLf & _
               "Restart Excel to activate the add-in.", vbInformation
    Else
        MsgBox "Error writing manifest to registry.", vbCritical
    End If
End Sub

'@function RemoveManifestPrompt
'@description Prompts the user for a manifest path and deletes the related registry entry.
'@param None
'@returns Sub – Shows success or failure dialog boxes.
'@usedBy Menu button created via CreateManifestMenu or manual invocation.
'@example RemoveManifestPrompt
Public Sub RemoveManifestPrompt()
    Dim manifestPath As String
    manifestPath = PromptForManifestPath()
    If manifestPath = "" Then Exit Sub

    If DeleteManifestRegistry(manifestPath) Then
        SaveLastManifestPath manifestPath
        MsgBox "Manifest removed successfully!" & vbCrLf & _
               "Restart Excel to complete the removal.", vbInformation
    Else
        MsgBox "Manifest not found or removal failed.", vbExclamation
    End If
End Sub

'@function ShowRegisteredManifests
'@description Reads the registry and lists all registered manifest entries.
'@param None
'@returns Sub – Presents the registered entries in a message box.
'@usedBy Menu button created via CreateManifestMenu or manual invocation.
'@example ShowRegisteredManifests
Public Sub ShowRegisteredManifests()
    Dim manifestList As String
    manifestList = QueryRegisteredManifests()

    If manifestList = "" Then
        manifestList = "(No registered manifests found.)"
    End If

    MsgBox "Registered Add-in Manifests:" & vbCrLf & vbCrLf & manifestList, _
           vbInformation, "Registry Entries"
End Sub

'@function CreateManifestMenu
'@description Injects the custom manifest management popup into the Excel Add-ins menu.
'@param None
'@returns Sub – Affects Excel UI only.
'@usedBy Workbook Open or manual macro execution.
'@example CreateManifestMenu
Public Sub CreateManifestMenu()
    Dim cb As CommandBar
    Dim popUp As CommandBarPopup
    Dim cfg As ManifestConfig

    cfg = GetManifestConfig()

    On Error Resume Next
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub

    RemoveManifestMenu

    Set popUp = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    popUp.Caption = cfg.MenuCaption

    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = cfg.MenuLoadCaption
        .OnAction = "LoadManifestPrompt"
        .FaceId = cfg.MenuLoadFaceId
    End With

    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = cfg.MenuRemoveCaption
        .OnAction = "RemoveManifestPrompt"
        .FaceId = cfg.MenuRemoveFaceId
    End With

    With popUp.Controls.Add(Type:=msoControlButton)
        .Caption = cfg.MenuListCaption
        .OnAction = "ShowRegisteredManifests"
        .FaceId = cfg.MenuListFaceId
    End With
End Sub

'@function RemoveManifestMenu
'@description Removes the injected manifest management popup from the Excel Add-ins menu.
'@param None
'@returns Sub – Performs UI cleanup only.
'@usedBy CreateManifestMenu, Workbook_BeforeClose handlers, or manual macros.
'@example RemoveManifestMenu
Public Sub RemoveManifestMenu()
    Dim ctrl As CommandBarControl
    Dim cfg As ManifestConfig

    cfg = GetManifestConfig()

    On Error Resume Next
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Caption = cfg.MenuCaption Then
            ctrl.Delete
        End If
    Next ctrl
    On Error GoTo 0
End Sub

'@function PromptForManifestPath
'@description Presents a file picker dialog seeded with the last used manifest path.
'@param None
'@returns String – Full manifest path or empty string if the user cancels.
'@usedBy LoadManifestPrompt, RemoveManifestPrompt
'@example Dim path As String: path = PromptForManifestPath
Private Function PromptForManifestPath() As String
    Dim fd As FileDialog
    Dim defaultPath As String

    defaultPath = GetLastManifestPath()
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Manifest XML File"
        .Filters.Clear
        .Filters.Add "XML Files", "*.xml"
        If defaultPath <> "" Then .InitialFileName = defaultPath
        If .Show = -1 Then
            PromptForManifestPath = .SelectedItems(1)
        Else
            PromptForManifestPath = ""
        End If
    End With
End Function

'@function SaveLastManifestPath
'@description Persists the most recently used manifest path in a workbook name.
'@param path String – File system path to store.
'@returns Sub – Updates or creates the named range.
'@usedBy LoadManifestPrompt, RemoveManifestPrompt
'@example SaveLastManifestPath "C:\\Manifests\\addin.xml"
Private Sub SaveLastManifestPath(ByVal path As String)
    Dim cfg As ManifestConfig
    Dim n As Name

    cfg = GetManifestConfig()

    On Error Resume Next
    Set n = ThisWorkbook.Names(cfg.LastPathName)
    On Error GoTo 0

    If n Is Nothing Then
        ThisWorkbook.Names.Add Name:=cfg.LastPathName, RefersTo:="=""" & path & """"
    Else
        n.RefersTo = "=""" & path & """"
    End If
End Sub

'@function GetLastManifestPath
'@description Retrieves the stored manifest path from the workbook if present.
'@param None
'@returns String – Stored path or empty when missing.
'@usedBy PromptForManifestPath
'@example Dim lastPath As String: lastPath = GetLastManifestPath
Private Function GetLastManifestPath() As String
    Dim cfg As ManifestConfig
    Dim storedValue As String

    cfg = GetManifestConfig()

    On Error Resume Next
    storedValue = ThisWorkbook.Names(cfg.LastPathName).RefersTo
    On Error GoTo 0

    storedValue = Replace(storedValue, "=", "")
    storedValue = Replace(storedValue, """", "")
    GetLastManifestPath = storedValue
End Function

'@function WriteManifestRegistry
'@description Creates or overwrites a registry entry pointing to the manifest.
'@param manifestPath String – Full path to the manifest file.
'@returns Boolean – True when successful, False if an error occurs.
'@usedBy LoadManifestPrompt
'@example success = WriteManifestRegistry("C:\\Manifests\\addin.xml")
Private Function WriteManifestRegistry(ByVal manifestPath As String) As Boolean
    On Error GoTo Failed
    Dim wsh As Object

    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite BuildRegistryEntryPath(manifestPath), manifestPath, "REG_SZ"
    WriteManifestRegistry = True
    Exit Function
Failed:
    WriteManifestRegistry = False
End Function

'@function DeleteManifestRegistry
'@description Deletes the registry entry for the specified manifest path.
'@param manifestPath String – Manifest path originally used for registration.
'@returns Boolean – True on success, False if deletion fails.
'@usedBy RemoveManifestPrompt
'@example success = DeleteManifestRegistry("C:\\Manifests\\addin.xml")
Private Function DeleteManifestRegistry(ByVal manifestPath As String) As Boolean
    On Error GoTo Failed
    Dim wsh As Object

    Set wsh = CreateObject("WScript.Shell")
    wsh.RegDelete BuildRegistryEntryPath(manifestPath)
    DeleteManifestRegistry = True
    Exit Function
Failed:
    DeleteManifestRegistry = False
End Function

'@function QueryRegisteredManifests
'@description Uses the Windows reg command to return every stored manifest path.
'@param None
'@returns String – Line-delimited list of manifest entries.
'@usedBy ShowRegisteredManifests
'@example MsgBox QueryRegisteredManifests
Private Function QueryRegisteredManifests() As String
    On Error GoTo Failed
    Dim cfg As ManifestConfig
    Dim wsh As Object
    Dim cmd As String
    Dim execObj As Object
    Dim line As String
    Dim output As String

    cfg = GetManifestConfig()
    cmd = "reg query """ & GetRegistryQueryPath() & cfg.RegSubKey & """"

    Set wsh = CreateObject("WScript.Shell")
    Set execObj = wsh.Exec("cmd /c " & cmd)

    Do While Not execObj.StdOut.AtEndOfStream
        line = execObj.StdOut.ReadLine
        If InStr(1, LCase$(line), ".xml", vbTextCompare) > 0 Then
            output = output & line & vbCrLf
        End If
    Loop

    QueryRegisteredManifests = Trim$(output)
    Exit Function
Failed:
    QueryRegisteredManifests = ""
End Function

'@function BuildRegistryEntryPath
'@description Constructs the full registry entry path for the manifest.
'@param manifestPath String – Manifest file path.
'@returns String – Full registry path suitable for RegWrite/RegDelete.
'@usedBy WriteManifestRegistry, DeleteManifestRegistry
'@example fullPath = BuildRegistryEntryPath("C:\\Manifests\\addin.xml")
Private Function BuildRegistryEntryPath(ByVal manifestPath As String) As String
    Dim cfg As ManifestConfig
    cfg = GetManifestConfig()
    BuildRegistryEntryPath = GetRegistryBasePath() & cfg.RegSubKey & manifestPath
End Function

'@function GetRegistryBasePath
'@description Determines the base registry hive for the current Excel version.
'@param None
'@returns String – e.g., HKEY_CURRENT_USER\Software\Microsoft\Office\16.0
'@usedBy BuildRegistryEntryPath, QueryRegisteredManifests
'@example Debug.Print GetRegistryBasePath
Private Function GetRegistryBasePath() As String
    Dim cfg As ManifestConfig
    cfg = GetManifestConfig()
    GetRegistryBasePath = cfg.RootRegPath & Application.Version
End Function

'@function GetRegistryQueryPath
'@description Converts the long registry path into the short HKCU-prefixed path for cmd usage.
'@param None
'@returns String – e.g., HKCU\Software\Microsoft\Office\16.0
'@usedBy QueryRegisteredManifests
'@example Debug.Print GetRegistryQueryPath
Private Function GetRegistryQueryPath() As String
    Dim basePath As String
    basePath = GetRegistryBasePath()
    GetRegistryQueryPath = "HKCU" & Mid$(basePath, Len("HKEY_CURRENT_USER") + 1)
End Function