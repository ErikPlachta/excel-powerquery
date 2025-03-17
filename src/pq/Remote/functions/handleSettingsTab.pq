let
    // Check if the "Settings" tab exists
    SettingsTab = try Excel.CurrentWorkbook(){[Name="Settings"]}[Content] otherwise null,

    // Load the settings from the workbook if present, otherwise from the remote config
    Settings = if SettingsTab <> null then
        SettingsTab
    else
        Json.Document(Web.Contents("https://example.com/config.json"))[Settings]
in
    Settings