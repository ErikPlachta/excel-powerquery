let
    generateSettingsTable = (headers as list, values as list) as table =>
    let
        // Combine headers and values into records
        settingsRecords = List.Transform({headers, values}, each Record.FromList(_, headers)),
        
        // Create table from the records
        settingsTable = Table.FromRecords(settingsRecords),

        // Check if the Settings table exists, if not create it
        existingTable = try Excel.CurrentWorkbook(){[Name="Settings"]}[Content] otherwise Table.FromRows({}, headers),
        finalTable = Table.Combine({existingTable, settingsTable})
    in
        finalTable
in
    generateSettingsTable