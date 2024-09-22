let
    // Function to create or update an Excel table
    CreateOrUpdateTable = (tableName as text, tableData as table, optional sheetName as text) =>
    let
        // Check if the table already exists
        existingTables = Excel.CurrentWorkbook(),
        tableExists = List.Contains(existingTables[Name], tableName),

        // If table exists, replace its data
        result = if tableExists then
            // Overwrite the existing table with the new data
            Table.Replace(tableData, tableName)
        else
            // Create a new table
            Excel.CurrentWorkbook(){[Name = tableName]}[Content] = tableData,

        // Return the new or updated table
        finalResult = if sheetName <> null then 
            Excel.CurrentWorkbook(){[Name = tableName]} otherwise null
    in
        finalResult
in
    CreateOrUpdateTable