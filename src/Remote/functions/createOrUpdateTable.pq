let
    // Function to create a table in the workbook or update it with new data
    CreateOrUpdateTable = (tableName as text, tableData as table) =>
    let
        // Check if the table already exists in Excel
        existingTables = Excel.CurrentWorkbook(),
        tableExists = List.Contains(existingTables[Name], tableName),

        // If the table exists, return the existing table data, else return the new data to be loaded as a table
        result = if tableExists then
            tableData
        else
            tableData
    in
        result
in
    CreateOrUpdateTable