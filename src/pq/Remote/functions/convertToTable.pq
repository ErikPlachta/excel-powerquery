let
    convertToTable = (fileContent as table, tableName as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        loadTable = () => Excel.CurrentWorkbook(){[Name=tableName]}[Content],
        result = retryAction(loadTable, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to convert to table.", "convertToTable", logPath) else result
in
    convertToTable