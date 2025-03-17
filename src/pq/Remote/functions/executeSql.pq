let
    executeSql = (connection as text, sqlQuery as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        validateConnection = if Text.Length(connection) > 0 then true else error "Invalid connection string",
        executeQuery = () => Sql.Database(connection, sqlQuery),
        result = retryAction(executeQuery, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to execute SQL query.", "executeSql", logPath) else result
in
    executeSql