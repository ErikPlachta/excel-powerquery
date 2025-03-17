let
    executeDax = (connection as text, daxQuery as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        executeQuery = () => AnalysisServices.Execute(connection, daxQuery),
        result = retryAction(executeQuery, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to execute DAX query.", "executeDax", logPath) else result
in
    executeDax