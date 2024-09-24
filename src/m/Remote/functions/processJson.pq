let
    processJson = (jsonContent as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        parseJson = () => Json.Document(jsonContent),
        result = retryAction(parseJson, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to process JSON.", "processJson", logPath) else result
in
    processJson