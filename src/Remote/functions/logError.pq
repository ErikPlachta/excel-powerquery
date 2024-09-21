let
    logError = (errorMessage as text, stepName as text, logPath as text) as record =>
    let
        logEntry = [
            Step = stepName,
            Error = errorMessage,
            Timestamp = DateTime.LocalNow()
        ],
        // Append log entry to a table, or save to file based on logPath
        logResult = try File.Contents(logPath) otherwise "Logging not configured"
    in
        logEntry
in
    logError