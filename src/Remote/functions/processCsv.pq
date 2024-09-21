let
    processCsv = (csvContent as text, delimiter as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        validateDelimiter = if Text.Length(delimiter) > 0 then true else error "Invalid delimiter",
        parseCsv = () => Csv.Document(csvContent, [Delimiter=delimiter, Encoding=1252, QuoteStyle=QuoteStyle.None]),
        result = retryAction(parseCsv, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to process CSV.", "processCsv", logPath) else result
in
    processCsv