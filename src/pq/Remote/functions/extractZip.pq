let
    extractZip = (zipFileContent as binary, extractPath as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as table =>
    let
        validateExtractPath = validateFilePath(extractPath),
        extractFiles = () => Zip.Contents(zipFileContent, extractPath),
        result = retryAction(extractFiles, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to extract zip file.", "extractZip", logPath) else result
in
    extractZip