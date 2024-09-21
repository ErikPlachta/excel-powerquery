let
    getLocalFile = (filePath as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text, useLocalCache as logical, remoteContentFunction as function) as text =>
    let
        validate = validateFilePath(filePath),
        fetchContent = () => manageCache(filePath, remoteContentFunction, useLocalCache),
        result = retryAction(fetchContent, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to fetch local file.", "getLocalFile", logPath) else result
in
    getLocalFile