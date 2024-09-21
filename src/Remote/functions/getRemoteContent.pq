let
    getRemoteContent = (url as text, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text, logPath as text) as text =>
    let
        validate = validateUrl(url),
        fetchContent = () => Text.FromBinary(Web.Contents(url)),
        result = retryAction(fetchContent, {}, retryAttempts, waitTimeSeconds, failoverBehavior)
    in
        if result = null then logError("Failed to fetch remote content.", "getRemoteContent", logPath) else result
in
    getRemoteContent