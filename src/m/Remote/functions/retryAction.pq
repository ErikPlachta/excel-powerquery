let
    retryAction = (action as function, parameters as list, retryAttempts as number, waitTimeSeconds as number, failoverBehavior as text) as any =>
    let
        executeWithRetry = List.Accumulate({1..retryAttempts}, null, (state, attempt) =>
            if state <> null then state 
            else try action(parameters) otherwise (
                if attempt < retryAttempts then 
                    (Function.InvokeAfter(() => null, #duration(0, 0, 0, waitTimeSeconds)); null) 
                else 
                    if failoverBehavior = "stop" then error "Action failed after " & Text.From(attempt) & " attempts."
                    else if failoverBehavior = "logAndContinue" then logError("Action failed after " & Text.From(attempt) & " attempts.", "action", "[logPath]")
                    else null
            )
        )
    in
        executeWithRetry
in
    retryAction