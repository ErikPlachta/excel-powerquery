let
    fn_PerformanceLogger = (queryFunction as function) as function =>
    let
        LoggedFunction = (params as record) =>
        let
            // Start time
            StartTime = DateTime.LocalNow(),
            // Run the actual query function
            Result = queryFunction(params),
            // End time
            EndTime = DateTime.LocalNow(),
            // Calculate execution duration
            Duration = Duration.ToText(Duration.From(StartTime - EndTime)),
            // Log the duration
            LoggedResult = Diagnostics.LogValue("Query Execution Duration", Duration),
            FinalResult = LoggedResult
        in
            FinalResult
    in
        LoggedFunction
in
    fn_PerformanceLogger