let
    // Parameter-driven function to handle diagnostics dynamically
    DynamicDiagnosticFunction = (
        queryFunction as function, 
        traceLevel as number, 
        logStepNames as list, 
        logValues as list
    ) as function =>
    let
        // Wrap the query steps in Diagnostics.Trace based on logStepNames
        WrappedFunction = List.Accumulate(logStepNames, queryFunction, (currentFunc, stepName) =>
            let
                // Wrap each step in a Diagnostics.Trace
                WrappedStep = (params as record) =>
                    let
                        result = try Record.Field(currentFunc(params), stepName) otherwise null,
                        logMessage = Text.Format("Step '#{0}' executed.", {stepName})
                    in
                        Diagnostics.Trace(traceLevel, logMessage, result)
            in
                WrappedStep
        ),
        
        // Optionally log values at specific steps based on logValues
        LogValuesFunction = List.Accumulate(logValues, WrappedFunction, (currentFunc, valueToLog) =>
            let
                LoggedValue = (params as record) =>
                    let
                        value = try Record.Field(currentFunc(params), valueToLog) otherwise null,
                        logMessage = Text.Format("Logged value at step '#{0}': #(#)", {valueToLog, Text.From(value)})
                    in
                        Diagnostics.LogValue(logMessage, value)
            in
                LoggedValue
        )
    in
        LogValuesFunction
in
    DynamicDiagnosticFunction