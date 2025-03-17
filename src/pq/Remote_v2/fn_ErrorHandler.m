let
    fn_ErrorHandler = (queryFunction as function) as function =>
    let
        SafeFunction = (params as record) =>
        let
            // Try to execute the query, return a fallback result in case of error
            Result = try queryFunction(params) otherwise
                Diagnostics.LogFailure("Error in query execution", () => error "Error occurred in query.")
        in
            Result
    in
        SafeFunction
in
    fn_ErrorHandler