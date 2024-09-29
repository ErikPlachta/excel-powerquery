let
    fn_RetryLogic = (queryFunction as function, retries as number) as function =>
    let
        RetryFunction = (params as record, attempt as number) =>
        let
            Result = try queryFunction(params) otherwise
                if attempt < retries then
                    RetryFunction(params, attempt + 1)
                else
                    Diagnostics.LogFailure("All retry attempts failed", () => error "Failed after multiple retries.")
        in
            Result
    in
        (params as record) => RetryFunction(params, 0)
in
    fn_RetryLogic