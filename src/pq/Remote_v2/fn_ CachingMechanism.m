let
    fn_CachingMechanism = (queryFunction as function, refreshInterval as duration) as function =>
    let
        CachedFunction = (params as record) =>
        let
            // Define a key for the cache based on parameters and time
            CacheKey = Text.FromBinary(Binary.FromText(Text.From(params), BinaryEncoding.Base64)),
            // Check if cache exists and is valid
            Cache = if CacheKey = Text.Combine({"Cache", Text.From(DateTime.FixedLocalNow())}) then
                Diagnostics.LogValue("Cache Hit", Cache)
            else
                let
                    // Cache miss: run the query function and store the result
                    Result = queryFunction(params),
                    // Store result with the current timestamp
                    Cache = Result
                in
                    Result
        in
            Cache
    in
        CachedFunction
in
    fn_CachingMechanism