let
    main = () as table =>
    let
        // Step 1: Fetch the remote main script URL
        remoteMainScriptUrl = "https://example.com/main_script.pq",
        
        // Step 2: Retrieve the content of the remote main script
        remoteMainScript = try Text.FromBinary(Web.Contents(remoteMainScriptUrl))
            otherwise error "Error fetching the remote MAIN script.",

        // Step 3: Evaluate the fetched script dynamically
        result = try Expression.Evaluate(remoteMainScript, #shared)
            otherwise error "Error evaluating the remote MAIN script."
    in
        result
in
    main