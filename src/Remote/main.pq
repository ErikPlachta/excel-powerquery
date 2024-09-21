let
    main = () as table =>
    let
        // Step 1: Fetch the remote configuration file
        configUrl = "https://example.com/config.json",
        configSource = try Text.FromBinary(Web.Contents(configUrl))
            otherwise error "Error fetching the configuration file.",

        // Step 2: Parse the configuration JSON
        configJson = try Json.Document(configSource)
            otherwise error "Error parsing the configuration JSON.",

        // Step 3: Loop through each file in the config and execute its actions
        processedFiles = Table.TransformRows(configJson[Files][files], each
            let
                fileTitle = [title],
                fileActions = [actions],
                processFile = Table.TransformRows(fileActions, each
                    let
                        actionFunction = Expression.Evaluate([function], #shared),
                        actionParams = [parameters],
                        actionResult = try actionFunction(actionParams)
                            otherwise error "Error executing action: " & [function]
                    in
                        actionResult
                )
            in
                processFile
        )
    in
        processedFiles
in
    main