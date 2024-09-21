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

        // Step 3: Ensure that the files key is treated as a table (if it’s a list, convert it)
        filesList = configJson[Files][files],
        filesTable = if Value.Is(filesList, List.Type) then Table.FromList(filesList, Record.FieldNames(filesList{0})) else filesList,

        // Step 4: Loop through each file in the config and process its actions
        processedFiles = Table.TransformRows(filesTable, each
            let
                fileTitle = [title],
                fileActions = [actions],

                // Ensure the actions field is treated as a table (if it's a list, convert it)
                actionsList = if Value.Is(fileActions, List.Type) then Table.FromList(fileActions, Record.FieldNames(fileActions{0})) else fileActions,

                processFile = Table.TransformRows(actionsList, each
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