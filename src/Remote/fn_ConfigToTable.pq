let
    // Modular function to dynamically generate a config for any JSON object
    GenerateDynamicConfig = (jsonObject as record, parentKeys as nullable list) =>
    let
        // Iterate over each key in the current JSON object and accumulate paths
        Fields = Record.FieldNames(jsonObject),
        DynamicConfig = List.Accumulate(Fields, {}, (accumulatedConfig, field) =>
            let
                fieldValue = Record.Field(jsonObject, field),
                // Build the full path (parent + current field)
                newParentKeys = if parentKeys = null then {field} else List.Combine({parentKeys, {field}}),
                // Recursively generate config for nested records and lists
                nestedConfig = if Value.Is(fieldValue, type record) then
                    GenerateDynamicConfig(fieldValue, newParentKeys)
                else if Value.Is(fieldValue, type list) then
                    List.Accumulate(fieldValue, accumulatedConfig, (listAccum, item) =>
                        if Value.Is(item, type record) then
                            // Recursively process list items if they are records
                            List.Combine({listAccum, GenerateDynamicConfig(item, newParentKeys)})
                        else
                            listAccum
                    )
                else
                    {[title = Text.Combine(newParentKeys, "."), target = newParentKeys]}
            in
                List.Combine({accumulatedConfig, nestedConfig})
        )
    in
        DynamicConfig,

    // Modular function to validate ParentID and ensure no orphan rows
    ValidateParentID = (table as table) =>
    let
        OrphanRows = Table.SelectRows(table, each [ParentID] <> null and Table.IsEmpty(Table.SelectRows(table, each [ID] = [ParentID]))),
        ValidationResult = if Table.IsEmpty(OrphanRows) then
            "ParentID validation passed"
        else
            Table.ToText(OrphanRows, "Orphan Rows Found: ", " | ")
    in
        ValidationResult,

    // Modular function to extract fields based on the config and apply recursion, including Group ID
    ExtractFields = (data as record, config as list, optional parentID as nullable number, optional currentID as number, optional groupID as nullable number) =>
    let
        // Assign a Group ID for the current object (if not already assigned)
        newGroupID = if groupID = null then if currentID = null then 1 else currentID + 1 else groupID,
        
        // Assign current ID for the current row
        newID = if currentID = null then 1 else currentID + 1,

        // Extract each field based on the provided config
        Rows = List.Transform(config, (fieldConfig) =>
            let
                fieldTitle = fieldConfig[title],
                fieldValue = try Record.Field(data, List.First(fieldConfig[target])) otherwise null,
                row = [
                    ID = newID,
                    ParentID = parentID,
                    GroupID = newGroupID,
                    Title = fieldTitle,
                    Value = fieldValue
                ]
            in
                row
        ),

        // Recursively process nested records and lists
        NestedRows = List.Accumulate(Record.FieldNames(data), Rows, (accumulatedRows, field) =>
            let
                fieldValue = Record.Field(data, field),
                childRows = if Value.Is(fieldValue, type record) then
                    // Recursively process nested records
                    ExtractFields(fieldValue, config, newID, newID, newGroupID)
                else if Value.Is(fieldValue, type list) then
                    // Process list items recursively
                    List.Accumulate(fieldValue, accumulatedRows, (listAccum, item) =>
                        if Value.Is(item, type record) then
                            List.Combine({listAccum, ExtractFields(item, config, newID, newID, newGroupID)})
                        else
                            listAccum
                    )
                else
                    accumulatedRows
            in
                Table.Combine({accumulatedRows, childRows})
        )
    in
        Table.FromRecords(NestedRows),

    // Main function to process the JSON object with optional config
    ProcessJson = (jsonObject as record, optional config as nullable list) =>
    let
        // If no config is passed, dynamically generate it by traversing the JSON
        configToUse = if config = null then
            GenerateDynamicConfig(jsonObject, null)
        else
            config,

        // Process the main JSON object using the extracted config
        FinalTable = ExtractFields(jsonObject, configToUse),

        // Validate the ParentID assignments to catch any orphan rows
        ValidationResult = ValidateParentID(FinalTable)
    in
        if ValidationResult = "ParentID validation passed" then
            FinalTable
        else
            error ValidationResult
in
    ProcessJson