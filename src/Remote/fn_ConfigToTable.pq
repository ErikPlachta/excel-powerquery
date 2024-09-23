let
    // Modular function to dynamically generate a config for any JSON object
    GenerateDynamicConfig = (jsonObject as record, parentKeys as nullable list) =>
    let
        // Iterate over each key in the current JSON object and accumulate paths
        Fields = Record.FieldNames(jsonObject),
        DynamicConfig = List.Accumulate(Fields, {}, (accumulatedConfig, field) =>
            let
                fieldValue = try Record.Field(jsonObject, field) otherwise 
                    error "Error retrieving the field '" & field & "' during config generation.",
                
                // Build the full path (parent + current field)
                newParentKeys = if parentKeys = null then {field} else List.Combine({parentKeys, {field}}),
                
                // Recursively generate config for nested records and lists
                nestedConfig = if Value.Is(fieldValue, type record) then
                    try GenerateDynamicConfig(fieldValue, newParentKeys) otherwise 
                        error "Error generating config for nested record '" & field & "' at path: " & Text.Combine(newParentKeys, "."),
                else if Value.Is(fieldValue, type list) then
                    try List.Accumulate(fieldValue, accumulatedConfig, (listAccum, item) =>
                        if Value.Is(item, type record) then
                            // Recursively process list items if they are records
                            List.Combine({listAccum, GenerateDynamicConfig(item, newParentKeys)})
                        else
                            listAccum
                    ) otherwise error "Error processing list field '" & field & "' at path: " & Text.Combine(newParentKeys, "."),
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
                fieldTitle = try fieldConfig[title] otherwise error "Missing 'title' key in config for current item.",
                fieldValue = try Record.Field(data, List.First(fieldConfig[target])) otherwise 
                    error "Error retrieving value for target path '" & Text.Combine(List.First(fieldConfig[target]), ".") & "'.",
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
                fieldValue = try Record.Field(data, field) otherwise 
                    error "Error retrieving field '" & field & "' while processing nested rows.",
                
                childRows = if Value.Is(fieldValue, type record) then
                    try ExtractFields(fieldValue, config, newID, newID, newGroupID) otherwise 
                        error "Error processing nested record for field '" & field & "'.",
                else if Value.Is(fieldValue, type list) then
                    try List.Accumulate(fieldValue, accumulatedRows, (listAccum, item) =>
                        if Value.Is(item, type record) then
                            List.Combine({listAccum, ExtractFields(item, config, newID, newID, newGroupID)})
                        else
                            // Add list items as separate rows
                            List.Combine({listAccum, {[ID = newID, ParentID = parentID, GroupID = newGroupID, Title = field, Value = item]}})
                    ) otherwise error "Error processing list items for field '" & field & "'.",
                else
                    accumulatedRows
            in
                Table.FromRecords(accumulatedRows)
        )
    in
        Table.FromRecords(NestedRows),

    // Main function to process the JSON object with optional config
    ProcessJson = (jsonObject as record, optional config as nullable list) =>
    let
        // If no config is passed, dynamically generate it by traversing the JSON
        configToUse = try if config = null then
            GenerateDynamicConfig(jsonObject, null)
        else
            config otherwise error "Error generating or using config.",

        // Process the main JSON object using the extracted config
        FinalTable = try ExtractFields(jsonObject, configToUse) otherwise 
            error "Error extracting fields from JSON object.",

        // Validate the ParentID assignments to catch any orphan rows
        ValidationResult = try ValidateParentID(FinalTable) otherwise 
            error "Error during ParentID validation."
    in
        if ValidationResult = "ParentID validation passed" then
            FinalTable
        else
            error ValidationResult
in
    ProcessJson