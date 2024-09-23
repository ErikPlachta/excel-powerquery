let
    //--------------------------------------------------------------------------
    // Modular function to dynamically generate a config for any JSON object
    GenerateDynamicConfig = (
        jsonObject as record
        ,parentKeys as nullable list
    ) =>
    let
        // Iterate over each key in the current JSON object and accumulate paths
        Fields = Record.FieldNames(jsonObject)
        ,DynamicConfig = 
           List.Accumulate(
            Fields, {}, (accumulatedConfig, field
          ) =>
          let
              tryFieldValue = try Record.Field(jsonObject, field),
              fieldValue = if tryFieldValue[HasError] then 
                   error "Error retrieving the field '" & field & "' at Line 12: " & tryFieldValue[Error][Message]
               else 
                   tryFieldValue[Value],

              // Build the full path (parent + current field)
              newParentKeys = 
                if (parentKeys = null) then
                 {field}
                else
                 List.Combine({parentKeys, {field}})

              // Recursively generate config for nested records and lists
              ,nestedConfig =
                if Value.Is(fieldValue, type record) then
                  let
                      tryGenerateConfig = try GenerateDynamicConfig(fieldValue, newParentKeys)
                  in
                      if tryGenerateConfig[HasError] then
                          error "Error generating config for nested record '" & field & "' at Line 20, path: " & Text.Combine(newParentKeys, ".") & ": " & tryGenerateConfig[Error][Message]
                      else
                          tryGenerateConfig[Value]
              else if Value.Is(fieldValue, type list) then
                  let
                      tryListAccumulate = 
                        try List.Accumulate(
                          fieldValue
                          ,accumulatedConfig
                          ,(listAccum, item) =>
                            if Value.Is(item, type record) then
                                List.Combine({listAccum, GenerateDynamicConfig(item, newParentKeys)})
                            else
                                listAccum
                        )
                  in
                      if tryListAccumulate[HasError] then
                          error "Error processing list field '" & field & "' at Line 32, path: " & Text.Combine(newParentKeys, ".") & ": " & tryListAccumulate[Error][Message]
                      else
                          tryListAccumulate[Value]
              else
                  {[
                    title = Text.Combine(newParentKeys, ".")
                    ,target = newParentKeys
                  ]}
          in
              List.Combine({accumulatedConfig, nestedConfig})
        )
    in
        DynamicConfig,
    //--------------------------------------------------------------------------
    // Modular function to validate ParentID and ensure no orphan rows
    ValidateParentID = (
      table as table
    ) =>
    let
        // Identify orphan rows where ParentID doesn't match any existing ID
        OrphanRows = try Table.SelectRows(
            table,
            each [ParentID] <> null and Table.IsEmpty(Table.SelectRows(table, each [ID] = [ParentID]))
        ) otherwise error "Error in ParentID validation at Line 52: Ensure the input is a valid table.",

        // Provide validation result
        ValidationResult = 
        if Table.IsEmpty(OrphanRows) then
            "ParentID validation passed"
        else
            Table.ToText(OrphanRows, "Orphan Rows Found: ", " | ")
    in
        ValidationResult,
    //--------------------------------------------------------------------------
    // Modular function to extract fields based on the config and apply recursion, including Group ID
    ExtractFields = (
      data as record
      ,config as list
      ,optional parentID as nullable number
      ,optional currentID as number
      ,optional groupID as nullable number
    ) =>
    let
        // Assign a Group ID for the current object (if not already assigned)
        newGroupID = if groupID = null then if currentID = null then 1 else currentID + 1 else groupID,
        //----------------------------------------------------------------------
        // Assign current ID for the current row
        newID = if currentID = null then 1 else currentID + 1,
        //----------------------------------------------------------------------
        // Extract each field based on the provided config
        Rows = List.Transform(config, (fieldConfig) =>
            let
                tryFieldTitle = try fieldConfig[title],
                fieldTitle = if tryFieldTitle[HasError] then
                    error "Missing 'title' key in config for current item at Line 68: " & tryFieldTitle[Error][Message]
                else 
                    tryFieldTitle[Value],
                
                tryFieldValue = try Record.Field(data, List.First(fieldConfig[target])),
                fieldValue = if tryFieldValue[HasError] then
                    error "Error retrieving value for target path '" & Text.Combine(List.First(fieldConfig[target]), ".") & "' at Line 73: " & tryFieldValue[Error][Message]
                else 
                    tryFieldValue[Value],

                row = [
                    ID = newID,
                    ParentID = parentID,
                    GroupID = newGroupID,
                    Title = fieldTitle,
                    Value = fieldValue
                ]
            in
                row
        )
        //--------------------------------------------------------------------------
        // Recursively process nested records and lists
        ,NestedRows = List.Accumulate(Record.FieldNames(data), Rows, (accumulatedRows, field) =>
            let
                tryFieldValue = try Record.Field(data, field),
                fieldValue = if tryFieldValue[HasError] then
                    error "Error retrieving field '" & field & "' while processing nested rows at Line 82: " & tryFieldValue[Error][Message]
                else
                    tryFieldValue[Value],

                childRows = if Value.Is(fieldValue, type record) then
                    let
                        tryExtractFields = try ExtractFields(fieldValue, config, newID, newID, newGroupID)
                    in
                        if tryExtractFields[HasError] then
                            error "Error processing nested record for field '" & field & "' at Line 88: " & tryExtractFields[Error][Message]
                        else
                            tryExtractFields[Value]
                else if Value.Is(fieldValue, type list) then
                    let
                        tryListAccumulate = try List.Accumulate(fieldValue, accumulatedRows, (listAccum, item) =>
                            if Value.Is(item, type record) then
                                List.Combine({listAccum, ExtractFields(item, config, newID, newID, newGroupID)})
                            else
                                List.Combine({listAccum, {[ID = newID, ParentID = parentID, GroupID = newGroupID, Title = field, Value = item]}})
                        )
                    in
                        if tryListAccumulate[HasError] then
                            error "Error processing list items for field '" & field & "' at Line 94: " & tryListAccumulate[Error][Message]
                        else
                            tryListAccumulate[Value]
                else
                    accumulatedRows
            in
                Table.FromRecords(accumulatedRows)
        )
    in
        Table.FromRecords(NestedRows),
    //--------------------------------------------------------------------------
    // Main function to process the JSON object with optional config
    ProcessJson = (jsonObject as record, optional config as nullable list) =>
    let
        //----------------------------------------------------------------------
        // If no config is passed, dynamically generate it by traversing the JSON
        tryConfigToUse = try if config = null then
            GenerateDynamicConfig(jsonObject, null)
        else
            config,
        
        configToUse = if tryConfigToUse[HasError] then
            error "Error generating or using config at Line 109: " & tryConfigToUse[Error][Message]
        else
            tryConfigToUse[Value],

        // Process the main JSON object using the extracted config
        tryFinalTable = try ExtractFields(jsonObject, configToUse),
        FinalTable = if tryFinalTable[HasError] then
            error "Error extracting fields from JSON object at Line 114: " & tryFinalTable[Error][Message]
        else
            tryFinalTable[Value],
        //----------------------------------------------------------------------
        // Validate the ParentID assignments to catch any orphan rows
        tryValidationResult = try ValidateParentID(FinalTable)
        ,ValidationResult = 
          if tryValidationResult[HasError] then
            error "Error during ParentID validation at Line 119: " & tryValidationResult[Error][Message]
         else
              tryValidationResult[Value]
    in
        if ValidationResult = "ParentID validation passed" then
            FinalTable
        else
            error ValidationResult
in
    ProcessJson