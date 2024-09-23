let
    //--------------------------------------------------------------------------
    // Centralized function to handle recursion for both records and lists
    ProcessItem = (
        data as any,
        config as list,
        parentID as nullable number,
        currentID as number,
        groupID as number
    ) =>
    let
        // Check if the data is null, if so, return an empty list (no processing)
        Result = if data = null then
            {}
        else
            let
                // If the data is a record, get the field names, otherwise, return an empty list
                FieldNames = if Value.Is(data, type record) then Record.FieldNames(data) else {},

                // Extract values based on the config or directly process fields
                ProcessedFields = List.Transform(FieldNames, (field) =>
                    let
                        FieldValue = Record.Field(data, field),
                        Row = if Value.Is(FieldValue, type record) then
                            // Recursively process nested records
                            ProcessItem(FieldValue, config, currentID, currentID + 1, groupID + 1)
                        else if Value.Is(FieldValue, type list) then
                            // Handle lists separately, passing each item recursively
                            List.Accumulate(FieldValue, {}, (accum, item) =>
                                List.Combine({accum, ProcessItem(item, config, currentID, currentID + 1, groupID + 1)})
                            )
                        else
                            // Process individual fields as rows
                            {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                    in
                        Row
                ),

                // Combine the rows into one table
                FinalResult = Table.FromRecords(List.Combine(ProcessedFields))
            in
                FinalResult
    in
        Result,

    //--------------------------------------------------------------------------
    // Function to process the entire JSON and generate the table
    ProcessJson = (jsonObject as record, config as list) =>
    let
        // Initialize root level variables
        RootID = 1,
        RootGroupID = 1,

        // Start processing the entire JSON object from root
        FinalTable = ProcessItem(jsonObject, config, null, RootID, RootGroupID),

        // Validate that ParentID relationships are valid
        ValidateParentID = (table as table) =>
        let
            OrphanRows = Table.SelectRows(table, each [ParentID] <> null and Table.IsEmpty(Table.SelectRows(table, each [ID] = [ParentID]))),
            ValidationResult = if Table.IsEmpty(OrphanRows) then "ParentID validation passed" else "Orphan Rows Found"
        in
            ValidationResult,

        // Perform validation
        ValidationMessage = ValidateParentID(FinalTable),

        // Return result if validation passes
        Output = if ValidationMessage = "ParentID validation passed" then FinalTable else error ValidationMessage
    in
        Output
in
    ProcessJson