let
    //--------------------------------------------------------------------------
    // Function to process individual records and lists
    ProcessItem = (data as any, parentID as nullable number, currentID as number, groupID as number) =>
    let
        ProcessedFields = if data = null then
            // Return an empty list if data is null
            {}
        else if Value.Is(data, type record) then
            // If the data is a record, get its fields and process them
            let
                FieldNames = Record.FieldNames(data),
                Rows = List.Accumulate(FieldNames, {}, (accum, field) =>
                    let
                        FieldValue = Record.Field(data, field),
                        Row = if Value.Is(FieldValue, type record) then
                            // Recursively process the nested record
                            ProcessItem(FieldValue, currentID, currentID + 1, groupID + 1)
                        else if Value.Is(FieldValue, type list) then
                            // Recursively process the list
                            List.Accumulate(FieldValue, {}, (listAccum, item) =>
                                List.Combine({listAccum, ProcessItem(item, currentID, currentID + 1, groupID + 1)})
                            )
                        else
                            // Process the field as a simple value
                            {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                    in
                        List.Combine({accum, Row})
                )
            in
                Rows
        else if Value.Is(data, type list) then
            // If the data is a list, process each item in the list
            List.Accumulate(data, {}, (accum, item) =>
                List.Combine({accum, ProcessItem(item, parentID, currentID + 1, groupID + 1)})
            )
        else
            // Process the data as a simple value
            {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = "SimpleValue", Value = data]}
    in
        ProcessedFields,

    //--------------------------------------------------------------------------
    // Function to process the entire JSON and generate the table
    ProcessJson = (jsonObject as record) =>
    let
        // Initialize root level variables
        RootID = 1,
        RootGroupID = 1,

        // Start processing the entire JSON object from root
        FinalRows = ProcessItem(jsonObject, null, RootID, RootGroupID),
        FinalTable = Table.FromRecords(FinalRows),

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