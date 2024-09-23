let
    //--------------------------------------------------------------------------
    // Centralized function to handle recursion for both records and lists
    ProcessItem = (
        data as any, 
        parentID as nullable number, 
        currentID as number, 
        groupID as number
    ) =>
    let
        // Initialize empty list to hold processed rows
        Result = if data = null then
            {}
        else if Value.Is(data, type record) then
            // If data is a record, process its fields
            let
                FieldNames = Record.FieldNames(data),
                ProcessedFields = List.Accumulate(FieldNames, {}, (accum, field) =>
                    let
                        FieldValue = Record.Field(data, field),
                        Row = if Value.Is(FieldValue, type record) then
                            // If it's a nested record, recursively process it
                            ProcessItem(FieldValue, currentID, currentID + 1, groupID + 1)
                        else if Value.Is(FieldValue, type list) then
                            // If it's a list, process each item in the list
                            List.Accumulate(FieldValue, {}, (listAccum, item) =>
                                List.Combine({listAccum, ProcessItem(item, currentID, currentID + 1, groupID + 1)})
                            )
                        else
                            // It's a simple field, so process it as a row
                            {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                    in
                        List.Combine({accum, Row})
                ),
                FinalResult = List.Combine(ProcessedFields)
            in
                FinalResult
        else if Value.Is(data, type list) then
            // If data is a list, process each item in the list
            List.Accumulate(data, {}, (accum, item) =>
                List.Combine({accum, ProcessItem(item, parentID, currentID + 1, groupID + 1)})
            )
        else
            // If it's a simple value, return as a row
            {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = "SimpleValue", Value = data]}
    in
        Result,
        
    //--------------------------------------------------------------------------
    // Function to process the entire JSON and generate the table
    ProcessJson = (jsonObject as record) =>
    let
        // Initialize root level variables
        RootID = 1,
        RootGroupID = 1,

        // Start processing the entire JSON object from root
        FinalTable = Table.FromRecords(ProcessItem(jsonObject, null, RootID, RootGroupID)),

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