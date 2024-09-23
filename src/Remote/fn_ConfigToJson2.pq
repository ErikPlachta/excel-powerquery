let
    //--------------------------------------------------------------------------
    // Function to process individual records and lists iteratively
    ProcessItem = (jsonObject as record) =>
    let
        // Initialize queue with the root JSON object
        Queue = { [data = jsonObject, parentID = null, currentID = 1, groupID = 1] },
        FinalRows = List.Generate(
            () => [Queue = Queue, Processed = {}], // Initial state
            each List.Count([Queue]) > 0, // Continue while there are items in the queue
            each let
                // Dequeue the first item
                CurrentItem = List.First([Queue]),
                RemainingQueue = List.RemoveFirstN([Queue], 1),
                
                // Process the current item
                data = CurrentItem[data],
                parentID = CurrentItem[parentID],
                currentID = CurrentItem[currentID],
                groupID = CurrentItem[groupID],

                // Initialize processed rows for this iteration
                ProcessedRows = if data = null then
                    // Return empty if null
                    {}
                else if Value.Is(data, type record) then
                    // If it's a record, process its fields
                    let
                        FieldNames = Record.FieldNames(data),
                        NewRows = List.Accumulate(FieldNames, {}, (accum, field) =>
                            let
                                FieldValue = Record.Field(data, field),
                                Row = if Value.Is(FieldValue, type record) then
                                    // Add nested record to queue for later processing
                                    [Queue = List.Combine({[Queue], { [data = FieldValue, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})]
                                else if Value.Is(FieldValue, type list) then
                                    // Add each item in the list to the queue
                                    let
                                        ProcessedList = List.Accumulate(FieldValue, {}, (listAccum, item) =>
                                            if Value.Is(item, type record) then
                                                List.Combine({listAccum, { [data = item, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})
                                            else
                                                listAccum
                                        )
                                    in
                                        [Queue = List.Combine({[Queue], ProcessedList})]
                                else
                                    // Simple value, create a row
                                    {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                            in
                                List.Combine({accum, Row})
                        )
                    in
                        NewRows
                else if Value.Is(data, type list) then
                    // If it's a list, process each item
                    let
                        ProcessedList = List.Accumulate(data, {}, (accum, item) =>
                            if Value.Is(item, type record) then
                                List.Combine({accum, { [data = item, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})
                            else
                                accum
                        )
                    in
                        [Queue = List.Combine({[Queue], ProcessedList})]
                else
                    // If it's a simple value
                    {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = "SimpleValue", Value = data}]
            in
                [Queue = RemainingQueue, Processed = List.Combine({[Processed], ProcessedRows})]
        )[Processed],
        FinalTable = Table.FromRecords(FinalRows)
    in
        FinalTable,

    //--------------------------------------------------------------------------
    // Function to process the entire JSON and generate the table
    ProcessJson = (jsonObject as record) =>
    let
        // Process the root JSON object iteratively
        FinalTable = ProcessItem(jsonObject),

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