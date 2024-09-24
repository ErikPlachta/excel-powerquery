let
    ProcessJson = (jsonObject as record) =>
    let
        //--------------------------------------------------------------------------
        // Function to process individual records and lists iteratively
        ProcessItem = (jsonObject as record) =>
        let
            Queue = { [data = jsonObject, parentID = null, currentID = 1, groupID = 1] },
            FinalRows = List.Generate(
                () => [Queue = Queue, Processed = {}], // Initial state
                each List.Count([Queue]) > 0, // Continue while there are items in the queue
                each let
                    CurrentItem = List.First([Queue]),
                    RemainingQueue = List.RemoveFirstN([Queue], 1),
                    data = CurrentItem[data],
                    parentID = CurrentItem[parentID],
                    currentID = CurrentItem[currentID],
                    groupID = CurrentItem[groupID],

                    // Process the current item by checking its type
                    ProcessedRows = if data = null then
                        {}
                    else if Value.Is(data, type record) then
                        ProcessRecord(data, currentID, parentID, groupID)
                    else if Value.Is(data, type list) then
                        ProcessList(data, currentID, parentID, groupID)
                    else
                        ProcessSimpleValue(data, currentID, parentID, groupID)
                in
                    [Queue = RemainingQueue, Processed = List.Combine({[Processed], ProcessedRows})]
            )[Processed],

            FinalTable = Table.FromRecords(FinalRows)
        in
            FinalTable,

        //--------------------------------------------------------------------------
        // Function to process records (objects with named fields)
        ProcessRecord = (data as record, currentID as number, parentID as nullable number, groupID as number) as list =>
        let
            FieldNames = Record.FieldNames(data),
            NewRows = List.Accumulate(FieldNames, {}, (accum, field) =>
                let
                    FieldValue = Record.Field(data, field),
                    Row = if Value.Is(FieldValue, type record) then
                        // Add nested record to queue for later processing
                        [Queue = List.Combine({[Queue], { [data = FieldValue, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})]
                    else if Value.Is(FieldValue, type list) then
                        // Process the list within the record
                        ProcessList(FieldValue, currentID, parentID, groupID)
                    else
                        // Simple value within the record
                        {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                in
                    List.Combine({accum, Row})
            )
        in
            NewRows,

        //--------------------------------------------------------------------------
        // Function to process lists
        ProcessList = (data as list, currentID as number, parentID as nullable number, groupID as number) as list =>
        let
            ProcessedList = List.Accumulate(data, {}, (accum, item) =>
                if Value.Is(item, type record) then
                    // Add each record in the list to the queue for later processing
                    List.Combine({accum, { [data = item, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})
                else
                    // Simple item within the list
                    List.Combine({accum, { [ID = currentID, ParentID = parentID, GroupID = groupID, Title = "ListItem", Value = item] }})
            )
        in
            ProcessedList,

        //--------------------------------------------------------------------------
        // Function to process simple values
        ProcessSimpleValue = (data as any, currentID as number, parentID as nullable number, groupID as number) as list =>
        let
            DynamicRow = [
                ID = currentID,
                ParentID = parentID,
                GroupID = groupID,
                Title = "SimpleValue",
                Value = data
            ]
        in
            {DynamicRow},

        //--------------------------------------------------------------------------
        // Validate ParentID relationships in the table
        ValidateParentID = (table as table) =>
        let
            OrphanRows = Table.SelectRows(table, each [ParentID] <> null and Table.IsEmpty(Table.SelectRows(table, each [ID] = [ParentID]))),
            ValidationResult = if Table.IsEmpty(OrphanRows) then "ParentID validation passed" else "Orphan Rows Found"
        in
            ValidationResult,

        //--------------------------------------------------------------------------
        // Process the JSON object and create the final table
        FinalTable = ProcessItem(jsonObject),

        // Perform validation
        ValidationMessage = ValidateParentID(FinalTable),

        // Return the result if validation passes
        Output = if ValidationMessage = "ParentID validation passed" then FinalTable else error ValidationMessage
    in
        Output
in
    ProcessJson