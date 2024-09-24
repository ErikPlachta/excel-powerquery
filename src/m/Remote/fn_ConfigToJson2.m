let
    // Main function to process the JSON passed directly as input
    ProcessJson = (jsonObject as record) =>
    let
        // Function to process individual records and lists iteratively
        ProcessItem = (jsonObject as record) =>
        let
            Queue = { [data = jsonObject, parentID = null, currentID = 1, groupID = 1] },
            FinalRows = List.Generate(
                () => [Queue = Queue, Processed = {}],
                each List.Count([Queue]) > 0,
                each let
                    CurrentItem = List.First([Queue]),
                    RemainingQueue = List.RemoveFirstN([Queue], 1),
                    data = CurrentItem[data],
                    parentID = CurrentItem[parentID],
                    currentID = CurrentItem[currentID],
                    groupID = CurrentItem[groupID],
                    ProcessedRows = if data = null then
                        {}
                    else if Value.Is(data, type record) then
                        let
                            FieldNames = Record.FieldNames(data),
                            NewRows = List.Accumulate(FieldNames, {}, (accum, field) =>
                                let
                                    FieldValue = Record.Field(data, field),
                                    Row = if Value.Is(FieldValue, type record) then
                                        [Queue = List.Combine({[Queue], { [data = FieldValue, parentID = currentID, currentID = currentID + 1, groupID = groupID + 1] }})]
                                    else if Value.Is(FieldValue, type list) then
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
                                        {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = field, Value = FieldValue]}
                                in
                                    List.Combine({accum, Row})
                            )
                        in
                            NewRows
                    else if Value.Is(data, type list) then
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
                        {[ID = currentID, ParentID = parentID, GroupID = groupID, Title = "SimpleValue", Value = data}]
                in
                    [Queue = RemainingQueue, Processed = List.Combine({[Processed], ProcessedRows})]
            )[Processed],
            FinalTable = Table.FromRecords(FinalRows)
        in
            FinalTable,

        // Validate ParentID relationships in the table
        ValidateParentID = (table as table) =>
        let
            OrphanRows = Table.SelectRows(table, each [ParentID] <> null and Table.IsEmpty(Table.SelectRows(table, each [ID] = [ParentID]))),
            ValidationResult = if Table.IsEmpty(OrphanRows) then "ParentID validation passed" else "Orphan Rows Found"
        in
            ValidationResult,

        // Generate the final table
        FinalTable = ProcessItem(jsonObject),

        // Perform validation
        ValidationMessage = ValidateParentID(FinalTable),

        // Return the result if validation passes
        Output = if ValidationMessage = "ParentID validation passed" then FinalTable else error ValidationMessage
    in
        Output
in
    ProcessJson