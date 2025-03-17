let
    fn_DataQualityCheck = (data as table, rules as record) as table =>
    let
        CheckNulls = if Record.HasFields(rules, "NoNulls") and rules[NoNulls] = true then
            Table.SelectRows(data, each List.NonNullCount(Record.ToList(_)) = Table.ColumnCount(data))
        else data,
        
        CheckDataType = if Record.HasFields(rules, "DataTypeRules") then
            List.Accumulate(Record.FieldNames(rules[DataTypeRules]), CheckNulls, (state, columnName) =>
                Table.TransformColumns(state, {{columnName, each if Type.Is(_, rules[DataTypeRules][columnName]) then _ else null}})
            )
        else CheckNulls
    in
        CheckDataType
in
    fn_DataQualityCheck