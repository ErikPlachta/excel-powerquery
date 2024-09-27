let
    // Sample target table
    TargetTable = Table.FromRecords({
        [ID = 1, Name = "John", Department = "HR"],
        [ID = 2, Name = "Jane", Department = "Finance"],
        [ID = 3, Name = "Bob", Department = "HR"]
    }),
    
    // Array of column name and value pairs to create new columns
    ConditionsArray = {
        [ColumnName = "Department", Value = "HR", NewColumn = "IsHR"],
        [ColumnName = "Name", Value = "Jane", NewColumn = "IsJane"]
    },

    // Function to loop through the array and add columns based on conditions
    AddConditionalColumns = (table as table, conditions as list) =>
        List.Accumulate(
            conditions,
            table,
            (state, current) =>
                Table.AddColumn(
                    state,
                    current[NewColumn],
                    each if Record.Field(_, current[ColumnName]) = current[Value] then true else false
                )
        ),

    // Call the function to add the columns to the target table
    ResultTable = AddConditionalColumns(TargetTable, ConditionsArray)

in
    ResultTable