let
    // Sample data for two tables
    Table1Data = Table.FromRecords({
        [Column1="A1", Column2="B1"],
        [Column1="A2", Column2="B2"]
    }),
    Table2Data = Table.FromRecords({
        [Column1="C1", Column2="D1"],
        [Column1="C2", Column2="D2"]
    }),

    // Create or update two tables dynamically
    ResultTable1 = CreateOrUpdateTable("SalesTable", Table1Data),
    ResultTable2 = CreateOrUpdateTable("FinanceTable", Table2Data)
in
    [ResultTable1, ResultTable2]