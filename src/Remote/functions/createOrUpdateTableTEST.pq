let
    // Table 1 data
    Table1Data = Table.FromRecords({
        [Column1="A1", Column2="B1"],
        [Column1="A2", Column2="B2"]
    }),

    // Table 2 data
    Table2Data = Table.FromRecords({
        [Column1="C1", Column2="D1"],
        [Column1="C2", Column2="D2"]
    }),

    // Create or update two tables dynamically
    ResultTable1 = CreateOrUpdateTable("SalesTable", Table1Data, "SalesSheet"),
    ResultTable2 = CreateOrUpdateTable("FinanceTable", Table2Data, "FinanceSheet")
in
    [ResultTable1, ResultTable2]