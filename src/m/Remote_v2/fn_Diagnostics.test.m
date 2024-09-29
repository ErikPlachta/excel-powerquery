let
    // Example query function with multiple steps
    ExampleQueryFunction = (params as record) as table =>
        let
            Source = Table.FromRows({{1, "A"}, {2, "B"}}, {"ID", "Value"}),
            Filtered = Table.SelectRows(Source, each [ID] = params[ID]),
            Transformed = Table.TransformColumns(Filtered, {"Value", Text.Lower})
        in
            Transformed,

    // Call the dynamic diagnostic function, specifying steps to trace and values to log
    DiagnosticQuery = DynamicDiagnosticFunction(
        ExampleQueryFunction, 
        TraceLevel.Information, 
        {"Source", "Filtered"}, 
        {"Transformed"}
    ),

    // Pass parameter values to the diagnostic query
    Result = DiagnosticQuery([ID = 1])
in
    Result