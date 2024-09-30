let
    // Fetch the remote JSON file (replace with actual URL)
    Source = Json.Document(Web.Contents("https://your-json-url-here.com")),

    // Extract the parameters list and function metadata from the JSON
    ParametersList = Source[Parameters],
    FunctionMeta = Source[FunctionMetadata],

    // Map JSON "Type" to Power Query types (ensure this matches JSON input)
    MapParamType = (paramType as text) =>
        if paramType = "choice" then type text
        else if paramType = "checkbox" or paramType = "logical" then type logical
        else if paramType = "date" then type date
        else if paramType = "number" then type number
        else if paramType = "integer" then type number
        else if paramType = "text" then type text
        else error "Unsupported parameter type",

    // Helper function to generate parameter meta dynamically from JSON
    GenerateParameterMeta = (param as record) =>
        let
            Title = param[Title],
            Type = MapParamType(param[Type]),

            // Handle optional fields
            FieldCaption = try param[FieldCaption] otherwise "No Caption",
            FieldDescription = try param[FieldDescription] otherwise "No Description",
            SampleValues = try param[SampleValues] otherwise null,
            AllowedValues = try param[AllowedValues] otherwise null,

            // Build parameter metadata
            ParamMeta = [
                Name = Title,
                Type = Type meta [
                    Documentation.FieldCaption = FieldCaption,
                    Documentation.FieldDescription = FieldDescription,
                    Documentation.SampleValues = SampleValues
                ]
            ]
        in
            ParamMeta,

    // Generate parameter meta list from the validated JSON
    ParametersMeta = List.Transform(ParametersList, each GenerateParameterMeta(_)),

    // Function implementation - concatenate parameter values for demonstration
    DynamicFunctionImpl = (paramValues as record) as text =>
    let
        ConcatenatedParams = List.Accumulate(ParametersMeta, "", (state, paramMeta) =>
            let
                paramName = paramMeta[Name],
                paramValue = try Record.Field(paramValues, paramName) otherwise "",
                paramCaption = Value.Metadata(paramMeta[Type])[Documentation.FieldCaption]
            in
                state & paramCaption & ": " & Text.From(paramValue) & "; "
        )
    in
        ConcatenatedParams,

    // Construct the dynamic function type manually
    DynamicFunctionType = type function (
        Report as (type text meta [
            Documentation.FieldCaption = "Report Selection",
            Documentation.FieldDescription = "Select the report to generate."
        ]),
        AsOfDate as (type logical meta [
            Documentation.FieldCaption = "As of Date Usage",
            Documentation.FieldDescription = "Whether to use an 'As of Date' in the report."
        ]),
        StartDate as (type date meta [
            Documentation.FieldCaption = "Report Start Date",
            Documentation.FieldDescription = "Select the start date for the report."
        ]),
        Threshold as (type number meta [
            Documentation.FieldCaption = "Filter Threshold",
            Documentation.FieldDescription = "Set a threshold value to filter the report data."
        ])
    ) as text,

    // Define the final function with dynamic parameter definitions
    DynamicReportFunction = (paramValues as record) as text => 
        DynamicFunctionImpl(paramValues),

    // Replace the function type dynamically, adding metadata from the JSON
    DynamicReportFunctionWithMeta = Value.ReplaceType(
        DynamicReportFunction,
        DynamicFunctionType meta [
            Documentation.Name = FunctionMeta[Documentation.Name],
            Documentation.LongDescription = FunctionMeta[Documentation.LongDescription],
            Documentation.Examples = FunctionMeta[Documentation.Examples]
        ]
    )
in
    // Return the function with documentation
    DynamicReportFunctionWithMeta