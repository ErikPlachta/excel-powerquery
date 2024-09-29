let
    // Fetch the remote JSON file (replace with actual URL)
    Source = Json.Document(Web.Contents("https://your-json-url-here.com")),

    // Call the external JSON schema validation function (assume it exists as ValidateJson)
    // Wrap in a try...otherwise to handle invalid JSON structure gracefully
    IsJsonValid = try ValidateJson(Source) otherwise false,

    // If the JSON is invalid, return an error
    ErrorCheck = if IsJsonValid = false then
        error "Invalid JSON structure based on the schema. Please check the source."
    else
        Source,

    // Extract the parameters list and function metadata from the validated JSON
    ParametersList = ErrorCheck[Parameters],
    FunctionMeta = ErrorCheck[FunctionMetadata],

    // Map JSON "Type" to Power Query types
    MapParamType = (paramType as text) =>
        if paramType = "choice" then "text"
        else if paramType = "checkbox" or paramType = "logical" then "logical"
        else if paramType = "date" then "date"
        else if paramType = "number" then "number"
        else if paramType = "integer" then "int64"
        else if paramType = "text" then "text"
        else error "Unsupported parameter type",

    // Helper function to generate parameter meta dynamically from JSON with error handling
    GenerateParameterMeta = (param as record) =>
        let
            Title = param[Title],
            Type = MapParamType(param[Type]),
            
            // Use try...otherwise for optional fields and handle missing cases gracefully
            FieldCaption = try param[FieldCaption] otherwise "No Caption",
            FieldDescription = try param[FieldDescription] otherwise "No Description",
            SampleValues = try param[SampleValues] otherwise null,
            AllowedValues = try param[AllowedValues] otherwise null,
            IsRequired = param[Required],
            DefaultValue = try param[Default] otherwise null,

            // Validate that DefaultValue conforms to the expected Type
            ValidatedDefault = if Type = "number" and Type.Is(DefaultValue, type number) then DefaultValue
                               else if Type = "text" and Type.Is(DefaultValue, type text) then DefaultValue
                               else if Type = "logical" and Type.Is(DefaultValue, type logical) then DefaultValue
                               else if Type = "date" and Type.Is(DefaultValue, type date) then DefaultValue
                               else if Type = "int64" and Type.Is(DefaultValue, type number) then DefaultValue
                               else if DefaultValue = null then DefaultValue
                               else error Text.Format("Invalid Default value for parameter: #(#)", {Title}),
            
            // Build parameter metadata
            ParamMeta = [
                Name = Title,
                Type = Type meta [
                    Documentation.FieldCaption = FieldCaption,
                    Documentation.FieldDescription = FieldDescription,
                    Documentation.SampleValues = SampleValues,
                    Documentation.AllowedValues = AllowedValues
                ],
                DefaultValue = ValidatedDefault
            ]
        in
            ParamMeta,

    // Generate parameter meta list from the validated JSON
    ParametersMeta = List.Transform(ParametersList, each GenerateParameterMeta(_)),

    // Core function logic - apply parameters to some meaningful data processing
    DynamicFunctionImpl = (paramValues as record) =>
    let
        // This is where you'd process data based on the parameters
        // For now, we simulate data filtering logic
        AppliedLogic = List.Accumulate(ParametersMeta, "", (state, paramMeta) =>
            let
                paramName = paramMeta[Name],
                paramValue = try Record.Field(paramValues, paramName) otherwise null,
                paramCaption = Value.Metadata(paramMeta[Type])[Documentation.FieldCaption],
                
                // Process logic: for example, if paramValue is null, warn user
                newState = if paramValue = null then
                    Text.Combine({state, "Warning: ", paramCaption, " not provided; "})
                else
                    Text.Combine({state, paramCaption, ": ", Text.From(paramValue), "; "})
            in
                newState
        )
    in
        AppliedLogic,

    // Define the final function with dynamic parameter definitions
    DynamicReportFunction = (paramValues as record) =>
        DynamicFunctionImpl(paramValues),

    // Replace the function type dynamically, adding metadata from the JSON
    DynamicReportFunctionWithMeta = Value.ReplaceType(
        DynamicReportFunction,
        type function (
            // Generate parameter types dynamically from JSON metadata
            List.Accumulate(ParametersMeta, {}, (state, paramMeta) => 
                let
                    paramName = paramMeta[Name],
                    paramType = paramMeta[Type]
                in
                    Record.AddField(state, paramName, paramType)
            )
        )
        as text meta [
            Documentation.Name = FunctionMeta[Documentation.Name],
            Documentation.LongDescription = FunctionMeta[Documentation.LongDescription],
            Documentation.Examples = FunctionMeta[Documentation.Examples]
        ]
    )
in
    // Return the function with documentation
    DynamicReportFunctionWithMeta