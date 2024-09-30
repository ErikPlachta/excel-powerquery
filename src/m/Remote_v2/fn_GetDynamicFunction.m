(jsonUrl as text) =>
let
    // Fetch the remote JSON file (replace with actual URL)
    Source = Json.Document(Web.Contents(jsonUrl))

    // Extract the parameters list and function metadata from the validated JSON
    ,ParametersList = Source[Parameters]
    ,FunctionMeta = Source[FunctionMetadata]

    // Map JSON "Type" to Power Query types
    ,MapParamType = (paramType as text) =>
        if paramType = "choice" then type text
        else if paramType = "checkbox" or paramType = "logical" then type logical
        else if paramType = "date" then type date
        else if paramType = "number" then type number
        else if paramType = "integer" then type number
        else if paramType = "text" then type text
        else error "Unsupported parameter type"

    // Helper function to generate parameter meta dynamically from JSON with error handling
    ,GenerateParameterMeta = (param as record) =>
        let
            Title = param[Title]
            ,TypeType = param[Type] // The text value
            ,Type = MapParamType(param[Type]) // The actual Type
            // Use try...otherwise for optional fields and handle missing cases gracefully
            ,FieldCaption = try param[FieldCaption] otherwise "No Caption"
            ,FieldDescription = try param[FieldDescription] otherwise "No Description"
            ,SampleValues = try param[SampleValues] otherwise null
            ,AllowedValues = try param[AllowedValues] otherwise null
            ,IsRequired = param[Required]
            ,DefaultValue = try param[Default] otherwise null
            // Build parameter metadata
            ,ParamMeta = [
                Name = Title
                ,TypeType = TypeType
                ,Type = Type
                ,DefaultValue = DefaultValue
                ,Meta = Type meta [
                    Documentation.FieldCaption = FieldCaption
                    ,Documentation.FieldDescription = FieldDescription
                    ,Documentation.SampleValues = SampleValues
                    ,Documentation.AllowedValues = AllowedValues
                    ,Documentation.IsRequired = IsRequired
                ]
            ]
        in
            ParamMeta

    // Generate parameter meta list from the validated JSON
    ,ParametersMeta = List.Transform(ParametersList, each GenerateParameterMeta(_))

    // Core function logic - apply parameters to some meaningful data processing
    ,DynamicFunctionImpl = (paramValues as record) as text =>
        let
            ConcatenatedParams = 
                List.Accumulate(ParametersMeta, "", (state, paramMeta) =>
                    let
                        paramName = paramMeta[Name]
                        ,paramValue = try Record.Field(paramValues, paramName) otherwise ""
                        ,paramCaption = Value.Metadata(paramMeta[Type])[Documentation.FieldCaption]
                    in
                        state & paramCaption & ": " & Text.From(paramValue) & "; "
                )
        in
            ConcatenatedParams

    // Build the function type dynamically using Type.ForFunction
    ,ParameterNames = List.Transform(ParametersMeta, each _[Name])
    ,ParameterTypes = List.Transform(ParametersMeta, each _[Type])

    // Create a record of parameter names and types
    ,ParameterSignature = Record.FromList(ParameterTypes, ParameterNames)

    // Create the function type dynamically
    ,DynamicFunctionType = Type.ForFunction(ParameterSignature, type text)

    // Define the final function with dynamic parameter definitions
    ,DynamicReportFunction = (paramValues as record) => DynamicFunctionImpl(paramValues)

    // Replace the function type dynamically, adding metadata from the JSON
    ,DynamicReportFunctionWithMeta = Value.ReplaceType(
        DynamicReportFunction
        ,DynamicFunctionType meta [
            Documentation.Name = FunctionMeta[Documentation.Name]
            ,Documentation.LongDescription = FunctionMeta[Documentation.LongDescription]
            ,Documentation.Examples = FunctionMeta[Documentation.Examples]
        ]
    )
in
    // Return the function with documentation
    DynamicReportFunctionWithMeta