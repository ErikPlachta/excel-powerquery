let
    // The remote JSON (replace with actual URL)
    Source = Json.Document(Web.Contents("https://your-json-url-here.com")),

    // Validation Function: Check if all required fields are present in an object
    ValidateRequiredFields = (record as record, requiredFields as list) as logical =>
        List.Accumulate(requiredFields, true, (state, field) => 
            state and Record.HasFields(record, field)
        ),

    // Validation Function: Check if the field types match expected types
    ValidateFieldTypes = (record as record, fieldSchemas as list) as logical =>
        List.Accumulate(fieldSchemas, true, (state, fieldSchema) => 
            let
                fieldName = fieldSchema[FieldName],
                fieldType = fieldSchema[FieldType],
                fieldValue = try Record.Field(record, fieldName) otherwise null,
                isValidType = if fieldType = "string" then Type.Is(fieldValue, type text)
                              else if fieldType = "array" then Type.Is(fieldValue, type list)
                              else if fieldType = "boolean" then Type.Is(fieldValue, type logical)
                              else if fieldType = "object" then Type.Is(fieldValue, type record)
                              else if fieldType = "null" then fieldValue = null
                              else true
            in
                state and isValidType
        ),

    // Function to validate the entire JSON against the schema
    ValidateJson = (json as record) as logical =>
        let
            // Define required fields at the top level
            requiredFieldsTopLevel = {"Parameters", "FunctionMetadata"},
            topLevelFieldsValid = ValidateRequiredFields(json, requiredFieldsTopLevel),

            // Validate Parameters array
            parameters = json[Parameters],
            requiredParameterFields = {"Title", "Type", "Required"},
            parameterSchemas = {
                [FieldName="Title", FieldType="string"],
                [FieldName="Description", FieldType="string"],
                [FieldName="Type", FieldType="string"],
                [FieldName="Default", FieldType="any"],
                [FieldName="Options", FieldType="array"],
                [FieldName="Required", FieldType="boolean"],
                [FieldName="FieldCaption", FieldType="string"],
                [FieldName="FieldDescription", FieldType="string"],
                [FieldName="SampleValues", FieldType="array"],
                [FieldName="AllowedValues", FieldType="array"],
                [FieldName="Examples", FieldType="array"]
            },
            parametersValid = List.Accumulate(parameters, true, (state, param) => 
                state and ValidateRequiredFields(param, requiredParameterFields) 
                and ValidateFieldTypes(param, parameterSchemas)
            ),

            // Validate FunctionMetadata object
            functionMetadata = json[FunctionMetadata],
            requiredFunctionFields = {"Documentation.Name", "Documentation.LongDescription", "Documentation.Examples"},
            functionSchemas = {
                [FieldName="Documentation.Name", FieldType="string"],
                [FieldName="Documentation.LongDescription", FieldType="string"],
                [FieldName="Documentation.Examples", FieldType="array"]
            },
            functionMetadataValid = ValidateRequiredFields(functionMetadata, requiredFunctionFields) 
                                    and ValidateFieldTypes(functionMetadata, functionSchemas)
        in
            topLevelFieldsValid and parametersValid and functionMetadataValid
in
    // Check if the JSON is valid
    ValidateJson(Source)