let
    findAndReplaceParameters = (fileContent as text, parameters as list) as text =>
    let
        replacedContent = List.Accumulate(parameters, fileContent, (currentContent, param) =>
            if Record.HasFields(param, {"placeholder", "value"}) then
                Text.Replace(currentContent, param[placeholder], param[value])
            else
                error "Invalid parameter structure."
        )
    in
        replacedContent
in
    findAndReplaceParameters