let
    processExtractedFiles = (extractedFiles as table, subActions as list) as table =>
    let
        result = executeSubActions(subActions)
    in
        if result = null then error "Sub-actions failed." else result
in
    processExtractedFiles