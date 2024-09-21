let
    validateFileType = (fileType as text, expectedType as text) as logical =>
    let
        isValid = if fileType = expectedType then true else error "Invalid file type. Expected: " & expectedType & ", Found: " & fileType
    in
        isValid
in
    validateFileType