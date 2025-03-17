let
    validateFilePath = (filePath as text) as logical =>
    let
        isValid = if Text.Length(filePath) > 0 and Text.Contains(filePath, "/") or Text.Contains(filePath, "\\") then true else error "Invalid file path: " & filePath
    in
        isValid
in
    validateFilePath