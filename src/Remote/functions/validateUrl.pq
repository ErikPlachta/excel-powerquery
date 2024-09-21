let
    validateUrl = (url as text) as logical =>
    let
        isValid = if Text.StartsWith(url, "http://") or Text.StartsWith(url, "https://") then true else error "Invalid URL: " & url
    in
        isValid
in
    validateUrl