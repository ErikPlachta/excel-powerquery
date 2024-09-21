let
    recordSuccessOrFail = (result as text, stepName as text, logPath as text) as record =>
    let
        recordData = if result = "Success" then
            [ Step = stepName, Status = "Success", Timestamp = DateTime.LocalNow() ]
        else
            logError("Step failed.", stepName, logPath)
    in
        recordData
in
    recordSuccessOrFail