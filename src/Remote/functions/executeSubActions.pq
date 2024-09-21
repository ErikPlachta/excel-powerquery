let
    executeSubActions = (subActions as list) as any =>
    let
        result = List.Accumulate(subActions, null, (state, action) =>
            if state = null or state[Status] = "Success" then
                try action() otherwise error "Sub-action failed."
            else 
                error "Previous action failed, skipping sub-action."
        )
    in
        result
in
    executeSubActions