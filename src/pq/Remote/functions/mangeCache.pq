let
    manageCache = (filePath as text, remoteContentFunction as function, useLocalCache as logical) as text =>
    let
        // Step 1: Check if the file exists locally and if local cache is enabled
        localFileExists = try File.Contents(filePath) otherwise null,

        cachedFile = if useLocalCache and localFileExists <> null then
            Text.FromBinary(localFileExists)
        else
            let
                // Step 2: Fetch the file from the remote source using the provided function
                remoteFileContent = try remoteContentFunction()
                    otherwise error "Error fetching remote file.",

                // Step 3: Save the fetched file to the local cache
                saveToLocalCache = try Binary.SaveAs(remoteFileContent, filePath)
                    otherwise error "Error saving file to local cache."
            in
                Text.FromBinary(remoteFileContent)
    in
        cachedFile
in
    manageCache