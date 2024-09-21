let
    saveToMetadataTable = (metadata as record, content as text, logPath as text) as table =>
    let
        validate = if Record.HasFields(metadata, {"Title", "Description", "Type", "Path"}) then true else error "Invalid metadata structure",
        newRow = Table.FromRecords({[
            Title = metadata[Title],
            Description = metadata[Description],
            Type = metadata[Type],
            Path = metadata[Path],
            ModifiedDateTime = metadata[ModifiedDateTime],
            CreatedDateTime = metadata[CreatedDateTime],
            AdditionalMetadata = Text.FromBinary(metadata[AdditionalMetadata]),
            Content = content
        ]}),
        MetadataTable = try Excel.CurrentWorkbook(){[Name="FileMetadata"]}[Content]
            otherwise Table.FromRecords({}),
        UpdatedMetadataTable = Table.Combine({MetadataTable, newRow})
    in
        if UpdatedMetadataTable = null then logError("Failed to save metadata.", "saveToMetadataTable", logPath) else UpdatedMetadataTable
in
    saveToMetadataTable