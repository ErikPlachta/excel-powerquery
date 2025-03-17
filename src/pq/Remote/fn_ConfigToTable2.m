let
    // Your JSON object here
    jsonInput = [
        title = "Power Query File Definitions",
        name = "PowerQueryDataDrivenConfig",
        settings = [
            models = [
                title = "Models",
                data = {
                    [
                        name = "SQLServer",
                        connection = "Server=myServer;Database=myDB;"
                    ],
                    [
                        name = "ExternalAPI",
                        connection = "https://api.example.com/data"
                    ]
                ]
            ]
        ]
    ],

    // Call the ProcessItem function to process the JSON input
    OutputTable = ProcessItem(jsonInput)
in
    OutputTable