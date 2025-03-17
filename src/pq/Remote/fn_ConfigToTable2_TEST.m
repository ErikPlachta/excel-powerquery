let
    // Define your JSON structure directly in Power Query
    jsonInput = 
    [
        title = "Power Query File Definitions",
        name = "PowerQueryDataDrivenConfig",
        description = "This configuration defines the data processing logic for Power Query files with settings, parameters, and actions defined remotely.",
        settings = [
            title = "Settings",
            description = "Configuration settings for connections, defaults, and caching",
            data = [
                models = [
                    title = "Models",
                    description = "Data connection models used in the configuration",
                    data = {
                        [
                            name = "SQLServer",
                            description = "Connection string to the SQL Server",
                            connection = "Server=myServer;Database=myDB;Trusted_Connection=True;",
                            sourceType = "sql",
                            dataType = "csv"
                        ],
                        [
                            name = "PowerBIModel",
                            description = "Connection string to the PowerBI data source",
                            connection = "Data Source=myPowerBI;Catalog=myCatalog",
                            sourceType = "dax",
                            dataType = "json"
                        ]
                    }
                ]
            ]
        ]
    ],
    
    // Call the function and pass the JSON input
    OutputTable = fn_JsonToTable(jsonInput)
in
    OutputTable