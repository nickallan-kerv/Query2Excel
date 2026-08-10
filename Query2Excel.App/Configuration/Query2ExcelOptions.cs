namespace Query2Excel.App.Configuration;

public sealed class Query2ExcelOptions
{
    public const string SectionName = "Query2Excel";
    public const string DefaultSqlScriptPath = "Query2Excel.App\\Scripts\\Example.sql";

    public string? ConnectionString { get; init; }

    public string? SqlScript { get; init; } = DefaultSqlScriptPath;

    public string? OutputFilePath { get; init; }

    public int CommandTimeoutSeconds { get; init; } = 120;
}
