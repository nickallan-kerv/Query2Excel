namespace Query2Excel.Core.Configuration;

public sealed class Query2ExcelOptions
{
    public const string SectionName = "Query2Excel";
    public const string DefaultSqlScriptPath = "Query2Excel.App\\Scripts\\Example.sql";

    public string? ConnectionString { get; init; }

    public string? SqlScript { get; init; } = DefaultSqlScriptPath;

    public string? OutputFilePath { get; init; }

    public int CommandTimeoutSeconds { get; init; } = 120;

    public Dictionary<string, RowStyleOptions> RowStyles { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class RowStyleOptions
{
    public string? BackgroundColor { get; init; }

    public string? ForegroundColor { get; init; }

    public bool Bold { get; init; }

    public double? FontSize { get; init; }

    public string? NumberFormat { get; init; }

    public bool NumberFormatNumericOnly { get; init; } = true;
}
