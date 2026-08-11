namespace Query2Excel.Core.Models;

public sealed record Query2ExcelRequest(
    string ConnectionString,
    string SqlScriptPath,
    string OutputFilePath,
    int CommandTimeoutSeconds,
    string? ConnectionStringTemplate = null
);

