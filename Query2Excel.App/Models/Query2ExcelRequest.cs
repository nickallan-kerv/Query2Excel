namespace Query2Excel.App.Models;

public sealed record Query2ExcelRequest(
    string ConnectionString,
    string SqlScriptPath,
    string OutputFilePath,
    int CommandTimeoutSeconds
);
