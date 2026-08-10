namespace Query2Excel.App.Models;

public sealed record QueryExecutionRequest(
    string ConnectionString,
    string SqlScript,
    int CommandTimeoutSeconds
);
