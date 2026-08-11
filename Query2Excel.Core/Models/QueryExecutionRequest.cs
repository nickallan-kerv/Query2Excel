namespace Query2Excel.Core.Models;

public sealed record QueryExecutionRequest(
    string ConnectionString,
    string SqlScript,
    int CommandTimeoutSeconds
);

