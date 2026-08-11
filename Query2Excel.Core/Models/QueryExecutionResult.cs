using System.Data;

namespace Query2Excel.Core.Models;

public sealed record QueryExecutionResult(
    IReadOnlyList<DataTable> ResultSets,
    string ExecutedSql,
    DateTimeOffset ExecutedAtUtc,
    TimeSpan Duration,
    string? ConnectionStringTemplate = null
)
{
    public int ResultSetCount => ResultSets.Count;

    public int TotalRowCount => ResultSets.Sum(table => table.Rows.Count);
}

