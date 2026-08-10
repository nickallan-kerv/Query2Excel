using System.Data;

namespace Query2Excel.App.Models;

public sealed record QueryExecutionResult(
    IReadOnlyList<DataTable> ResultSets,
    string ExecutedSql,
    DateTimeOffset ExecutedAtUtc,
    TimeSpan Duration
)
{
    public int ResultSetCount => ResultSets.Count;

    public int TotalRowCount => ResultSets.Sum(table => table.Rows.Count);
}
