using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using Query2Excel.Core.Abstractions;
using Query2Excel.Core.Models;

namespace Query2Excel.Core.Services;

public sealed class DatabaseExecutor(ILogger<DatabaseExecutor> logger) : IDatabaseExecutor
{
    public async Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.ConnectionString);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.SqlScript);

        var startedAt = DateTimeOffset.UtcNow;
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            logger.LogInformation("Opening database connection.");

            await using var connection = new SqlConnection(request.ConnectionString);
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

            logger.LogInformation("Executing SQL script.");

            await using var command = new SqlCommand(request.SqlScript, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = request.CommandTimeoutSeconds
            };

            using var adapter = new SqlDataAdapter(command);
            var dataSet = new DataSet();
            adapter.Fill(dataSet);

            var resultSets = dataSet.Tables
                .Cast<DataTable>()
                .ToList();

            stopwatch.Stop();

            var totalRows = resultSets.Sum(set => set.Rows.Count);

            logger.LogInformation(
                "SQL script execution completed. ResultSets: {ResultSetCount}. Rows returned: {RowCount}. DurationMs: {DurationMs}.",
                resultSets.Count,
                totalRows,
                stopwatch.Elapsed.TotalMilliseconds);

            return new QueryExecutionResult(resultSets, request.SqlScript, startedAt, stopwatch.Elapsed);
        }
        catch (SqlException sqlException)
        {
            logger.LogError(sqlException, "SQL script execution failed.");
            throw new InvalidOperationException(
                "Database execution failed. Verify the connection string and SQL script.",
                sqlException);
        }
        catch (Exception exception)
        {
            logger.LogError(exception, "Unexpected error occurred while executing SQL.");
            throw;
        }
    }
}

