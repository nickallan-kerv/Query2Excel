using Query2Excel.App.Models;

namespace Query2Excel.App.Abstractions;

public interface IDatabaseExecutor
{
    Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken);
}
