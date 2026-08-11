using Query2Excel.Core.Models;

namespace Query2Excel.Core.Abstractions;

public interface IDatabaseExecutor
{
    Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken);
}

