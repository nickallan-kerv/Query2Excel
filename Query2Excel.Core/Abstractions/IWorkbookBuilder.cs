using Query2Excel.Core.Models;

namespace Query2Excel.Core.Abstractions;

public interface IWorkbookBuilder
{
    Task BuildWorkbookAsync(QueryExecutionResult result, string outputFilePath, CancellationToken cancellationToken);
}

