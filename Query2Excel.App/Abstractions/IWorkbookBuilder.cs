using Query2Excel.App.Models;

namespace Query2Excel.App.Abstractions;

public interface IWorkbookBuilder
{
    Task BuildWorkbookAsync(QueryExecutionResult result, string outputFilePath, CancellationToken cancellationToken);
}
