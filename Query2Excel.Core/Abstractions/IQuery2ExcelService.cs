using Query2Excel.Core.Models;

namespace Query2Excel.Core.Abstractions;

public interface IQuery2ExcelService
{
    Task RunAsync(Query2ExcelRequest request, CancellationToken cancellationToken);
}

