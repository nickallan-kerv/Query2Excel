using Query2Excel.App.Models;

namespace Query2Excel.App.Abstractions;

public interface IQuery2ExcelService
{
    Task RunAsync(Query2ExcelRequest request, CancellationToken cancellationToken);
}
