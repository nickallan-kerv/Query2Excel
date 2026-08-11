using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class Query2ExcelService_ErrorHandlingTests
{
    [Fact]
    public async Task RunAsync_WhenDatabaseExecutorThrows_DoesNotInvokeWorkbookBuilder()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        await File.WriteAllTextAsync(scriptPath, "SELECT broken_sql");

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var workbookBuilder = new TrackingWorkbookBuilder();
        var expectedException = new InvalidOperationException("Incorrect syntax near 'broken_sql'.");

        var service = new Query2ExcelService(
            new ThrowingDatabaseExecutor(expectedException),
            workbookBuilder,
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => service.RunAsync(request, CancellationToken.None));
            Assert.Equal(expectedException.Message, exception.Message);
            Assert.False(workbookBuilder.BuildCalled);
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }
        }
    }
}
