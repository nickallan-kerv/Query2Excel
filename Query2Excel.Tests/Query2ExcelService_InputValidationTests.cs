using System.Data;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class Query2ExcelService_InputValidationTests
{
    [Fact]
    public async Task RunAsync_WithMissingSqlScript_ThrowsArgumentException()
    {
        var missingPath = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.sql");
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(new QueryExecutionResult(Array.Empty<DataTable>(), "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(1))),
            new TrackingWorkbookBuilder(),
            NullLogger<Query2ExcelService>.Instance);

        var request = new Query2ExcelRequest(
            "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
            missingPath,
            outputPath,
            120);

        var exception = await Assert.ThrowsAsync<ArgumentException>(() => service.RunAsync(request, CancellationToken.None));
        Assert.Contains("SQL script file was not found", exception.Message);
    }

    [Fact]
    public async Task RunAsync_WithEmptySqlScript_ThrowsArgumentException()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"empty-{Guid.NewGuid():N}.sql");
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        await File.WriteAllTextAsync(scriptPath, string.Empty);

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(new QueryExecutionResult(Array.Empty<DataTable>(), string.Empty, DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(1))),
            new TrackingWorkbookBuilder(),
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            var exception = await Assert.ThrowsAsync<ArgumentException>(() => service.RunAsync(request, CancellationToken.None));
            Assert.Contains("SQL script file is empty", exception.Message);
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
