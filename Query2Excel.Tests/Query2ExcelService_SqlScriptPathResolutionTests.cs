using System.Data;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

[Collection("NonParallel-WorkingDirectory")]
public sealed class Query2ExcelService_SqlScriptPathResolutionTests
{
    [Fact]
    public async Task RunAsync_WithRelativeProjectPrefixedScriptPath_LoadsScriptContent()
    {
        var currentDirectory = Directory.GetCurrentDirectory();
        var scriptsDirectory = Path.Combine(currentDirectory, "Query2Excel.App", "Scripts");
        Directory.CreateDirectory(scriptsDirectory);

        var scriptFileName = $"smoke-{Guid.NewGuid():N}.sql";
        var absoluteScriptPath = Path.Combine(scriptsDirectory, scriptFileName);
        const string sql = "SELECT 2 AS Id, 'prefixed' AS Status;";
        await File.WriteAllTextAsync(absoluteScriptPath, sql);

        var fakeResult = new QueryExecutionResult(new[] { new DataTable() }, sql, DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));
        var databaseExecutor = new StaticResultDatabaseExecutor(fakeResult);
        var workbookBuilder = new TrackingWorkbookBuilder();
        var service = new Query2ExcelService(databaseExecutor, workbookBuilder, NullLogger<Query2ExcelService>.Instance);

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                Path.Combine("Query2Excel.App", "Scripts", scriptFileName),
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.Equal(sql, databaseExecutor.LastRequest?.SqlScript);
        }
        finally
        {
            if (File.Exists(absoluteScriptPath))
            {
                File.Delete(absoluteScriptPath);
            }
        }
    }

    [Fact]
    public async Task RunAsync_WithRelativeScriptsPath_LoadsScriptContent()
    {
        var currentDirectory = Directory.GetCurrentDirectory();
        var scriptsDirectory = Path.Combine(currentDirectory, "Scripts");
        Directory.CreateDirectory(scriptsDirectory);

        var scriptFileName = $"smoke-{Guid.NewGuid():N}.sql";
        var absoluteScriptPath = Path.Combine(scriptsDirectory, scriptFileName);
        const string sql = "SELECT 3 AS Id, 'relative' AS Status;";
        await File.WriteAllTextAsync(absoluteScriptPath, sql);

        var fakeResult = new QueryExecutionResult(new[] { new DataTable() }, sql, DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));
        var databaseExecutor = new StaticResultDatabaseExecutor(fakeResult);
        var workbookBuilder = new TrackingWorkbookBuilder();
        var service = new Query2ExcelService(databaseExecutor, workbookBuilder, NullLogger<Query2ExcelService>.Instance);

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                Path.Combine("Scripts", scriptFileName),
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.Equal(sql, databaseExecutor.LastRequest?.SqlScript);
        }
        finally
        {
            if (File.Exists(absoluteScriptPath))
            {
                File.Delete(absoluteScriptPath);
            }
        }
    }
}
