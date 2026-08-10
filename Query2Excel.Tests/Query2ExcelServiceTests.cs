using System.Data;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.App.Abstractions;
using Query2Excel.App.Models;
using Query2Excel.App.Services;

namespace Query2Excel.Tests;

public sealed class Query2ExcelServiceTests
{
    [Fact]
    public async Task RunAsync_ExecutesQueryAndBuildsWorkbook()
    {
        var expectedPath = @"C:\\temp\\query2excel.xlsx";
        var fakeResult = new QueryExecutionResult(new[] { new DataTable() }, "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));
        var databaseExecutor = new FakeDatabaseExecutor(fakeResult);
        var workbookBuilder = new FakeWorkbookBuilder();
        var service = new Query2ExcelService(databaseExecutor, workbookBuilder, NullLogger<Query2ExcelService>.Instance);

        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        await File.WriteAllTextAsync(scriptPath, "SELECT 1;");

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                expectedPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.True(databaseExecutor.ExecuteCalled);
            Assert.True(workbookBuilder.BuildCalled);
            Assert.Equal(expectedPath, workbookBuilder.OutputFilePath);
            Assert.Equal("SELECT 1;", databaseExecutor.LastRequest?.SqlScript);
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }
        }
    }

    private sealed class FakeDatabaseExecutor(QueryExecutionResult result) : IDatabaseExecutor
    {
        public bool ExecuteCalled { get; private set; }

        public QueryExecutionRequest? LastRequest { get; private set; }

        public Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken)
        {
            ExecuteCalled = true;
            LastRequest = request;
            return Task.FromResult(result);
        }
    }

    private sealed class FakeWorkbookBuilder : IWorkbookBuilder
    {
        public bool BuildCalled { get; private set; }

        public string? OutputFilePath { get; private set; }

        public Task BuildWorkbookAsync(QueryExecutionResult result, string outputFilePath, CancellationToken cancellationToken)
        {
            BuildCalled = true;
            OutputFilePath = outputFilePath;
            return Task.CompletedTask;
        }
    }
}
