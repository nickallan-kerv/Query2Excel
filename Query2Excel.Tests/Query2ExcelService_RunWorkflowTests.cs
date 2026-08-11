using System.Data;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class Query2ExcelService_RunWorkflowTests
{
    [Fact]
    public async Task RunAsync_ExecutesQueryAndBuildsWorkbook()
    {
        var expectedPath = @"C:\temp\query2excel.xlsx";
        var fakeResult = new QueryExecutionResult(new[] { new DataTable() }, "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));
        var databaseExecutor = new StaticResultDatabaseExecutor(fakeResult);
        var workbookBuilder = new TrackingWorkbookBuilder();
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

    [Fact]
    public async Task RunAsync_WithPortableQuery_WritesWorkbookWithSqlAndOutputSheets()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-workflow-{Guid.NewGuid():N}.xlsx");
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");

        var resultTable = new DataTable();
        resultTable.Columns.Add("Id", typeof(int));
        resultTable.Columns.Add("Status", typeof(string));
        resultTable.Rows.Add(1, "ok");

        var fakeResult = new QueryExecutionResult(
            new[] { resultTable },
            "SELECT 1 AS Id, 'ok' AS Status;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(25));

        var databaseExecutor = new StaticResultDatabaseExecutor(fakeResult);
        var workbookBuilder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);
        var service = new Query2ExcelService(databaseExecutor, workbookBuilder, NullLogger<Query2ExcelService>.Instance);

        await File.WriteAllTextAsync(scriptPath, "SELECT 1 AS Id, 'ok' AS Status;");

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.True(File.Exists(outputPath));

            using var workbook = new XLWorkbook(outputPath);
            Assert.NotNull(workbook.Worksheet("Output1"));
            Assert.NotNull(workbook.Worksheet("SQL"));
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }
}
