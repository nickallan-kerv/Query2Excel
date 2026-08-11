using System.Data;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

[Collection("NonParallel-WorkingDirectory")]
public sealed class Query2ExcelService_OutputPathBehaviorTests
{
    [Fact]
    public async Task RunAsync_WithOutputFileNameOnly_WritesWorkbookInCurrentDirectory()
    {
        var workingDirectory = Path.Combine(Path.GetTempPath(), $"query2excel-cwd-{Guid.NewGuid():N}");
        Directory.CreateDirectory(workingDirectory);

        var originalDirectory = Directory.GetCurrentDirectory();
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        const string outputFileName = "output.xlsx";

        await File.WriteAllTextAsync(scriptPath, "SELECT 1 AS Id, 'ok' AS Status;");

        var resultTable = new DataTable();
        resultTable.Columns.Add("Id", typeof(int));
        resultTable.Columns.Add("Status", typeof(string));
        resultTable.Rows.Add(1, "ok");

        var fakeResult = new QueryExecutionResult(
            new[] { resultTable },
            "SELECT 1 AS Id, 'ok' AS Status;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(10));

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(fakeResult),
            new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance),
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            Directory.SetCurrentDirectory(workingDirectory);

            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputFileName,
                120);

            await service.RunAsync(request, CancellationToken.None);

            var expectedPath = Path.Combine(workingDirectory, outputFileName);
            Assert.True(File.Exists(expectedPath));
        }
        finally
        {
            Directory.SetCurrentDirectory(originalDirectory);

            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }

            if (Directory.Exists(workingDirectory))
            {
                Directory.Delete(workingDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RunAsync_WithNestedOutputPath_CreatesMissingFolderAndWritesWorkbook()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        await File.WriteAllTextAsync(scriptPath, "SELECT 1;");

        var outputRoot = Path.Combine(Path.GetTempPath(), $"query2excel-output-{Guid.NewGuid():N}");
        var outputPath = Path.Combine(outputRoot, "nested", "query2excel-console.xlsx");

        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Rows.Add(1);

        var fakeResult = new QueryExecutionResult(new[] { table }, "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(fakeResult),
            new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance),
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.True(Directory.Exists(Path.GetDirectoryName(outputPath)!));
            Assert.True(File.Exists(outputPath));
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }

            if (Directory.Exists(outputRoot))
            {
                Directory.Delete(outputRoot, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RunAsync_WhenOutputFileAlreadyExists_CreatesVersionedBackupFile()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        await File.WriteAllTextAsync(scriptPath, "SELECT 1;");

        var outputRoot = Path.Combine(Path.GetTempPath(), $"query2excel-output-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputRoot);
        var outputPath = Path.Combine(outputRoot, "query2excel-console.xlsx");
        var backupPath = outputPath + "1";

        var originalContents = "previous workbook"u8.ToArray();
        await File.WriteAllBytesAsync(outputPath, originalContents);

        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Rows.Add(1);

        var fakeResult = new QueryExecutionResult(new[] { table }, "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(fakeResult),
            new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance),
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.True(File.Exists(outputPath));
            Assert.True(File.Exists(backupPath));
            Assert.Equal(originalContents, await File.ReadAllBytesAsync(backupPath));
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }

            if (Directory.Exists(outputRoot))
            {
                Directory.Delete(outputRoot, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RunAsync_WhenVersionedBackupAlreadyExists_UsesNextAvailableBackupVersion()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.sql");
        await File.WriteAllTextAsync(scriptPath, "SELECT 1;");

        var outputRoot = Path.Combine(Path.GetTempPath(), $"query2excel-output-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputRoot);
        var outputPath = Path.Combine(outputRoot, "query2excel-console.xlsx");
        var backupPath1 = outputPath + "1";
        var backupPath2 = outputPath + "2";

        await File.WriteAllBytesAsync(outputPath, "latest previous workbook"u8.ToArray());
        await File.WriteAllBytesAsync(backupPath1, "older backup"u8.ToArray());

        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Rows.Add(1);

        var fakeResult = new QueryExecutionResult(new[] { table }, "SELECT 1;", DateTimeOffset.UtcNow, TimeSpan.FromMilliseconds(10));

        var service = new Query2ExcelService(
            new StaticResultDatabaseExecutor(fakeResult),
            new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance),
            NullLogger<Query2ExcelService>.Instance);

        try
        {
            var request = new Query2ExcelRequest(
                "Server=tcp:sample.database.windows.net,1433;Database=test;User ID=u;Password=p;Encrypt=True;",
                scriptPath,
                outputPath,
                120);

            await service.RunAsync(request, CancellationToken.None);

            Assert.True(File.Exists(outputPath));
            Assert.True(File.Exists(backupPath1));
            Assert.True(File.Exists(backupPath2));
        }
        finally
        {
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
            }

            if (Directory.Exists(outputRoot))
            {
                Directory.Delete(outputRoot, recursive: true);
            }
        }
    }
}
