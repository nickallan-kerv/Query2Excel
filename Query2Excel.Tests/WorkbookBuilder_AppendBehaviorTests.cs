using System.Data;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class WorkbookBuilder_AppendBehaviorTests
{
    [Fact]
    public async Task BuildWorkbookAsync_WithAppendBelowPreviousTableMetadata_AppendsNextResultSetToSameWorksheet()
    {
        var firstResultSet = CreateResultTable();
        firstResultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var appendMetadataSet = new DataTable();
        appendMetadataSet.Columns.Add("__AppendBelowPreviousTable", typeof(bool));
        appendMetadataSet.Rows.Add(true);

        var secondResultSet = CreateResultTable();
        secondResultSet.Rows.Add(2, "Beta", new DateTime(2026, 8, 10, 10, 30, 0, DateTimeKind.Utc), false);

        var result = new QueryExecutionResult(
            new[] { firstResultSet, appendMetadataSet, secondResultSet },
            "EXEC sp_find;\nSELECT __AppendBelowPreviousTable = 1;\nEXEC sp_who2;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(120));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.DoesNotContain(workbook.Worksheets, worksheet => worksheet.Name == "Output2");
            Assert.Equal("Id", outputSheet.Cell("A1").GetString());
            Assert.Equal("Alpha", outputSheet.Cell("B2").GetString());
            Assert.Equal("Id", outputSheet.Cell("A3").GetString());
            Assert.Equal("Beta", outputSheet.Cell("B4").GetString());
            Assert.Equal(2, outputSheet.Tables.Count());
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public async Task BuildWorkbookAsync_WithAppendBelowPreviousTableMetadata_PreservesFirstTableFreezeRow()
    {
        var metadataForFirstResultSet = new DataTable();
        metadataForFirstResultSet.Columns.Add("__Title", typeof(string));
        metadataForFirstResultSet.Rows.Add("List of Objects");

        var firstResultSet = CreateResultTable();
        firstResultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var appendMetadataSet = new DataTable();
        appendMetadataSet.Columns.Add("__AppendBelowPreviousTable", typeof(string));
        appendMetadataSet.Rows.Add("true");

        var secondResultSet = CreateResultTable();
        secondResultSet.Rows.Add(2, "Beta", new DateTime(2026, 8, 10, 10, 30, 0, DateTimeKind.Utc), false);

        var result = new QueryExecutionResult(
            new[] { metadataForFirstResultSet, firstResultSet, appendMetadataSet, secondResultSet },
            "SELECT __Title = 'List of Objects';\nEXEC sp_find;\nSELECT __AppendBelowPreviousTable = 1;\nEXEC sp_who2;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(150));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("List of Objects", outputSheet.Cell("A1").GetString());
            Assert.Equal("Id", outputSheet.Cell("A2").GetString());
            Assert.Equal("Id", outputSheet.Cell("A4").GetString());
            Assert.Equal(2d, outputSheet.SheetView.SplitRow);
            Assert.Equal(2, outputSheet.Tables.Count());
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    private static DataTable CreateResultTable()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("CreatedAtUtc", typeof(DateTime));
        dataTable.Columns.Add("IsActive", typeof(bool));
        return dataTable;
    }
}
