using System.Data;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class WorkbookBuilder_MetadataLayoutTests
{
    [Fact]
    public async Task BuildWorkbookAsync_WithMetadataRow_RenamesNextOutputWorksheetOnly()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__SheetName", typeof(string));
        metadataSet.Rows.Add("Objects");

        var firstResultSet = CreateResultTable();
        firstResultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var secondResultSet = CreateResultTable();
        secondResultSet.Rows.Add(2, "Beta", new DateTime(2026, 8, 10, 10, 30, 0, DateTimeKind.Utc), false);

        var result = new QueryExecutionResult(
            new[] { metadataSet, firstResultSet, secondResultSet },
            "SELECT __SheetName = 'Objects';\nEXEC sp_find;\nEXEC sp_who2;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(120));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);

            Assert.Contains(workbook.Worksheets, worksheet => worksheet.Name == "Objects");
            Assert.Contains(workbook.Worksheets, worksheet => worksheet.Name == "Output2");
            Assert.DoesNotContain(workbook.Worksheets, worksheet => worksheet.Name == "Output1");

            var objectsSheet = workbook.Worksheet("Objects");
            Assert.Equal("Alpha", objectsSheet.Cell(2, 2).GetString());

            var output2Sheet = workbook.Worksheet("Output2");
            Assert.Equal("Beta", output2Sheet.Cell(2, 2).GetString());
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
    public async Task BuildWorkbookAsync_WithSheetTitleMetadata_InsertsTitleAboveTable()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__Title", typeof(string));
        metadataSet.Rows.Add("List of Objects");

        var resultSet = CreateResultTable();
        resultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var result = new QueryExecutionResult(
            new[] { metadataSet, resultSet },
            "SELECT __Title = 'List of Objects';\nEXEC sp_find;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(120));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("List of Objects", outputSheet.Cell("A1").GetString());
            Assert.Equal("Id", outputSheet.Cell("A2").GetString());
            Assert.Equal("Alpha", outputSheet.Cell("B3").GetString());
            Assert.Equal(2, outputSheet.Tables.Single().RangeAddress.FirstAddress.RowNumber);
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
    public async Task BuildWorkbookAsync_WithSheetNameAndSheetTitleMetadata_AppliesBothToNextResultSet()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__SheetName", typeof(string));
        metadataSet.Columns.Add("__Title", typeof(string));
        metadataSet.Rows.Add("Objects", "List of Objects");

        var firstResultSet = CreateResultTable();
        firstResultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var secondResultSet = CreateResultTable();
        secondResultSet.Rows.Add(2, "Beta", new DateTime(2026, 8, 10, 10, 30, 0, DateTimeKind.Utc), false);

        var result = new QueryExecutionResult(
            new[] { metadataSet, firstResultSet, secondResultSet },
            "SELECT __SheetName = 'Objects', __Title = 'List of Objects';\nEXEC sp_find;\nEXEC sp_who2;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(150));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);

            var objectsSheet = workbook.Worksheet("Objects");
            Assert.Equal("List of Objects", objectsSheet.Cell("A1").GetString());
            Assert.Equal("Id", objectsSheet.Cell("A2").GetString());

            var output2Sheet = workbook.Worksheet("Output2");
            Assert.Equal("Beta", output2Sheet.Cell("B2").GetString());
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
    public async Task BuildWorkbookAsync_WithDescriptionOnlyMetadata_PlacesDescriptionInA1AndTableInA2()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__Description", typeof(string));
        metadataSet.Rows.Add("Object inventory for the current server.");

        var resultSet = CreateResultTable();
        resultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var result = new QueryExecutionResult(
            new[] { metadataSet, resultSet },
            "SELECT __Description = 'Object inventory for the current server.';\nEXEC sp_find;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(100));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("Object inventory for the current server.", outputSheet.Cell("A1").GetString());
            Assert.Equal("Id", outputSheet.Cell("A2").GetString());
            Assert.Equal("Alpha", outputSheet.Cell("B3").GetString());
            Assert.Equal(2, outputSheet.Tables.Single().RangeAddress.FirstAddress.RowNumber);
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
    public async Task BuildWorkbookAsync_WithTitleAndDescriptionMetadata_PlacesDescriptionAfterTitleAndPreservesAutosize()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__SheetName", typeof(string));
        metadataSet.Columns.Add("__Title", typeof(string));
        metadataSet.Columns.Add("__Description", typeof(string));
        metadataSet.Rows.Add(
            "Objects",
            "This is an intentionally very long title that should not control table column width.",
            "This is a very long description that should not expand column A beyond table content sizing.");

        var resultSet = new DataTable();
        resultSet.Columns.Add("Id", typeof(int));
        resultSet.Columns.Add("Name", typeof(string));
        resultSet.Rows.Add(1, "A");

        var result = new QueryExecutionResult(
            new[] { metadataSet, resultSet },
            "SELECT __SheetName = 'Objects', __Title = 'Long title', __Description = 'Long description';\nSELECT 1 AS Id, 'A' AS Name;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(100));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var objectsSheet = workbook.Worksheet("Objects");

            Assert.Equal("This is an intentionally very long title that should not control table column width.", objectsSheet.Cell("A1").GetString());
            Assert.Equal("This is a very long description that should not expand column A beyond table content sizing.", objectsSheet.Cell("A2").GetString());
            Assert.Equal("Id", objectsSheet.Cell("A3").GetString());
            Assert.Equal("A", objectsSheet.Cell("B4").GetString());
            Assert.Equal(3, objectsSheet.Tables.Single().RangeAddress.FirstAddress.RowNumber);

            Assert.True(objectsSheet.Column(1).Width < 20d);
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
    public async Task BuildWorkbookAsync_WithSheetNameAndAppendBelowPreviousTable_ThrowsInvalidOperationException()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__SheetName", typeof(string));
        metadataSet.Columns.Add("__AppendBelowPreviousTable", typeof(string));
        metadataSet.Rows.Add("Objects", "1");

        var resultSet = CreateResultTable();
        resultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var result = new QueryExecutionResult(
            new[] { metadataSet, resultSet },
            "SELECT __SheetName = 'Objects', __AppendBelowPreviousTable = 1;\nEXEC sp_find;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(100));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance);

        try
        {
            var exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None));

            Assert.Equal("Workbook creation failed.", exception.Message);
            Assert.NotNull(exception.InnerException);
            Assert.Equal(
                "Metadata validation failed: __SheetName cannot be combined with __AppendBelowPreviousTable.",
                exception.InnerException!.Message);
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
