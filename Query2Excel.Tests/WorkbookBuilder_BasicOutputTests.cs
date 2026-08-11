using System.Data;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Query2Excel.Core.Configuration;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class WorkbookBuilder_BasicOutputTests
{
    [Fact]
    public async Task BuildWorkbookAsync_WithMultipleResultSets_CreatesOutputWorksheetsPerResultSet()
    {
        var firstResultSet = CreateResultTable();
        firstResultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var secondResultSet = CreateResultTable();
        secondResultSet.Rows.Add(2, "Beta", new DateTime(2026, 8, 10, 10, 30, 0, DateTimeKind.Utc), false);

        var result = new QueryExecutionResult(
            new[] { firstResultSet, secondResultSet },
            "EXEC dbo.ProcA;\nEXEC dbo.ProcB;",
            new DateTimeOffset(2026, 8, 10, 12, 0, 0, TimeSpan.Zero),
            TimeSpan.FromMilliseconds(248));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);

            Assert.NotNull(workbook.Worksheet("Output1"));
            Assert.NotNull(workbook.Worksheet("Output2"));
            Assert.NotNull(workbook.Worksheet("SQL"));

            var output1Sheet = workbook.Worksheet("Output1");
            Assert.Equal("Id", output1Sheet.Cell(1, 1).GetString());
            Assert.Equal("Name", output1Sheet.Cell(1, 2).GetString());
            Assert.Equal("Alpha", output1Sheet.Cell(2, 2).GetString());
            Assert.Single(output1Sheet.Tables);

            var output2Sheet = workbook.Worksheet("Output2");
            Assert.Equal("Beta", output2Sheet.Cell(2, 2).GetString());
            Assert.Single(output2Sheet.Tables);

            var sqlSheet = workbook.Worksheet("SQL");
            Assert.Equal("Connection String Template", sqlSheet.Cell("A7").GetString());
            Assert.True(sqlSheet.Cell("B7").IsMerged());
            Assert.Equal(result.ExecutedSql, sqlSheet.Cell("A10").GetString());
            Assert.Equal(result.TotalRowCount, sqlSheet.Cell("B5").GetValue<int>());
            Assert.Equal(result.ResultSetCount, sqlSheet.Cell("B6").GetValue<int>());
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
    public async Task BuildWorkbookAsync_WithNoRows_WritesHeadersAndInformationalMessage()
    {
        var dataTable = CreateResultTable();
        var result = new QueryExecutionResult(
            new[] { dataTable },
            "SELECT Id, Name, CreatedAtUtc, IsActive FROM dbo.Sample WHERE 1 = 0;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(80));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("Id", outputSheet.Cell(1, 1).GetString());
            Assert.Equal("Name", outputSheet.Cell(1, 2).GetString());
            Assert.Equal("No rows were returned by the query.", outputSheet.Cell(3, 1).GetString());
            Assert.Single(outputSheet.Tables);
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
    public async Task BuildWorkbookAsync_WithRowFormatMetadata_AppliesRowStyleAndOmitsMetadataColumn()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add(1, "Alpha", "Good");
        dataTable.Rows.Add(2, "Beta", "Neutral");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 'Alpha' AS Name, 'Good' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(110));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("Id", outputSheet.Cell(1, 1).GetString());
            Assert.Equal("Name", outputSheet.Cell(1, 2).GetString());
            Assert.NotEqual("__RowFormat", outputSheet.Cell(1, 3).GetString());

            Assert.Equal(XLColor.FromHtml("#C6EFCE"), outputSheet.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#006100"), outputSheet.Cell(2, 1).Style.Font.FontColor);

            Assert.Equal(XLColor.FromHtml("#FFEB9C"), outputSheet.Cell(3, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#9C5700"), outputSheet.Cell(3, 1).Style.Font.FontColor);
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
    public async Task BuildWorkbookAsync_WithMissingRowFormatColumn_ThrowsInvalidOperationException()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Rows.Add(1, "Alpha");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 'Alpha' AS Name;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(110));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            var exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None));

            Assert.Equal("Workbook creation failed.", exception.Message);
            Assert.NotNull(exception.InnerException);
            Assert.Equal(
                "Metadata validation failed: __RowFormatColumn '__RowFormat' was not found in the following data result set.",
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

    [Fact]
    public async Task BuildWorkbookAsync_WithCheckCellRowFormat_AppliesExpectedRowStyle()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add(1, "Alpha", "Check Cell");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 'Alpha' AS Name, 'Check Cell' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(95));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal(XLColor.FromHtml("#E2EFDA"), outputSheet.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#375623"), outputSheet.Cell(2, 1).Style.Font.FontColor);
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
    public async Task BuildWorkbookAsync_WithAccent2AndHeading1RowFormats_AppliesExpectedStyles()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add(1, "Alpha", "Accent2");
        dataTable.Rows.Add(2, "Beta", "Heading 1");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 'Alpha' AS Name, 'Accent2' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(95));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal(XLColor.FromHtml("#FBE2D5"), outputSheet.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#000000"), outputSheet.Cell(2, 1).Style.Font.FontColor);

            Assert.Equal(XLColor.FromHtml("#1F4E78"), outputSheet.Cell(3, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.White, outputSheet.Cell(3, 1).Style.Font.FontColor);
            Assert.True(outputSheet.Cell(3, 1).Style.Font.Bold);
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
    public async Task BuildWorkbookAsync_WithCurrencyRowFormat_AppliesCurrencyNumberFormatToNumericCells()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Amount", typeof(decimal));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add(1, 123.45m, "Currency");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 123.45 AS Amount, 'Currency' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(95));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal("£#,##0.00", outputSheet.Cell(2, 1).Style.NumberFormat.Format);
            Assert.Equal("£#,##0.00", outputSheet.Cell(2, 2).Style.NumberFormat.Format);
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
    public async Task BuildWorkbookAsync_WithRawCredentialsInConnectionStringTemplate_MasksCredentialsInSqlWorksheet()
    {
        var rawUser = $"user-{Guid.NewGuid():N}";
        var rawPassword = $"pass-{Guid.NewGuid():N}";

        var resultSet = CreateResultTable();
        resultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var result = new QueryExecutionResult(
            new[] { resultSet },
            "SELECT 1;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(50),
            $"Server=localhost;Database=Sandbox;User ID={rawUser};Password={rawPassword};TrustServerCertificate=True;");

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var templateValue = workbook.Worksheet("SQL").Cell("B7").GetString();

            Assert.Contains("User ID={UserId}", templateValue, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Password={Password}", templateValue, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(rawUser, templateValue, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(rawPassword, templateValue, StringComparison.Ordinal);
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
    public async Task BuildWorkbookAsync_WithTokenizedConnectionStringTemplate_PreservesTokensInSqlWorksheet()
    {
        var resultSet = CreateResultTable();
        resultSet.Rows.Add(1, "Alpha", new DateTime(2026, 8, 10, 9, 0, 0, DateTimeKind.Utc), true);

        var result = new QueryExecutionResult(
            new[] { resultSet },
            "SELECT 1;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(50),
            "Server=localhost;Database=Sandbox;User ID={UserId};Password={Password};TrustServerCertificate=True;");

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithConfiguredRowStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var templateValue = workbook.Worksheet("SQL").Cell("B7").GetString();

            Assert.Contains("User ID={UserId}", templateValue, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Password={Password}", templateValue, StringComparison.OrdinalIgnoreCase);
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

    private static WorkbookBuilder CreateBuilderWithConfiguredRowStyles()
    {
        var options = Options.Create(new Query2ExcelOptions
        {
            RowStyles = new Dictionary<string, RowStyleOptions>(StringComparer.OrdinalIgnoreCase)
            {
                ["Good"] = new() { BackgroundColor = "#C6EFCE", ForegroundColor = "#006100" },
                ["Bad"] = new() { BackgroundColor = "#FFC7CE", ForegroundColor = "#9C0006" },
                ["Neutral"] = new() { BackgroundColor = "#FFEB9C", ForegroundColor = "#9C5700" },
                ["Check Cell"] = new() { BackgroundColor = "#E2EFDA", ForegroundColor = "#375623" },
                ["Accent2"] = new() { BackgroundColor = "#FBE2D5", ForegroundColor = "#000000" },
                ["Heading 1"] = new() { BackgroundColor = "#1F4E78", ForegroundColor = "#FFFFFF", Bold = true, FontSize = 12 },
                ["Currency"] = new() { NumberFormat = "£#,##0.00", NumberFormatNumericOnly = true }
            }
        });

        return new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance, options);
    }
}
