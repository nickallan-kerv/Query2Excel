using System.Data;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Query2Excel.Core.Configuration;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;

namespace Query2Excel.Tests;

public sealed class WorkbookBuilder_RowStyleSupportTests
{
    [Theory]
    [InlineData("Good", "#C6EFCE", "#006100")]
    [InlineData("Bad", "#FFC7CE", "#9C0006")]
    [InlineData("Neutral", "#FFEB9C", "#9C5700")]
    [InlineData("Check Cell", "#E2EFDA", "#375623")]
    [InlineData("Accent1", "#DCE6F1", "#000000")]
    [InlineData("Accent2", "#F2DCDB", "#000000")]
    [InlineData("Accent3", "#EBF1DE", "#000000")]
    [InlineData("Accent4", "#E4DFEC", "#000000")]
    [InlineData("Accent5", "#DAEEF3", "#000000")]
    [InlineData("Accent6", "#FDE9D9", "#000000")]
    public async Task BuildWorkbookAsync_WithSupportedRowStyle_AppliesConfiguredBackgroundAndForeground(
        string rowFormat,
        string expectedBackgroundColor,
        string expectedForegroundColor)
    {
        var result = CreateSingleRowFormatResultSet(rowFormat);
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithSupportedStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal(XLColor.FromHtml(expectedBackgroundColor), outputSheet.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml(expectedForegroundColor), outputSheet.Cell(2, 1).Style.Font.FontColor);
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
    public async Task BuildWorkbookAsync_WithHeading1RowStyle_AppliesConfiguredTypography()
    {
        var result = CreateSingleRowFormatResultSet("Heading 1");
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithSupportedStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var styledCell = workbook.Worksheet("Output1").Cell(2, 1);

            Assert.Equal(XLColor.FromHtml("#1F4E78"), styledCell.Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#FFFFFF"), styledCell.Style.Font.FontColor);
            Assert.True(styledCell.Style.Font.Bold);
            Assert.Equal(12d, styledCell.Style.Font.FontSize);
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
    public async Task BuildWorkbookAsync_WithCurrencyRowStyle_AppliesNumberFormatToNumericCellsOnly()
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Label", typeof(string));
        dataTable.Columns.Add("Amount", typeof(decimal));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add("Alpha", 123.45m, "Currency");

        var result = new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            "SELECT __RowFormatColumn = '__RowFormat';\nSELECT 'Alpha' AS Label, 123.45 AS Amount, 'Currency' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(90));

        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithSupportedStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var outputSheet = workbook.Worksheet("Output1");

            Assert.Equal(string.Empty, outputSheet.Cell(2, 1).Style.NumberFormat.Format);
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

    [Theory]
    [InlineData("check_cell")]
    [InlineData("Check-Cell")]
    [InlineData("CHECK CELL")]
    [InlineData("  check    cell  ")]
    public async Task BuildWorkbookAsync_WithNormalizedRowStyleName_MatchesConfiguredStyle(string rowFormat)
    {
        var result = CreateSingleRowFormatResultSet(rowFormat);
        var outputPath = Path.Combine(Path.GetTempPath(), $"query2excel-{Guid.NewGuid():N}.xlsx");
        var builder = CreateBuilderWithSupportedStyles();

        try
        {
            await builder.BuildWorkbookAsync(result, outputPath, CancellationToken.None);

            using var workbook = new XLWorkbook(outputPath);
            var styledCell = workbook.Worksheet("Output1").Cell(2, 1);

            Assert.Equal(XLColor.FromHtml("#E2EFDA"), styledCell.Style.Fill.BackgroundColor);
            Assert.Equal(XLColor.FromHtml("#375623"), styledCell.Style.Font.FontColor);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    private static QueryExecutionResult CreateSingleRowFormatResultSet(string rowFormat)
    {
        var metadataSet = new DataTable();
        metadataSet.Columns.Add("__RowFormatColumn", typeof(string));
        metadataSet.Rows.Add("__RowFormat");

        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("__RowFormat", typeof(string));
        dataTable.Rows.Add(1, "Alpha", rowFormat);

        return new QueryExecutionResult(
            new[] { metadataSet, dataTable },
            $"SELECT __RowFormatColumn = '__RowFormat';\nSELECT 1 AS Id, 'Alpha' AS Name, '{rowFormat}' AS __RowFormat;",
            DateTimeOffset.UtcNow,
            TimeSpan.FromMilliseconds(90));
    }

    private static WorkbookBuilder CreateBuilderWithSupportedStyles()
    {
        var options = Options.Create(new Query2ExcelOptions
        {
            RowStyles = new Dictionary<string, RowStyleOptions>(StringComparer.OrdinalIgnoreCase)
            {
                ["Good"] = new() { BackgroundColor = "#C6EFCE", ForegroundColor = "#006100" },
                ["Bad"] = new() { BackgroundColor = "#FFC7CE", ForegroundColor = "#9C0006" },
                ["Neutral"] = new() { BackgroundColor = "#FFEB9C", ForegroundColor = "#9C5700" },
                ["Check Cell"] = new() { BackgroundColor = "#E2EFDA", ForegroundColor = "#375623" },
                ["Accent1"] = new() { BackgroundColor = "#DCE6F1", ForegroundColor = "#000000" },
                ["Accent2"] = new() { BackgroundColor = "#F2DCDB", ForegroundColor = "#000000" },
                ["Accent3"] = new() { BackgroundColor = "#EBF1DE", ForegroundColor = "#000000" },
                ["Accent4"] = new() { BackgroundColor = "#E4DFEC", ForegroundColor = "#000000" },
                ["Accent5"] = new() { BackgroundColor = "#DAEEF3", ForegroundColor = "#000000" },
                ["Accent6"] = new() { BackgroundColor = "#FDE9D9", ForegroundColor = "#000000" },
                ["Heading 1"] = new() { BackgroundColor = "#1F4E78", ForegroundColor = "#FFFFFF", Bold = true, FontSize = 12 },
                ["Currency"] = new() { NumberFormat = "£#,##0.00", NumberFormatNumericOnly = true }
            }
        });

        return new WorkbookBuilder(NullLogger<WorkbookBuilder>.Instance, options);
    }
}
