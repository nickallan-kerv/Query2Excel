using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using Query2Excel.App.Abstractions;
using Query2Excel.App.Models;
using System.Data;

namespace Query2Excel.App.Services;

public sealed class WorkbookBuilder(ILogger<WorkbookBuilder> logger) : IWorkbookBuilder
{
    private const string SheetNameMetadataField = "__SheetName";
    private const string TitleMetadataField = "__Title";
    private const string DescriptionMetadataField = "__Description";

    public Task BuildWorkbookAsync(QueryExecutionResult result, string outputFilePath, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFilePath);
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            logger.LogInformation("Building workbook at path {OutputFilePath}.", outputFilePath);

            var directory = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var workbook = new XLWorkbook();

            BuildOutputWorksheets(workbook, result.ResultSets);
            BuildSqlWorksheet(workbook, result);

            workbook.SaveAs(outputFilePath);

            logger.LogInformation("Workbook successfully created at path {OutputFilePath}.", outputFilePath);
            return Task.CompletedTask;
        }
        catch (Exception exception)
        {
            logger.LogError(exception, "Failed to create workbook at path {OutputFilePath}.", outputFilePath);
            throw new InvalidOperationException("Workbook creation failed.", exception);
        }
    }

    private static void BuildOutputWorksheets(XLWorkbook workbook, IReadOnlyList<DataTable> resultSets)
    {
        if (resultSets.Count == 0)
        {
            BuildSingleOutputWorksheet(workbook, new DataTable(), "Output1", 1, null, null);
            return;
        }

        var outputIndex = 1;
        WorksheetMetadata? pendingMetadata = null;

        foreach (var resultSet in resultSets)
        {
            if (TryReadMetadata(resultSet, out var metadata))
            {
                pendingMetadata = metadata;
                continue;
            }

            var defaultName = $"Output{outputIndex}";
            var requestedName = string.IsNullOrWhiteSpace(pendingMetadata?.SheetName) ? defaultName : pendingMetadata.Value.SheetName;
            var worksheetName = ResolveWorksheetName(workbook, requestedName, defaultName);

            BuildSingleOutputWorksheet(workbook, resultSet, worksheetName, outputIndex, pendingMetadata?.Title, pendingMetadata?.Description);

            pendingMetadata = null;
            outputIndex++;
        }

        if (outputIndex == 1)
        {
            var worksheetName = ResolveWorksheetName(workbook, pendingMetadata?.SheetName, "Output1");
            BuildSingleOutputWorksheet(workbook, new DataTable(), worksheetName, 1, pendingMetadata?.Title, pendingMetadata?.Description);
        }
    }

    private static void BuildSingleOutputWorksheet(XLWorkbook workbook, DataTable resultSet, string worksheetName, int outputIndex, string? title, string? description)
    {
        var worksheet = workbook.Worksheets.Add(worksheetName);
        var totalColumns = resultSet.Columns.Count;
        var hasTitle = !string.IsNullOrWhiteSpace(title);
        var hasDescription = !string.IsNullOrWhiteSpace(description);
        var tableHeaderRow = 1 + (hasTitle ? 1 : 0) + (hasDescription ? 1 : 0);
        var tableDataStartRow = tableHeaderRow + 1;

        if (hasTitle)
        {
            worksheet.Cell(1, 1).Value = title;
            ApplyTitleStyle(worksheet.Cell(1, 1));
        }

        if (hasDescription)
        {
            var descriptionRow = hasTitle ? 2 : 1;
            worksheet.Cell(descriptionRow, 1).Value = description;
            ApplyDescriptionStyle(worksheet.Cell(descriptionRow, 1));
        }

        if (totalColumns == 0)
        {
            var infoCell = worksheet.Cell(tableHeaderRow, 1);
            infoCell.Value = "No tabular result set was returned by the SQL query.";
            infoCell.Style.Font.Bold = true;
            worksheet.ColumnsUsed().AdjustToContents();
            return;
        }

        for (var columnIndex = 0; columnIndex < totalColumns; columnIndex++)
        {
            worksheet.Cell(tableHeaderRow, columnIndex + 1).Value = resultSet.Columns[columnIndex].ColumnName;
        }

        if (resultSet.Rows.Count > 0)
        {
            for (var rowIndex = 0; rowIndex < resultSet.Rows.Count; rowIndex++)
            {
                for (var columnIndex = 0; columnIndex < totalColumns; columnIndex++)
                {
                    var value = resultSet.Rows[rowIndex][columnIndex];
                    SetCellValue(worksheet.Cell(tableDataStartRow + rowIndex, columnIndex + 1), value);
                }
            }

            var tableRange = worksheet.Range(tableHeaderRow, 1, tableHeaderRow + resultSet.Rows.Count, totalColumns);
            var table = tableRange.CreateTable($"QueryResults{outputIndex}");
            table.Theme = XLTableTheme.TableStyleMedium9;
            table.ShowAutoFilter = true;

            worksheet.Columns(1, totalColumns).AdjustToContents(tableHeaderRow, tableHeaderRow + resultSet.Rows.Count);
        }
        else
        {
            var tableRange = worksheet.Range(tableHeaderRow, 1, tableHeaderRow, totalColumns);
            var table = tableRange.CreateTable($"QueryResults{outputIndex}");
            table.Theme = XLTableTheme.TableStyleMedium9;
            table.ShowAutoFilter = true;

            var infoRow = tableDataStartRow + 1;
            worksheet.Cell(infoRow, 1).Value = "No rows were returned by the query.";
            worksheet.Cell(infoRow, 1).Style.Font.Italic = true;

            if (totalColumns > 1)
            {
                worksheet.Range(infoRow, 1, infoRow, totalColumns).Merge();
            }

            worksheet.Columns(1, totalColumns).AdjustToContents(tableHeaderRow, tableHeaderRow);
        }

        worksheet.SheetView.FreezeRows(tableHeaderRow);
    }

    private static bool TryReadMetadata(DataTable resultSet, out WorksheetMetadata metadata)
    {
        metadata = default;

        if (resultSet.Rows.Count != 1 || resultSet.Columns.Count == 0)
        {
            return false;
        }

        foreach (DataColumn column in resultSet.Columns)
        {
            if (!IsRecognizedMetadataField(column.ColumnName))
            {
                return false;
            }
        }

        var row = resultSet.Rows[0];
        var sheetName = GetMetadataValue(row, SheetNameMetadataField);
        var title = GetMetadataValue(row, TitleMetadataField);
        var description = GetMetadataValue(row, DescriptionMetadataField);

        metadata = new WorksheetMetadata(sheetName, title, description);
        return true;
    }

    private static bool IsRecognizedMetadataField(string columnName)
    {
        return string.Equals(columnName, SheetNameMetadataField, StringComparison.OrdinalIgnoreCase)
            || string.Equals(columnName, TitleMetadataField, StringComparison.OrdinalIgnoreCase)
            || string.Equals(columnName, DescriptionMetadataField, StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetMetadataValue(DataRow row, string fieldName)
    {
        foreach (DataColumn column in row.Table.Columns)
        {
            if (!string.Equals(column.ColumnName, fieldName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return row[column] == DBNull.Value ? null : row[column]?.ToString();
        }

        return null;
    }

    private static string ResolveWorksheetName(XLWorkbook workbook, string? requestedName, string fallbackName)
    {
        var baseName = string.IsNullOrWhiteSpace(requestedName)
            ? fallbackName
            : SanitizeWorksheetName(requestedName);

        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = fallbackName;
        }

        var uniqueName = baseName;
        var suffix = 1;

        while (workbook.Worksheets.Any(worksheet => string.Equals(worksheet.Name, uniqueName, StringComparison.OrdinalIgnoreCase)))
        {
            var suffixText = $"_{suffix}";
            var maxBaseLength = 31 - suffixText.Length;
            var trimmedBase = baseName.Length > maxBaseLength ? baseName[..maxBaseLength] : baseName;
            uniqueName = $"{trimmedBase}{suffixText}";
            suffix++;
        }

        return uniqueName;
    }

    private static string SanitizeWorksheetName(string rawName)
    {
        var trimmed = rawName.Trim().Trim('\'');
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        var invalidChars = new[] { '[', ']', ':', '*', '?', '/', '\\' };
        foreach (var invalidChar in invalidChars)
        {
            trimmed = trimmed.Replace(invalidChar, '_');
        }

        return trimmed.Length > 31 ? trimmed[..31] : trimmed;
    }

    private static void ApplyTitleStyle(IXLCell cell)
    {
        cell.Style.Font.Bold = true;
        cell.Style.Font.FontSize = 16;
        cell.Style.Font.FontName = "Calibri";
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void ApplyDescriptionStyle(IXLCell cell)
    {
        cell.Style.Font.Italic = true;
        cell.Style.Font.FontSize = 11;
        cell.Style.Font.FontName = "Calibri";
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private readonly record struct WorksheetMetadata(string? SheetName, string? Title, string? Description);

    private static void SetCellValue(IXLCell cell, object value)
    {
        if (value is DBNull)
        {
            cell.SetValue(string.Empty);
            return;
        }

        switch (value)
        {
            case string stringValue:
                cell.SetValue(stringValue);
                break;
            case bool boolValue:
                cell.SetValue(boolValue);
                break;
            case byte byteValue:
                cell.SetValue(byteValue);
                break;
            case short shortValue:
                cell.SetValue(shortValue);
                break;
            case int intValue:
                cell.SetValue(intValue);
                break;
            case long longValue:
                cell.SetValue(longValue);
                break;
            case float floatValue:
                cell.SetValue(floatValue);
                break;
            case double doubleValue:
                cell.SetValue(doubleValue);
                break;
            case decimal decimalValue:
                cell.SetValue(decimalValue);
                break;
            case DateTime dateTimeValue:
                cell.SetValue(dateTimeValue);
                break;
            case DateTimeOffset dateTimeOffsetValue:
                cell.SetValue(dateTimeOffsetValue.DateTime);
                break;
            default:
                cell.SetValue(value.ToString() ?? string.Empty);
                break;
        }
    }

    private static void BuildSqlWorksheet(XLWorkbook workbook, QueryExecutionResult result)
    {
        var worksheet = workbook.Worksheets.Add("SQL");

        worksheet.Cell("A1").Value = "Query2Excel Execution Details";
        worksheet.Cell("A1").Style.Font.Bold = true;
        worksheet.Cell("A1").Style.Font.FontSize = 14;

        worksheet.Cell("A3").Value = "Execution Timestamp (UTC)";
        worksheet.Cell("B3").Value = result.ExecutedAtUtc.UtcDateTime;
        worksheet.Cell("B3").Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";

        worksheet.Cell("A4").Value = "Execution Duration (ms)";
        worksheet.Cell("B4").Value = Math.Round(result.Duration.TotalMilliseconds, 2);

        worksheet.Cell("A5").Value = "Rows Returned";
        worksheet.Cell("B5").Value = result.TotalRowCount;

        worksheet.Cell("A6").Value = "Result Sets Returned";
        worksheet.Cell("B6").Value = result.ResultSetCount;

        worksheet.Cell("A8").Value = "Executed SQL";
        worksheet.Cell("A8").Style.Font.Bold = true;

        worksheet.Range("A9:F9").Merge();
        worksheet.Cell("A9").Value = result.ExecutedSql;
        worksheet.Cell("A9").Style.Font.FontName = "Consolas";
        worksheet.Cell("A9").Style.Font.FontSize = 11;
        worksheet.Cell("A9").Style.Alignment.WrapText = true;
        worksheet.Cell("A9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        worksheet.Row(9).Height = 200;

        worksheet.Column("A").Width = 28;
        worksheet.Column("B").Width = 22;
        worksheet.Columns("C:F").Width = 16;
    }
}
