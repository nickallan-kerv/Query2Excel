using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Query2Excel.Core.Abstractions;
using Query2Excel.Core.Configuration;
using Query2Excel.Core.Models;
using System.Globalization;
using System.Data;

namespace Query2Excel.Core.Services;

public sealed class WorkbookBuilder : IWorkbookBuilder
{
    private readonly ILogger<WorkbookBuilder> logger;
    private readonly IReadOnlyDictionary<string, RowStyleDefinition> rowStyleDefinitions;

    public WorkbookBuilder(ILogger<WorkbookBuilder> logger)
        : this(logger, Options.Create(new Query2ExcelOptions()))
    {
    }

    public WorkbookBuilder(ILogger<WorkbookBuilder> logger, IOptions<Query2ExcelOptions> options)
    {
        this.logger = logger;
        rowStyleDefinitions = BuildRowStyleDefinitions(options.Value.RowStyles);
    }

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

            BuildOutputWorksheets(workbook, result.ResultSets, rowStyleDefinitions);
            BuildSqlWorksheet(workbook, result);

            CreateVersionedBackupIfFileExists(outputFilePath);
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

    private static void BuildOutputWorksheets(
        XLWorkbook workbook,
        IReadOnlyList<DataTable> resultSets,
        IReadOnlyDictionary<string, RowStyleDefinition> rowStyleDefinitions)
    {
        if (resultSets.Count == 0)
        {
            var worksheet = workbook.Worksheets.Add("Output1");
            BuildTableSection(worksheet, new DataTable(), 1, 1, null, null, null, true, rowStyleDefinitions);
            return;
        }

        var outputIndex = 1;
        OutputWorksheetMetadata? pendingMetadata = null;
        IXLWorksheet? lastOutputWorksheet = null;
        var nextAppendStartRow = 1;

        foreach (var resultSet in resultSets)
        {
            if (OutputWorksheetMetadataParser.TryParse(resultSet, out var metadata))
            {
                ValidateMetadata(metadata!);
                pendingMetadata = metadata;
                continue;
            }

            var appendBelowPreviousTable = pendingMetadata?.AppendBelowPreviousTable == true && lastOutputWorksheet is not null;
            if (appendBelowPreviousTable)
            {
                nextAppendStartRow = BuildTableSection(
                    lastOutputWorksheet!,
                    resultSet,
                    nextAppendStartRow,
                    outputIndex,
                    pendingMetadata?.Title,
                    pendingMetadata?.Description,
                    pendingMetadata?.RowFormatColumn,
                    freezeRows: false,
                    rowStyleDefinitions: rowStyleDefinitions);
            }
            else
            {
                var defaultName = $"Output{outputIndex}";
                var requestedName = string.IsNullOrWhiteSpace(pendingMetadata?.SheetName) ? defaultName : pendingMetadata.SheetName;
                var worksheetName = ResolveWorksheetName(workbook, requestedName, defaultName);

                var worksheet = workbook.Worksheets.Add(worksheetName);
                nextAppendStartRow = BuildTableSection(
                    worksheet,
                    resultSet,
                    1,
                    outputIndex,
                    pendingMetadata?.Title,
                    pendingMetadata?.Description,
                    pendingMetadata?.RowFormatColumn,
                    freezeRows: true,
                    rowStyleDefinitions: rowStyleDefinitions);

                lastOutputWorksheet = worksheet;
            }

            pendingMetadata = null;
            outputIndex++;
        }

        if (outputIndex == 1)
        {
            var worksheetName = ResolveWorksheetName(workbook, pendingMetadata?.SheetName, "Output1");
            var worksheet = workbook.Worksheets.Add(worksheetName);
            BuildTableSection(
                worksheet,
                new DataTable(),
                1,
                1,
                pendingMetadata?.Title,
                pendingMetadata?.Description,
                pendingMetadata?.RowFormatColumn,
                true,
                rowStyleDefinitions);
        }
    }

    private static void ValidateMetadata(OutputWorksheetMetadata metadata)
    {
        if (!string.IsNullOrWhiteSpace(metadata.SheetName) && metadata.AppendBelowPreviousTable)
        {
            throw new InvalidOperationException("Metadata validation failed: __SheetName cannot be combined with __AppendBelowPreviousTable.");
        }
    }

    private static int BuildTableSection(
        IXLWorksheet worksheet,
        DataTable resultSet,
        int startRow,
        int outputIndex,
        string? title,
        string? description,
        string? rowFormatColumn,
        bool freezeRows,
        IReadOnlyDictionary<string, RowStyleDefinition> rowStyleDefinitions)
    {
        var rowFormatColumnIndex = ResolveRowFormatColumnIndex(resultSet, rowFormatColumn);
        var visibleColumnIndices = Enumerable.Range(0, resultSet.Columns.Count)
            .Where(index => index != rowFormatColumnIndex)
            .ToArray();

        var totalColumns = visibleColumnIndices.Length;
        var hasTitle = !string.IsNullOrWhiteSpace(title);
        var hasDescription = !string.IsNullOrWhiteSpace(description);
        var tableHeaderRow = startRow + (hasTitle ? 1 : 0) + (hasDescription ? 1 : 0);
        var tableDataStartRow = tableHeaderRow + 1;

        if (hasTitle)
        {
            var titleRow = startRow;
            worksheet.Cell(titleRow, 1).Value = title;
            ApplyTitleStyle(worksheet.Cell(titleRow, 1));
        }

        if (hasDescription)
        {
            var descriptionRow = startRow + (hasTitle ? 1 : 0);
            worksheet.Cell(descriptionRow, 1).Value = description;
            ApplyDescriptionStyle(worksheet.Cell(descriptionRow, 1));
        }

        if (totalColumns == 0)
        {
            var infoCell = worksheet.Cell(tableHeaderRow, 1);
            infoCell.Value = "No tabular result set was returned by the SQL query.";
            infoCell.Style.Font.Bold = true;
            worksheet.ColumnsUsed().AdjustToContents();

            if (freezeRows)
            {
                worksheet.SheetView.FreezeRows(tableHeaderRow);
            }

            return tableHeaderRow + 1;
        }

        for (var visibleColumnPosition = 0; visibleColumnPosition < totalColumns; visibleColumnPosition++)
        {
            var sourceColumnIndex = visibleColumnIndices[visibleColumnPosition];
            worksheet.Cell(tableHeaderRow, visibleColumnPosition + 1).Value = resultSet.Columns[sourceColumnIndex].ColumnName;
        }

        if (resultSet.Rows.Count > 0)
        {
            for (var rowIndex = 0; rowIndex < resultSet.Rows.Count; rowIndex++)
            {
                for (var visibleColumnPosition = 0; visibleColumnPosition < totalColumns; visibleColumnPosition++)
                {
                    var sourceColumnIndex = visibleColumnIndices[visibleColumnPosition];
                    var value = resultSet.Rows[rowIndex][sourceColumnIndex];
                    SetCellValue(worksheet.Cell(tableDataStartRow + rowIndex, visibleColumnPosition + 1), value);
                }
            }

            var tableRange = worksheet.Range(tableHeaderRow, 1, tableHeaderRow + resultSet.Rows.Count, totalColumns);
            var table = tableRange.CreateTable($"QueryResults{outputIndex}");
            table.Theme = XLTableTheme.TableStyleMedium9;
            table.ShowAutoFilter = true;

            ApplyRowStyles(worksheet, resultSet, rowFormatColumnIndex, tableDataStartRow, totalColumns, rowStyleDefinitions);

            worksheet.Columns(1, totalColumns).AdjustToContents(tableHeaderRow, tableHeaderRow + resultSet.Rows.Count);

            if (freezeRows)
            {
                worksheet.SheetView.FreezeRows(tableHeaderRow);
            }

            return tableHeaderRow + resultSet.Rows.Count + 1;
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

            if (freezeRows)
            {
                worksheet.SheetView.FreezeRows(tableHeaderRow);
            }

            return infoRow + 1;
        }
    }

    private static int ResolveRowFormatColumnIndex(DataTable resultSet, string? rowFormatColumn)
    {
        if (string.IsNullOrWhiteSpace(rowFormatColumn))
        {
            return -1;
        }

        var columnName = rowFormatColumn.Trim();
        for (var index = 0; index < resultSet.Columns.Count; index++)
        {
            if (string.Equals(resultSet.Columns[index].ColumnName, columnName, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        throw new InvalidOperationException(
            $"Metadata validation failed: __RowFormatColumn '{columnName}' was not found in the following data result set.");
    }

    private static void ApplyRowStyles(
        IXLWorksheet worksheet,
        DataTable resultSet,
        int rowFormatColumnIndex,
        int tableDataStartRow,
        int visibleColumnCount,
        IReadOnlyDictionary<string, RowStyleDefinition> rowStyleDefinitions)
    {
        if (rowFormatColumnIndex < 0 || visibleColumnCount <= 0)
        {
            return;
        }

        for (var rowIndex = 0; rowIndex < resultSet.Rows.Count; rowIndex++)
        {
            var rawStyle = resultSet.Rows[rowIndex][rowFormatColumnIndex];
            var styleName = rawStyle == DBNull.Value ? string.Empty : rawStyle?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(styleName)
                || string.Equals(styleName, "Normal", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rowRange = worksheet.Range(
                tableDataStartRow + rowIndex,
                1,
                tableDataStartRow + rowIndex,
                visibleColumnCount);

            ApplyNamedRowStyle(rowRange, styleName, rowStyleDefinitions);
        }
    }

    private static void ApplyNamedRowStyle(
        IXLRange rowRange,
        string styleName,
        IReadOnlyDictionary<string, RowStyleDefinition> rowStyleDefinitions)
    {
        var styleKey = NormalizeStyleKey(styleName);
        if (string.IsNullOrWhiteSpace(styleKey)
            || !rowStyleDefinitions.TryGetValue(styleKey, out var definition))
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(definition.BackgroundColor))
        {
            rowRange.Style.Fill.BackgroundColor = XLColor.FromHtml(definition.BackgroundColor);
        }

        if (!string.IsNullOrWhiteSpace(definition.ForegroundColor))
        {
            rowRange.Style.Font.FontColor = XLColor.FromHtml(definition.ForegroundColor);
        }

        if (definition.Bold)
        {
            rowRange.Style.Font.Bold = true;
        }

        if (definition.FontSize.HasValue)
        {
            rowRange.Style.Font.FontSize = definition.FontSize.Value;
        }

        if (string.IsNullOrWhiteSpace(definition.NumberFormat))
        {
            return;
        }

        if (definition.NumberFormatNumericOnly)
        {
            foreach (var cell in rowRange.Cells())
            {
                if (cell.DataType == XLDataType.Number)
                {
                    cell.Style.NumberFormat.Format = definition.NumberFormat;
                }
            }

            return;
        }

        rowRange.Style.NumberFormat.Format = definition.NumberFormat;
    }

    private static IReadOnlyDictionary<string, RowStyleDefinition> BuildRowStyleDefinitions(
        IReadOnlyDictionary<string, RowStyleOptions>? configuredStyles)
    {
        var definitions = new Dictionary<string, RowStyleDefinition>(StringComparer.OrdinalIgnoreCase);

        if (configuredStyles is null)
        {
            return definitions;
        }

        foreach (var configuredStyle in configuredStyles)
        {
            var styleKey = NormalizeStyleKey(configuredStyle.Key);
            if (string.IsNullOrWhiteSpace(styleKey))
            {
                continue;
            }

            definitions[styleKey] = RowStyleDefinition.From(configuredStyle.Value);
        }

        return definitions;
    }

    private static string NormalizeStyleKey(string styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
        {
            return string.Empty;
        }

        var compact = new string(styleName
            .Trim()
            .Where(character => !char.IsWhiteSpace(character) && character != '-' && character != '_')
            .ToArray());

        return compact.ToLowerInvariant();
    }

    private sealed record RowStyleDefinition(
        string? BackgroundColor,
        string? ForegroundColor,
        bool Bold,
        double? FontSize,
        string? NumberFormat,
        bool NumberFormatNumericOnly)
    {
        public static RowStyleDefinition From(RowStyleOptions options)
        {
            return new RowStyleDefinition(
                options.BackgroundColor,
                options.ForegroundColor,
                options.Bold,
                options.FontSize,
                options.NumberFormat,
                options.NumberFormatNumericOnly);
        }
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

    private static void CreateVersionedBackupIfFileExists(string outputFilePath)
    {
        if (!File.Exists(outputFilePath))
        {
            return;
        }

        var backupPath = GetNextVersionedBackupPath(outputFilePath);
        File.Move(outputFilePath, backupPath);
    }

    private static string GetNextVersionedBackupPath(string filePath)
    {
        for (var version = 1; version < int.MaxValue; version++)
        {
            var candidate = filePath + version.ToString(CultureInfo.InvariantCulture);
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException($"Unable to create a versioned backup for '{filePath}'.");
    }

    private static void BuildSqlWorksheet(XLWorkbook workbook, QueryExecutionResult result)
    {
        var worksheet = workbook.Worksheets.Add("SQL");
        var safeConnectionStringTemplate = ConnectionStringTemplateProtector.SanitizeForWorkbook(result.ConnectionStringTemplate);

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

        worksheet.Cell("A7").Value = "Connection String Template";
        worksheet.Range("B7:F7").Merge();
        worksheet.Cell("B7").Value = safeConnectionStringTemplate;
        worksheet.Cell("B7").Style.Alignment.WrapText = true;
        worksheet.Cell("B7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

        worksheet.Cell("A9").Value = "Executed SQL";
        worksheet.Cell("A9").Style.Font.Bold = true;

        worksheet.Range("A10:F10").Merge();
        worksheet.Cell("A10").Value = result.ExecutedSql;
        worksheet.Cell("A10").Style.Font.FontName = "Consolas";
        worksheet.Cell("A10").Style.Font.FontSize = 11;
        worksheet.Cell("A10").Style.Alignment.WrapText = true;
        worksheet.Cell("A10").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        worksheet.Row(10).Height = 200;

        worksheet.Column("A").Width = 28;
        worksheet.Column("B").Width = 22;
        worksheet.Columns("C:F").Width = 16;
    }
}

