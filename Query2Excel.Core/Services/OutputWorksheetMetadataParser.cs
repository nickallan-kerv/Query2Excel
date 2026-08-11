using System.Data;
using Query2Excel.Core.Models;

namespace Query2Excel.Core.Services;

public static class OutputWorksheetMetadataParser
{
    public static bool TryParse(DataTable resultSet, out OutputWorksheetMetadata? metadata)
    {
        metadata = null;

        if (resultSet.Rows.Count != 1 || resultSet.Columns.Count == 0)
        {
            return false;
        }

        foreach (DataColumn column in resultSet.Columns)
        {
            if (!OutputWorksheetMetadata.IsRecognizedField(column.ColumnName))
            {
                return false;
            }
        }

        var row = resultSet.Rows[0];
        var sheetName = GetMetadataValue(row, OutputWorksheetMetadata.SheetNameField);
        var title = GetMetadataValue(row, OutputWorksheetMetadata.TitleField);
        var description = GetMetadataValue(row, OutputWorksheetMetadata.DescriptionField);
        var appendBelowPreviousTable = GetMetadataBooleanValue(row, OutputWorksheetMetadata.AppendBelowPreviousTableField);
        var rowFormatColumn = GetMetadataValue(row, OutputWorksheetMetadata.RowFormatColumnField);

        metadata = new OutputWorksheetMetadata(sheetName, title, description, appendBelowPreviousTable, rowFormatColumn);
        return true;
    }

    private static bool GetMetadataBooleanValue(DataRow row, string fieldName)
    {
        var rawValue = GetMetadataValue(row, fieldName);
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return false;
        }

        if (bool.TryParse(rawValue, out var parsedBoolean))
        {
            return parsedBoolean;
        }

        if (int.TryParse(rawValue, out var parsedInt))
        {
            return parsedInt != 0;
        }

        var normalized = rawValue.Trim();
        if (string.Equals(normalized, "yes", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "y", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "on", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(normalized, "no", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "n", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "off", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return false;
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
}

