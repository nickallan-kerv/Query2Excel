using System.Data;
using Query2Excel.App.Models;

namespace Query2Excel.App.Services;

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

        metadata = new OutputWorksheetMetadata(sheetName, title, description);
        return true;
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
