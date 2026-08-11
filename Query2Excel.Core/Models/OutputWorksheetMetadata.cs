namespace Query2Excel.Core.Models;

public sealed record OutputWorksheetMetadata(
    string? SheetName,
    string? Title,
    string? Description,
    bool AppendBelowPreviousTable,
    string? RowFormatColumn)
{
    public const string SheetNameField = "__SheetName";
    public const string TitleField = "__Title";
    public const string DescriptionField = "__Description";
    public const string AppendBelowPreviousTableField = "__AppendBelowPreviousTable";
    public const string RowFormatColumnField = "__RowFormatColumn";

    private static readonly HashSet<string> RecognizedFields = new(StringComparer.OrdinalIgnoreCase)
    {
        SheetNameField,
        TitleField,
        DescriptionField,
        AppendBelowPreviousTableField,
        RowFormatColumnField
    };

    public static bool IsRecognizedField(string fieldName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);
        return RecognizedFields.Contains(fieldName);
    }
}

