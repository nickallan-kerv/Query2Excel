namespace Query2Excel.App.Models;

public sealed record OutputWorksheetMetadata(string? SheetName, string? Title, string? Description)
{
    public const string SheetNameField = "__SheetName";
    public const string TitleField = "__Title";
    public const string DescriptionField = "__Description";

    private static readonly HashSet<string> RecognizedFields = new(StringComparer.OrdinalIgnoreCase)
    {
        SheetNameField,
        TitleField,
        DescriptionField
    };

    public static bool IsRecognizedField(string fieldName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);
        return RecognizedFields.Contains(fieldName);
    }
}
