using System.Data;
using Query2Excel.App.Services;

namespace Query2Excel.Tests;

public sealed class OutputWorksheetMetadataParserTests
{
    [Fact]
    public void TryParse_WithRecognizedSingleRowMetadata_ReturnsMetadata()
    {
        var table = new DataTable();
        table.Columns.Add("__SheetName", typeof(string));
        table.Columns.Add("__Title", typeof(string));
        table.Columns.Add("__Description", typeof(string));
        table.Rows.Add("Objects", "List of Objects", "All objects in scope.");

        var parsed = OutputWorksheetMetadataParser.TryParse(table, out var metadata);

        Assert.True(parsed);
        Assert.NotNull(metadata);
        Assert.Equal("Objects", metadata.SheetName);
        Assert.Equal("List of Objects", metadata.Title);
        Assert.Equal("All objects in scope.", metadata.Description);
    }

    [Fact]
    public void TryParse_WithUnrecognizedField_ReturnsFalse()
    {
        var table = new DataTable();
        table.Columns.Add("__SheetName", typeof(string));
        table.Columns.Add("__Unknown", typeof(string));
        table.Rows.Add("Objects", "X");

        var parsed = OutputWorksheetMetadataParser.TryParse(table, out var metadata);

        Assert.False(parsed);
        Assert.Null(metadata);
    }

    [Fact]
    public void TryParse_WithMultipleRows_ReturnsFalse()
    {
        var table = new DataTable();
        table.Columns.Add("__Title", typeof(string));
        table.Rows.Add("First");
        table.Rows.Add("Second");

        var parsed = OutputWorksheetMetadataParser.TryParse(table, out var metadata);

        Assert.False(parsed);
        Assert.Null(metadata);
    }
}
