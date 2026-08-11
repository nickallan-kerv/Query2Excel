namespace Query2Excel.Tests;

public sealed class ConsolePathValidationTests
{
    [Fact]
    public void ValidateFileExtensions_WithValidExtensions_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
            ConsolePathValidation.ValidateFileExtensions(
                @"C:\Temp\query.sql",
                @"C:\Temp\output.xlsx"));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateFileExtensions_WithInvalidSqlExtension_ThrowsArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ConsolePathValidation.ValidateFileExtensions(
                @"C:\Temp\query.txt",
                @"C:\Temp\output.xlsx"));

        Assert.Equal("SQL script path must point to a .sql file.", exception.Message);
    }

    [Fact]
    public void ValidateFileExtensions_WithInvalidOutputExtension_ThrowsArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ConsolePathValidation.ValidateFileExtensions(
                @"C:\Temp\query.sql",
                @"C:\Temp\output.csv"));

        Assert.Equal("Output file path must point to a .xlsx file.", exception.Message);
    }
}
