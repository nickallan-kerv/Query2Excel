using System.Text;

namespace Query2Excel.Tests;

public sealed class ProgramEntry_HelpTests
{
    [Fact]
    public async Task RunAsync_WithHelpArgument_PrintsHelpAndReturnsZero()
    {
        var originalOut = Console.Out;
        var outputBuffer = new StringBuilder();
        await using var writer = new StringWriter(outputBuffer);

        try
        {
            Console.SetOut(writer);

            var exitCode = await ProgramEntry.RunAsync(["--help"]);
            await writer.FlushAsync();

            var output = outputBuffer.ToString();

            Assert.Equal(0, exitCode);
            Assert.Contains("Query2Excel Console", output);
            Assert.Contains("Usage:", output);
            Assert.Contains("--connectionString", output);
            Assert.Contains("--outputFilePath", output);
        }
        finally
        {
            Console.SetOut(originalOut);
        }
    }
}
