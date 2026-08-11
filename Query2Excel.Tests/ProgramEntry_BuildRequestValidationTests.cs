using System.Reflection;
using Microsoft.Extensions.Configuration;
using Query2Excel.Core.Models;

namespace Query2Excel.Tests;

public sealed class ProgramEntry_BuildRequestValidationTests
{
    [Fact]
    public void BuildRequest_WithMissingConnectionString_ThrowsArgumentException()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["sqlScript"] = @"C:\Temp\script.sql",
                ["outputFilePath"] = @"C:\Temp\output.xlsx"
            })
            .Build();

        var exception = AssertBuildRequestThrowsArgumentException(configuration);
        Assert.Equal("Connection string is required. Supply --connectionString, ConnectionStrings:Query2Excel, or Query2Excel:ConnectionString.", exception.Message);
    }

    [Fact]
    public void BuildRequest_WithInvalidTimeout_ThrowsArgumentException()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["connectionString"] = "Server=.;Database=Test;User ID=u;Password=p;",
                ["sqlScript"] = @"C:\Temp\script.sql",
                ["outputFilePath"] = @"C:\Temp\output.xlsx",
                ["commandTimeoutSeconds"] = "0"
            })
            .Build();

        var exception = AssertBuildRequestThrowsArgumentException(configuration);
        Assert.Equal("Command timeout must be a positive integer.", exception.Message);
    }

    [Fact]
    public void BuildRequest_WithValidValues_ReturnsRequest()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["connectionString"] = "Server=.;Database=Test;User ID=u;Password=p;",
                ["sqlScript"] = @"C:\Temp\script.sql",
                ["outputFilePath"] = @"C:\Temp\output.xlsx",
                ["commandTimeoutSeconds"] = "45"
            })
            .Build();

        var request = InvokeBuildRequest(configuration);

        Assert.Equal(@"C:\Temp\script.sql", request.SqlScriptPath);
        Assert.Equal(@"C:\Temp\output.xlsx", request.OutputFilePath);
        Assert.Equal(45, request.CommandTimeoutSeconds);
    }

    [Fact]
    public void BuildRequest_WithTokenizedConnectionStringTemplate_ResolvesCredentialsFromConsoleSecrets()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["connectionStringTemplate"] = "Server=.;Database=Test;User ID={UserId};Password={Password};",
                ["Query2ExcelApp:DatabaseUserId"] = "dbUser",
                ["Query2ExcelApp:DatabasePassword"] = "dbPassword",
                ["sqlScript"] = @"C:\Temp\script.sql",
                ["outputFilePath"] = @"C:\Temp\output.xlsx"
            })
            .Build();

        var request = InvokeBuildRequest(configuration);

        Assert.Equal("Server=.;Database=Test;User ID=dbUser;Password=dbPassword;", request.ConnectionString);
        Assert.Equal("Server=.;Database=Test;User ID={UserId};Password={Password};", request.ConnectionStringTemplate);
    }

    [Fact]
    public void BuildRequest_WithTokenizedTemplateAndMissingPasswordSecret_ThrowsArgumentException()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["connectionStringTemplate"] = "Server=.;Database=Test;User ID={UserId};Password={Password};",
                ["Query2ExcelApp:DatabaseUserId"] = "dbUser",
                ["sqlScript"] = @"C:\Temp\script.sql",
                ["outputFilePath"] = @"C:\Temp\output.xlsx"
            })
            .Build();

        var exception = AssertBuildRequestThrowsArgumentException(configuration);
        Assert.Equal("Database password token was provided, but no value was found for Query2ExcelApp:DatabasePassword or Query2Excel:DatabasePassword.", exception.Message);
    }

    private static Query2ExcelRequest InvokeBuildRequest(IConfiguration configuration)
    {
        var method = typeof(ProgramEntry).GetMethod("BuildRequest", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        try
        {
            return (Query2ExcelRequest)method.Invoke(null, new object[] { configuration })!;
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            throw ex.InnerException;
        }
    }

    private static ArgumentException AssertBuildRequestThrowsArgumentException(IConfiguration configuration)
    {
        var exception = Record.Exception(() => InvokeBuildRequest(configuration));
        Assert.NotNull(exception);
        return Assert.IsType<ArgumentException>(exception);
    }
}
