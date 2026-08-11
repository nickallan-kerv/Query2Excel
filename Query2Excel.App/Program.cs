using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Query2Excel.Core.Abstractions;
using Query2Excel.Core.Configuration;
using Query2Excel.Core.Models;
using Query2Excel.Core.Services;
using Serilog;

return await ProgramEntry.RunAsync(args).ConfigureAwait(false);

internal static class ProgramEntry
{
	public static async Task<int> RunAsync(string[] args)
	{
		if (IsHelpRequested(args))
		{
			PrintHelp();
			return 0;
		}

		using var host = CreateHostBuilder(args).Build();
		using var scope = host.Services.CreateScope();

		var logger = scope.ServiceProvider.GetRequiredService<ILoggerFactory>().CreateLogger("Query2Excel");

		try
		{
			var configuration = scope.ServiceProvider.GetRequiredService<IConfiguration>();
			var request = BuildRequest(configuration);
			var service = scope.ServiceProvider.GetRequiredService<IQuery2ExcelService>();

			await service.RunAsync(request, CancellationToken.None).ConfigureAwait(false);

			logger.LogInformation("Workbook generation completed successfully.");
			return 0;
		}
		catch (ArgumentException argumentException)
		{
			logger.LogError(argumentException, "Input validation failed.");
			return 1;
		}
		catch (InvalidOperationException invalidOperationException)
		{
			logger.LogError(invalidOperationException, "Execution failed.");
			return 2;
		}
		catch (Exception exception)
		{
			logger.LogError(exception, "Unexpected failure.");
			return 99;
		}
	}

	private static IHostBuilder CreateHostBuilder(string[] args)
	{
		return Host.CreateDefaultBuilder(args)
			.ConfigureAppConfiguration((context, configurationBuilder) =>
			{
				configurationBuilder.AddJsonFile(
					Path.Combine(AppContext.BaseDirectory, "Query2Excel.RowStyles.json"),
					optional: true,
					reloadOnChange: false);

				// Ensure the app project settings are loaded even when launched from a different working directory.
				configurationBuilder.AddJsonFile(
					Path.Combine(AppContext.BaseDirectory, "appsettings.json"),
					optional: true,
					reloadOnChange: false);

				configurationBuilder.AddJsonFile(
					Path.Combine(AppContext.BaseDirectory, $"appsettings.{context.HostingEnvironment.EnvironmentName}.json"),
					optional: true,
					reloadOnChange: false);

				configurationBuilder.AddUserSecrets(typeof(ProgramEntry).Assembly, optional: true);
			})
			.UseSerilog((context, services, loggerConfiguration) =>
			{
				loggerConfiguration
					.ReadFrom.Configuration(context.Configuration)
					.ReadFrom.Services(services)
					.Enrich.FromLogContext();
			})
			.ConfigureServices((context, services) =>
			{
				services.Configure<Query2ExcelOptions>(context.Configuration.GetSection(Query2ExcelOptions.SectionName));

				services.AddSingleton<IDatabaseExecutor, DatabaseExecutor>();
				services.AddSingleton<IWorkbookBuilder, WorkbookBuilder>();
				services.AddSingleton<IQuery2ExcelService, Query2ExcelService>();
			});
	}

	private static Query2ExcelRequest BuildRequest(IConfiguration configuration)
	{
		var connectionStringTemplate = FirstNonEmpty(
			configuration["connectionStringTemplate"],
			configuration["ConnectionStringTemplate"],
			configuration[$"{Query2ExcelOptions.SectionName}:ConnectionStringTemplate"]);

		var connectionString = FirstNonEmpty(
			configuration.GetConnectionString("Query2Excel"),
			configuration["connectionString"],
			configuration["ConnectionString"],
			configuration[$"{Query2ExcelOptions.SectionName}:ConnectionString"]);

		if (!string.IsNullOrWhiteSpace(connectionStringTemplate))
		{
			connectionString = BuildConnectionStringFromTemplate(connectionStringTemplate, configuration);
		}

		var sqlScriptPath = FirstNonEmpty(
			configuration["sqlScript"],
			configuration["SqlScript"],
			configuration[$"{Query2ExcelOptions.SectionName}:SqlScript"],
			Query2ExcelOptions.DefaultSqlScriptPath);

		var outputFilePath = FirstNonEmpty(
			configuration["outputFilePath"],
			configuration["OutputFilePath"],
			configuration[$"{Query2ExcelOptions.SectionName}:OutputFilePath"]);

		var timeoutRaw = FirstNonEmpty(
			configuration["commandTimeoutSeconds"],
			configuration["CommandTimeoutSeconds"],
			configuration[$"{Query2ExcelOptions.SectionName}:CommandTimeoutSeconds"]);

		if (string.IsNullOrWhiteSpace(connectionString))
		{
			throw new ArgumentException("Connection string is required. Supply --connectionString, ConnectionStrings:Query2Excel, or Query2Excel:ConnectionString.");
		}

		if (string.IsNullOrWhiteSpace(sqlScriptPath))
		{
			throw new ArgumentException("SQL script path is required. Supply --sqlScript or Query2Excel:SqlScript.");
		}

		if (string.IsNullOrWhiteSpace(outputFilePath))
		{
			throw new ArgumentException("Output file path is required. Supply --outputFilePath or Query2Excel:OutputFilePath.");
		}

		ConsolePathValidation.ValidateFileExtensions(sqlScriptPath, outputFilePath);

		var timeout = 120;
		if (!string.IsNullOrWhiteSpace(timeoutRaw) && (!int.TryParse(timeoutRaw, out timeout) || timeout <= 0))
		{
			throw new ArgumentException("Command timeout must be a positive integer.");
		}

		return new Query2ExcelRequest(connectionString, sqlScriptPath, outputFilePath, timeout, connectionStringTemplate);
	}

	private static string BuildConnectionStringFromTemplate(string connectionStringTemplate, IConfiguration configuration)
	{
		var template = connectionStringTemplate.Trim();
		if (string.IsNullOrWhiteSpace(template))
		{
			throw new ArgumentException("Connection string template is required when --connectionStringTemplate is provided.");
		}

		var userIdTokens = new[] { "{UserId}", "{{USER_ID}}" };
		var passwordTokens = new[] { "{Password}", "{{PASSWORD}}" };

		var containsUserIdToken = userIdTokens.Any(token => template.Contains(token, StringComparison.OrdinalIgnoreCase));
		var containsPasswordToken = passwordTokens.Any(token => template.Contains(token, StringComparison.OrdinalIgnoreCase));

		if (!containsUserIdToken && !containsPasswordToken)
		{
			return template;
		}

		var userId = FirstNonEmpty(
			configuration["databaseUserId"],
			configuration["DatabaseUserId"],
			configuration["Query2ExcelApp:DatabaseUserId"],
			configuration["Query2Excel:DatabaseUserId"]);

		var password = FirstNonEmpty(
			configuration["databasePassword"],
			configuration["DatabasePassword"],
			configuration["Query2ExcelApp:DatabasePassword"],
			configuration["Query2Excel:DatabasePassword"]);

		if (containsUserIdToken && string.IsNullOrWhiteSpace(userId))
		{
			throw new ArgumentException("Database user id token was provided, but no value was found for Query2ExcelApp:DatabaseUserId or Query2Excel:DatabaseUserId.");
		}

		if (containsPasswordToken && string.IsNullOrWhiteSpace(password))
		{
			throw new ArgumentException("Database password token was provided, but no value was found for Query2ExcelApp:DatabasePassword or Query2Excel:DatabasePassword.");
		}

		if (!string.IsNullOrWhiteSpace(userId))
		{
			foreach (var token in userIdTokens)
			{
				template = template.Replace(token, userId, StringComparison.OrdinalIgnoreCase);
			}
		}

		if (!string.IsNullOrWhiteSpace(password))
		{
			foreach (var token in passwordTokens)
			{
				template = template.Replace(token, password, StringComparison.OrdinalIgnoreCase);
			}
		}

		return template;
	}

	private static string? FirstNonEmpty(params string?[] candidates)
	{
		return candidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate));
	}

	private static bool IsHelpRequested(string[] args)
	{
		return args.Any(argument =>
			string.Equals(argument, "--help", StringComparison.OrdinalIgnoreCase)
			|| string.Equals(argument, "-h", StringComparison.OrdinalIgnoreCase)
			|| string.Equals(argument, "/?", StringComparison.OrdinalIgnoreCase));
	}

	private static void PrintHelp()
	{
		var helpText = """
Query2Excel Console

Usage:
	dotnet run --project .\\Query2Excel.App -- [options]

Options:
	--help, -h, /?                 Show this help text.
	--connectionString             Full database connection string.
	--connectionStringTemplate     Tokenized connection string template (supports {UserId} and {Password}).
	--databaseUserId               User id for template token replacement.
	--databasePassword             Password for template token replacement.
	--sqlScript                    Path to .sql script file.
	--outputFilePath               Path to output .xlsx file.
	--commandTimeoutSeconds        Command timeout in seconds.

Configuration keys (equivalent):
	Query2Excel:ConnectionString
	Query2Excel:ConnectionStringTemplate
	Query2Excel:DatabaseUserId
	Query2Excel:DatabasePassword
	Query2Excel:SqlScript
	Query2Excel:OutputFilePath
	Query2Excel:CommandTimeoutSeconds

Examples:
	dotnet run --project .\\Query2Excel.App
	dotnet run --project .\\Query2Excel.App -- --sqlScript .\\Query2Excel.App\\Scripts\\Smoke.sql --outputFilePath C:\\temp\\output.xlsx
	dotnet run --project .\\Query2Excel.App -- --connectionStringTemplate "Server=localhost;Database=Db;User ID={UserId};Password={Password};TrustServerCertificate=True;"
""";

		Console.WriteLine(helpText);
	}
}

internal static class ConsolePathValidation
{
	public static void ValidateFileExtensions(string sqlScriptPath, string outputFilePath)
	{
		if (!string.Equals(Path.GetExtension(sqlScriptPath), ".sql", StringComparison.OrdinalIgnoreCase))
		{
			throw new ArgumentException("SQL script path must point to a .sql file.");
		}

		if (!string.Equals(Path.GetExtension(outputFilePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
		{
			throw new ArgumentException("Output file path must point to a .xlsx file.");
		}
	}
}
