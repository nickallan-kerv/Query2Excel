using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Query2Excel.App.Abstractions;
using Query2Excel.App.Configuration;
using Query2Excel.App.Models;
using Query2Excel.App.Services;
using Serilog;

return await ProgramEntry.RunAsync(args).ConfigureAwait(false);

internal static class ProgramEntry
{
	public static async Task<int> RunAsync(string[] args)
	{
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
		var connectionString = FirstNonEmpty(
			configuration.GetConnectionString("Query2Excel"),
			configuration["connectionString"],
			configuration["ConnectionString"],
			configuration[$"{Query2ExcelOptions.SectionName}:ConnectionString"]);

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

		var timeout = 120;
		if (!string.IsNullOrWhiteSpace(timeoutRaw) && (!int.TryParse(timeoutRaw, out timeout) || timeout <= 0))
		{
			throw new ArgumentException("Command timeout must be a positive integer.");
		}

		return new Query2ExcelRequest(connectionString, sqlScriptPath, outputFilePath, timeout);
	}

	private static string? FirstNonEmpty(params string?[] candidates)
	{
		return candidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate));
	}
}
