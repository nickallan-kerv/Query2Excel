using Microsoft.Extensions.Logging;
using Query2Excel.Core.Abstractions;
using Query2Excel.Core.Models;

namespace Query2Excel.Core.Services;

public sealed class Query2ExcelService(
    IDatabaseExecutor databaseExecutor,
    IWorkbookBuilder workbookBuilder,
    ILogger<Query2ExcelService> logger) : IQuery2ExcelService
{
    public async Task RunAsync(Query2ExcelRequest request, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.ConnectionString);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.SqlScriptPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.OutputFilePath);

        var sqlScript = await LoadSqlScriptAsync(request.SqlScriptPath, cancellationToken).ConfigureAwait(false);

        logger.LogInformation(
            "Starting Query2Excel workflow. SqlScriptPath: {SqlScriptPath}. OutputPath: {OutputFilePath}.",
            request.SqlScriptPath,
            request.OutputFilePath);

        var executionResult = await databaseExecutor
            .ExecuteAsync(
                new QueryExecutionRequest(request.ConnectionString, sqlScript, request.CommandTimeoutSeconds),
                cancellationToken)
            .ConfigureAwait(false);

        var sanitizedConnectionStringTemplate = ConnectionStringTemplateProtector.SanitizeForWorkbook(
            string.IsNullOrWhiteSpace(request.ConnectionStringTemplate)
                ? request.ConnectionString
                : request.ConnectionStringTemplate);

        executionResult = executionResult with
        {
            ConnectionStringTemplate = sanitizedConnectionStringTemplate
        };

        await workbookBuilder
            .BuildWorkbookAsync(executionResult, request.OutputFilePath, cancellationToken)
            .ConfigureAwait(false);

        logger.LogInformation(
            "Query2Excel workflow completed. ResultSetCount: {ResultSetCount}. RowsReturned: {RowCount}. OutputPath: {OutputFilePath}.",
            executionResult.ResultSetCount,
            executionResult.TotalRowCount,
            request.OutputFilePath);
    }

    private static async Task<string> LoadSqlScriptAsync(string sqlScriptPath, CancellationToken cancellationToken)
    {
        if (!TryResolveSqlScriptPath(sqlScriptPath, out var normalizedPath, out var checkedPaths))
        {
            throw new ArgumentException(
                $"SQL script file was not found. Input path: '{sqlScriptPath}'. Checked: {string.Join("; ", checkedPaths)}");
        }

        var script = await File.ReadAllTextAsync(normalizedPath, cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(script))
        {
            throw new ArgumentException($"SQL script file is empty: {normalizedPath}");
        }

        return script;
    }

    private static bool TryResolveSqlScriptPath(string sqlScriptPath, out string resolvedPath, out List<string> checkedPaths)
    {
        checkedPaths = new List<string>();

        if (Path.IsPathRooted(sqlScriptPath))
        {
            resolvedPath = sqlScriptPath;
            checkedPaths.Add(resolvedPath);
            return File.Exists(resolvedPath);
        }

        var normalizedRelativePath = sqlScriptPath
            .Replace('/', Path.DirectorySeparatorChar)
            .Replace('\\', Path.DirectorySeparatorChar);

        var relativeCandidates = new List<string> { normalizedRelativePath };
        var projectPrefix = $"Query2Excel.App{Path.DirectorySeparatorChar}";
        if (normalizedRelativePath.StartsWith(projectPrefix, StringComparison.OrdinalIgnoreCase))
        {
            relativeCandidates.Add(normalizedRelativePath[projectPrefix.Length..]);
        }

        var basePaths = new[]
        {
            Directory.GetCurrentDirectory(),
            AppContext.BaseDirectory
        };

        foreach (var basePath in basePaths)
        {
            foreach (var relativeCandidate in relativeCandidates)
            {
                var candidate = Path.GetFullPath(Path.Combine(basePath, relativeCandidate));
                checkedPaths.Add(candidate);
                if (File.Exists(candidate))
                {
                    resolvedPath = candidate;
                    return true;
                }
            }

            foreach (var candidate in EnumerateAncestorCandidates(basePath, relativeCandidates))
            {
                checkedPaths.Add(candidate);
                if (File.Exists(candidate))
                {
                    resolvedPath = candidate;
                    return true;
                }
            }
        }

        resolvedPath = checkedPaths.FirstOrDefault() ?? sqlScriptPath;
        return false;
    }

    private static IEnumerable<string> EnumerateAncestorCandidates(string basePath, IReadOnlyList<string> relativeCandidates)
    {
        var current = new DirectoryInfo(basePath);
        while (current is not null)
        {
            foreach (var relativeCandidate in relativeCandidates)
            {
                yield return Path.GetFullPath(Path.Combine(current.FullName, relativeCandidate));
            }

            current = current.Parent;
        }
    }
}

