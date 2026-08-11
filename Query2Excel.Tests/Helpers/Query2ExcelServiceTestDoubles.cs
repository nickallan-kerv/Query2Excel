using Query2Excel.Core.Abstractions;
using Query2Excel.Core.Models;

namespace Query2Excel.Tests;

internal sealed class StaticResultDatabaseExecutor(QueryExecutionResult result) : IDatabaseExecutor
{
    public bool ExecuteCalled { get; private set; }

    public QueryExecutionRequest? LastRequest { get; private set; }

    public Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken)
    {
        ExecuteCalled = true;
        LastRequest = request;
        return Task.FromResult(result);
    }
}

internal sealed class ThrowingDatabaseExecutor(Exception exception) : IDatabaseExecutor
{
    public Task<QueryExecutionResult> ExecuteAsync(QueryExecutionRequest request, CancellationToken cancellationToken)
    {
        throw exception;
    }
}

internal sealed class TrackingWorkbookBuilder : IWorkbookBuilder
{
    public bool BuildCalled { get; private set; }

    public string? OutputFilePath { get; private set; }

    public Task BuildWorkbookAsync(QueryExecutionResult result, string outputFilePath, CancellationToken cancellationToken)
    {
        BuildCalled = true;
        OutputFilePath = outputFilePath;
        return Task.CompletedTask;
    }
}
