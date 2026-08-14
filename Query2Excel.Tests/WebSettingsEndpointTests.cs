extern alias webhost;

using System.Net.Http.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.Configuration;

namespace Query2Excel.Tests;

public sealed class WebSettingsEndpointTests
{
    [Fact]
    public async Task GetSettings_WithConfiguredRowStyles_ReturnsRowStyleNames()
    {
        await using var factory = new Query2ExcelWebFactory(new Dictionary<string, string?>
        {
            ["Query2ExcelWeb:LaunchBrowserOnStartup"] = "false",
            ["Query2Excel:RowStyles:Accent1:BackgroundColor"] = "#DCE6F1",
            ["Query2Excel:RowStyles:Accent2:BackgroundColor"] = "#F2DCDB"
        });

        using var client = factory.CreateClient();
        var settings = await client.GetFromJsonAsync<WebSettingsResponseContract>("/settings");

        Assert.NotNull(settings);
        Assert.Contains("Accent1", settings.RowStyleNames);
        Assert.Contains("Accent2", settings.RowStyleNames);
    }

    [Fact]
    public async Task GetSettings_WithoutConfiguredRowStyles_ReturnsEmptyRowStyleNames()
    {
        await using var factory = new Query2ExcelWebFactory(
            new Dictionary<string, string?>
            {
                ["Query2ExcelWeb:LaunchBrowserOnStartup"] = "false"
            },
            clearDefaultConfigSources: true);

        using var client = factory.CreateClient();
        var settings = await client.GetFromJsonAsync<WebSettingsResponseContract>("/settings");

        Assert.NotNull(settings);
        Assert.NotNull(settings.RowStyleNames);
        Assert.Empty(settings.RowStyleNames);
    }

    private sealed class Query2ExcelWebFactory : WebApplicationFactory<webhost::Program>
    {
        private readonly IReadOnlyDictionary<string, string?> _overrides;
        private readonly bool _clearDefaultConfigSources;

        public Query2ExcelWebFactory(
            IReadOnlyDictionary<string, string?> overrides,
            bool clearDefaultConfigSources = false)
        {
            _overrides = overrides;
            _clearDefaultConfigSources = clearDefaultConfigSources;
        }

        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseEnvironment("Development");

            builder.ConfigureAppConfiguration((_, config) =>
            {
                if (_clearDefaultConfigSources)
                {
                    config.Sources.Clear();
                }

                config.AddInMemoryCollection(_overrides);
            });
        }
    }

    private sealed class WebSettingsResponseContract
    {
        public string[] RowStyleNames { get; init; } = Array.Empty<string>();
    }
}
