using System.Data.Common;

namespace Query2Excel.Core.Services;

internal static class ConnectionStringTemplateProtector
{
    private static readonly string[] UserIdKeys = ["User ID", "UserId", "UID", "User"];
    private static readonly string[] PasswordKeys = ["Password", "Pwd"];

    public static string SanitizeForWorkbook(string? connectionStringOrTemplate)
    {
        if (string.IsNullOrWhiteSpace(connectionStringOrTemplate))
        {
            return string.Empty;
        }

        var raw = connectionStringOrTemplate.Trim();

        try
        {
            var builder = new DbConnectionStringBuilder
            {
                ConnectionString = raw
            };

            MaskValueIfNeeded(builder, UserIdKeys, "{UserId}");
            MaskValueIfNeeded(builder, PasswordKeys, "{Password}");
            return builder.ConnectionString;
        }
        catch
        {
            // Leave unmodified if parsing fails; callers should pass a valid connection string/template.
            return raw;
        }
    }

    private static void MaskValueIfNeeded(DbConnectionStringBuilder builder, string[] candidateKeys, string token)
    {
        foreach (var key in candidateKeys)
        {
            if (!builder.TryGetValue(key, out var value))
            {
                continue;
            }

            var rawValue = value?.ToString()?.Trim() ?? string.Empty;
            if (IsTokenized(rawValue, token))
            {
                builder[key] = token;
                return;
            }

            if (!string.IsNullOrWhiteSpace(rawValue))
            {
                builder[key] = token;
                return;
            }
        }
    }

    private static bool IsTokenized(string value, string canonicalToken)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (value.Equals(canonicalToken, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(canonicalToken, "{UserId}", StringComparison.Ordinal))
        {
            return value.Equals("{{USER_ID}}", StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(canonicalToken, "{Password}", StringComparison.Ordinal))
        {
            return value.Equals("{{PASSWORD}}", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }
}
