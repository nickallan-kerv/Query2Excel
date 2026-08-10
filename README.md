# Query2Excel

## SQL Script Configuration

`Query2Excel:SqlScript` points to a `.sql` file that can contain multiline SQL and multiple statements.

Default value in `appsettings.json`:

```json
"SqlScript": "Query2Excel.App\\Scripts\\Example.sql"
```

If the script returns multiple result sets, the workbook will contain:

- `Output1` for result set 1
- `Output2` for result set 2
- Additional sheets (`Output3`, `Output4`, ...) for any remaining result sets

The `SQL` worksheet still contains the exact executed script plus execution timestamp, duration, total rows, and total result set count.

## Local Secret Setup (Connection String)

This project uses .NET User Secrets so database credentials are not stored in source control.

### Project Configuration

- Project file: Query2Excel.App/Query2Excel.App.csproj
- UserSecretsId: query2excel-89aa31cd-b4bc-4f7c-a2f3-5fc7e6117f18
- Secret key used by the app: ConnectionStrings:Query2Excel

### Set the Secret (exact command)

Run from the repository root:

```powershell
dotnet user-secrets --project .\Query2Excel.App set "ConnectionStrings:Query2Excel" "Server=localhost;Database=CDMSandbox;User Id=nickallan;Password=YOUR_PASSWORD;TrustServerCertificate=True;"
```

### Verify the Secret

```powershell
dotnet user-secrets --project .\Query2Excel.App list
```

You should see a value for `ConnectionStrings:Query2Excel`.

### Where It Is Stored (Windows)

User Secrets are stored per-user outside the repository:

- `%APPDATA%\Microsoft\UserSecrets\<UserSecretsId>\secrets.json`
- Example for this project:
  `C:\Users\NickAllan\AppData\Roaming\Microsoft\UserSecrets\query2excel-89aa31cd-b4bc-4f7c-a2f3-5fc7e6117f18\secrets.json`

### Notes

- Keep `Query2Excel.App/appsettings.json` checked in with `"ConnectionString": ""`.
- Do not commit plaintext credentials.
- For CI/production, prefer environment variables or Azure Key Vault.
