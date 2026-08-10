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

## Metadata Fieldset

Query2Excel supports metadata result sets that control how the next output worksheet is rendered.

### How metadata is recognized

A result set is treated as metadata when all of the following are true:

- It has exactly 1 row.
- It has one or more columns.
- Every column name is a recognized metadata field.

If any column is unrecognized, the result set is treated as a normal output table.

### Supported metadata fields

- __SheetName
  - Optional
  - Overrides the default worksheet name for the next output result set.
  - If omitted, default naming is Output1, Output2, and so on.

- __Title
  - Optional
  - Inserts a title line above the table.
  - Styled as a worksheet heading.

- __Description
  - Optional
  - Inserts a descriptive line between title and table, or above the table when no title is provided.

### Row placement rules

- If __Title and __Description are both present:
  - A1 = title
  - A2 = description
  - Table header starts at A3

- If only __Title is present:
  - A1 = title
  - Table header starts at A2

- If only __Description is present:
  - A1 = description
  - Table header starts at A2

- If neither is present:
  - Table header starts at A1

### Metadata scope

- Metadata always applies to the immediately following data result set.
- Metadata result sets are consumed and are not emitted as output worksheets.

### Autosizing behavior

- Table column auto-fit uses only the table header/data region.
- Long __Title or __Description values do not affect table column widths.

### Example

```sql
SELECT __SheetName = 'Objects',
       __Title = 'List of Objects',
       __Description = 'This query returns a list of all objects in the database, including tables, views, and stored procedures.';
EXEC sp_find;

SELECT __Description = 'This query returns a list of currently active users and their processes.';
EXEC sp_who2;
```

Expected workbook output:

- Worksheet 1: Objects
  - A1 title, A2 description, table begins at A3
- Worksheet 2: Output2
  - A1 description, table begins at A2

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
