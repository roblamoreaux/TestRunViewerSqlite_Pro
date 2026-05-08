
# TestRunViewer (SQLite, WinForms, .NET 8)

This project registers **custom aggregate functions** `variance()` and `stdev()` in SQLite so statistics queries can use them directly. The implementation uses Welford's algorithm (sample variance & stdev) via **Microsoft.Data.Sqlite** `CreateAggregate`.

**Docs**
- Microsoft Learn: User-defined functions
- Microsoft Learn: `SqliteConnection.CreateAggregate`

Build:
```bash
dotnet build
```
Run:
```bash
dotnet run --project TestRunViewerSqlite.csproj
```
