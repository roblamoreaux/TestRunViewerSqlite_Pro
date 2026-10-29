
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
Open a results file on startup (loads the overview automatically):
```bash
TestRunViewerSqlite.exe "C:\path\to\Results.TapResults"
dotnet run --project TestRunViewerSqlite.csproj -- "C:\path\to\Results.TapResults"
```
Build the installer (needs the WiX tool: `dotnet tool install --global wix --version 5.0.2`):
```bash
powershell -ExecutionPolicy Bypass -File Installer\build-installer.ps1
```
Output: `bin\Installer\TestRunViewerSqlite-<version>-x64.msi` (self-contained, no .NET runtime needed on the target PC).
To release a new version, bump `<Version>` in `TestRunViewerSqlite.csproj` and rebuild; installing the new MSI replaces the old version.
