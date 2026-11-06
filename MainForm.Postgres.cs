using Microsoft.Data.Sqlite;
using Npgsql;
using System.Data;
using System.Data.Common;
using System.Text.RegularExpressions;

namespace TestRunViewerSqlite
{
    // PostgreSQL support. OpenTAP's PostgreSQL result listener uses the same tables as the
    // SQLite one, so the app's SQLite queries are reused and translated where the dialects differ.
    public partial class MainForm
    {
        private string? _pgConnectionString; // null = the SQLite file in _dbPath is the source
        private string _pgHost = "localhost";
        private int _pgPort = 5432;
        private string _pgDatabase = string.Empty;
        private string _pgUsername = "postgres";

        private bool IsPostgres => _pgConnectionString != null;
        private bool HasDatabase => IsPostgres || File.Exists(_dbPath);

        private async Task<DbConnection> OpenConnectionAsync()
        {
            DbConnection conn = IsPostgres
                ? new NpgsqlConnection(_pgConnectionString)
                : new SqliteConnection(ConnectionString);
            try
            {
                await conn.OpenAsync();
            }
            catch
            {
                conn.Dispose();
                throw;
            }
            if (conn is SqliteConnection sqlite)
                RegisterCustomFunctions(sqlite);
            return conn;
        }

        private DbCommand CreateCommand(DbConnection conn, string sql)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = IsPostgres ? ToPostgres(sql) : sql;
            return cmd;
        }

        private static void AddParameter(DbCommand cmd, string name, object? value)
        {
            var parameter = cmd.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(parameter);
        }

        // Rewrites the parts of the SQLite queries that PostgreSQL doesn't accept.
        private static string ToPostgres(string sql)
        {
            const RegexOptions ignoreCase = RegexOptions.IgnoreCase;

            // stdev/variance are the app's own SQLite aggregates (sample, 0 when there's one value)
            sql = Regex.Replace(sql, @"\bstdev\s*\(([^()]*)\)", "COALESCE(stddev_samp($1), 0.0)", ignoreCase);
            sql = Regex.Replace(sql, @"\bvariance\s*\(([^()]*)\)", "COALESCE(var_samp($1), 0.0)", ignoreCase);

            // group_concat -> string_agg; unnamed ones keep SQLite's column name (the expression text)
            sql = Regex.Replace(sql, @"\bgroup_concat\s*\(\s*([^,()]+?)\s*,\s*('[^']*')\s*\)(\s+as\b)?",
                m => $"string_agg(CAST({m.Groups[1].Value} AS text), {m.Groups[2].Value})" +
                     (m.Groups[3].Success ? m.Groups[3].Value : $" AS \"{m.Value}\""), ignoreCase);

            // COALESCE can't mix text and numbers in PostgreSQL
            sql = Regex.Replace(sql, @"\bCOALESCE\(\s*([\w.]+(?:MinLimit|MaxLimit))\s*,\s*0\.0\s*\)", "COALESCE($1, '0.0')", ignoreCase);
            sql = Regex.Replace(sql, @"\bCOALESCE\(\s*result\.Dim0\s*,\s*0\.0\s*\)", _ => SqliteReal("COALESCE(result.Dim0, '0.0')"), ignoreCase);

            // CAST(text AS REAL) fails on non-numeric text in PostgreSQL
            sql = Regex.Replace(sql, @"\bcast\s*\(\s*([^()]+?)\s+as\s+real\s*\)", m => SqliteReal(m.Groups[1].Value), ignoreCase);
            sql = Regex.Replace(sql, @"\bas\s+real\b", "AS double precision", ignoreCase);

            return WrapUngroupedColumns(sql);
        }

        // SQLite's CAST(x AS REAL): a double, NULL stays NULL, and non-numeric text becomes 0 instead of an error.
        private static string SqliteReal(string expression) =>
            $"(CASE WHEN ({expression}) IS NULL THEN NULL " +
            $@"WHEN CAST(({expression}) AS text) ~ '^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?\s*$' " +
            $"THEN CAST(CAST(({expression}) AS text) AS double precision) ELSE 0.0 END)";

        // Columns the queries' final SELECT returns from a GROUP BY without an aggregate. SQLite takes
        // the value from one row of the group; PostgreSQL requires an aggregate, so pick one with array_agg.
        private static readonly string[] UngroupedColumns =
        {
            "result.dim1", "result.Dim2", "result.Dim3", "GetStepMinLimit.MinLimit", "GetStepMaxLimit.MaxLimit",
            "GetStepVerdict.Verdict", "GetStepCheckLimit.CheckLimit",                                  // Details
            "GetScriptName.ScriptName", "GetReport.MinLimit", "GetReport.MaxLimit", "GetReport.SignalName",
            "GetReport.CheckLimit", "GetReport.SerialNumber", "GetReport.Signal",                      // Stats
        };

        private static string WrapUngroupedColumns(string sql)
        {
            const RegexOptions lines = RegexOptions.IgnoreCase | RegexOptions.Multiline;
            var selects = Regex.Matches(sql, @"^\s*select\s*$", lines);
            if (selects.Count == 0)
                return sql;

            int start = selects[^1].Index + selects[^1].Length;
            var from = Regex.Match(sql[start..], @"^\s*from\s*$", lines);
            if (!from.Success || !Regex.IsMatch(sql[start..], @"\bgroup\s+by\b", RegexOptions.IgnoreCase))
                return sql;

            int end = start + from.Index;
            var selectList = sql[start..end];
            foreach (var column in UngroupedColumns)
            {
                var c = Regex.Escape(column);
                // A column on its own line keeps its name; elsewhere (inside COALESCE, before "as") just wrap it
                selectList = Regex.Replace(selectList, $@"^(\s*)({c})(\s*,?\s*)$",
                    m => $"{m.Groups[1].Value}(array_agg({m.Groups[2].Value}))[1] AS {column.Split('.')[1]}{m.Groups[3].Value}", lines);
                selectList = Regex.Replace(selectList, $@"(?<![\w.])(?<!array_agg\(){c}(?!\w)",
                    m => $"(array_agg({m.Value}))[1]", RegexOptions.IgnoreCase);
            }
            return sql[..start] + selectList + sql[end..];
        }

        // PostgreSQL returns unquoted column names in lower case; use the casing written in the SQL.
        private static void RestoreColumnCase(DataTable table, string sql)
        {
            var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(sql, @"\bas\s+([A-Za-z_]\w*)", RegexOptions.IgnoreCase))
                names.TryAdd(m.Groups[1].Value, m.Groups[1].Value);
            foreach (Match m in Regex.Matches(sql, @"\b[A-Za-z_]\w*\b"))
                if (m.Value.Any(char.IsUpper))
                    names.TryAdd(m.Value, m.Value);

            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName == column.ColumnName.ToLowerInvariant() &&
                    names.TryGetValue(column.ColumnName, out var name))
                    column.ColumnName = name;
            }
        }

        private async void btnOpenPostgres_Click(object sender, EventArgs e)
        {
            using var dialog = new Form
            {
                Text = "Open PostgreSQL Database",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                AutoScaleMode = AutoScaleMode.Font
            };
            var layout = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Padding = new Padding(10) };

            TextBox AddField(string label, string value, bool password = false)
            {
                var box = new TextBox { Text = value, Width = 260, UseSystemPasswordChar = password };
                layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left });
                layout.Controls.Add(box);
                return box;
            }

            var host = AddField("Host:", _pgHost);
            var port = AddField("Port:", _pgPort.ToString());
            var database = AddField("Database:", _pgDatabase);
            var username = AddField("Username:", _pgUsername);
            var password = AddField("Password:", string.Empty, password: true);

            var connect = new Button { Text = "Connect", DialogResult = DialogResult.OK, AutoSize = true };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
            buttons.Controls.AddRange(new Control[] { cancel, connect });
            layout.Controls.Add(buttons);
            layout.SetColumnSpan(buttons, 2);

            dialog.Controls.Add(layout);
            dialog.AcceptButton = connect;
            dialog.CancelButton = cancel;
            dialog.Shown += (_, _) => (string.IsNullOrEmpty(database.Text) ? database : password).Focus();

            // Keep asking until the connection works or the user cancels
            while (dialog.ShowDialog(this) == DialogResult.OK)
            {
                if (!int.TryParse(port.Text.Trim(), out var portNumber))
                {
                    MessageBox.Show(this, "Port must be a number.", "Open PostgreSQL Database",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }

                var builder = new NpgsqlConnectionStringBuilder
                {
                    Host = host.Text.Trim(),
                    Port = portNumber,
                    Database = database.Text.Trim(),
                    Username = username.Text.Trim(),
                    Password = password.Text,
                    Options = "-c default_transaction_read_only=on", // the viewer never writes
                    CommandTimeout = 300
                };

                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    using var conn = new NpgsqlConnection(builder.ConnectionString);
                    await conn.OpenAsync();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "PostgreSQL Connection Failed",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }

                _pgHost = builder.Host!;
                _pgPort = portNumber;
                _pgDatabase = builder.Database!;
                _pgUsername = builder.Username!;
                _pgConnectionString = builder.ConnectionString;
                txtDbPath.Text = $"PostgreSQL: {_pgUsername}@{_pgHost}:{_pgPort}/{_pgDatabase}";

                SaveSettings();
                await LoadOverviewAsync();
                UpdateUiState();
                return;
            }
        }
    }
}
