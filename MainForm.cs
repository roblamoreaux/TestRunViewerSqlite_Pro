using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System.Data;
using System.Text;
using System.Text.Json;
namespace TestRunViewerSqlite
{
    public partial class MainForm : Form
    {
        private string _dbPath = string.Empty;
        private string ConnectionString => $"Data Source={_dbPath};Mode=ReadOnly;Cache=Shared"; private string SettingsFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TestRunViewerSqlite"); private string SettingsFile => Path.Combine(SettingsFolder, "settings.json");
        private readonly string _overviewSql = @"
WITH 
GetStepTimestamp AS (
  SELECT testrun2params.scope, testrun2params.runid, params.value AS TimeStamp
  FROM testrun
  INNER JOIN testrun2params ON testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE testrun2params.scope = testrun.runid AND testrun2params.runid = testrun.runid AND params.name = 'StartTime'
),
GetParentName AS (
  SELECT testrun.runid, testrun.parentrunid, params.Value AS ParentName
  FROM testrun
  INNER JOIN testrun2params ON testrun.ParentRunID = testrun2params.Scope AND testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE params.Name = 'Name' AND params.GroupName = ''
),
GetVerdict AS (
  SELECT planrun.planrunnumber, params.value as Verdict
  FROM planrun
  INNER JOIN testrun ON planrun.runid = testrun.runid
  INNER JOIN testrun2params ON testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE testrun2params.scope = testrun.runid AND testrun2params.runid = testrun.runid AND params.name = 'Verdict'
)
SELECT GetStepTimestamp.TimeStamp, GetParentName.ParentName as PlanRunName, planrun.PlanRunNumber, planrun.*, GetVerdict.Verdict
FROM planrun
INNER JOIN testrun ON planrun.runid = testrun.planrunid
LEFT  JOIN GetVerdict ON planrun.planrunnumber = GetVerdict.planrunnumber
LEFT  JOIN GetParentName ON testrun.RunID = GetParentName.runid
INNER JOIN GetStepTimestamp ON testrun.runid = GetStepTimestamp.runid
WHERE testrun.RunID = (PlanRun.RunID + 1)
ORDER BY planrun.PlanRunNumber, GetStepTimestamp.TimeStamp
";
        private readonly string _detailsSql = @"SELECT 1 AS Placeholder;";
        private readonly string _statsAllSql = @"SELECT 'metric' AS Metric, variance(1.0) AS Variance, stdev(1.0) AS StdDev;";
        private readonly string _statsBySerialSql = @"SELECT 'byserial' AS Report, variance(1.0) AS Variance, stdev(1.0) AS StdDev;";
        public MainForm()
        {
            InitializeComponent(); dgvOverview.AutoGenerateColumns = true;
            dgvDetails.AutoGenerateColumns = true; dgvStatsAll.AutoGenerateColumns = true;
            dgvStatsSerial.AutoGenerateColumns = true;
            cboVerdict.Items.AddRange(new object[] { "(Any)", "Pass", "Fail", "NotSet" });
            cboVerdict.SelectedIndex = 0; cboStatus.Items.Add("(Any)");
            cboStatus.SelectedIndex = 0; dtFrom.ShowCheckBox = true;
            dtTo.ShowCheckBox = true; LoadSettings();
            UpdateUiState();
        }
        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e); if (File.Exists(_dbPath)) { await LoadOverviewAsync(); }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSettings(); base.OnFormClosing(e);
        }
        private struct VarianceAccumulator
        {
            public long Count;
            public double Mean;
            public double M2; // Sum of squares of differences
        }
        private static void RegisterVariance(SqliteConnection conn)
        {
            conn.CreateAggregate<double?, VarianceAccumulator, double?>(
                name: "variance",
                seed: new VarianceAccumulator
                {
                    Count = 0,
                    Mean = 0.0,
                    M2 = 0.0
                },
                func: (acc, value) =>
                {
                    if (value.HasValue)
                    {
                        acc.Count++;
                        var delta = value.Value - acc.Mean;
                        acc.Mean += delta / acc.Count;
                        acc.M2 += delta * (value.Value - acc.Mean);
                    }
                    return acc;
                },
                resultSelector: acc =>
                    acc.Count > 1
                        ? acc.M2 / (acc.Count - 1)   // ✅ sample variance
                        : (double?)null,
                isDeterministic: true
            );
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("lastDbPath", out var el))
                    {
                        var p = el.GetString();
                        if (!string.IsNullOrWhiteSpace(p) && File.Exists(p))
                        {
                            _dbPath = p!; txtDbPath.Text = _dbPath;
                        }
                    }
                }
            }
            catch { }
        }
        private void SaveSettings() { try { Directory.CreateDirectory(SettingsFolder); var json = JsonSerializer.Serialize(new { lastDbPath = _dbPath }, new JsonSerializerOptions { WriteIndented = true }); File.WriteAllText(SettingsFile, json); } catch { } }
        private void UpdateUiState() { bool hasDb = File.Exists(_dbPath); btnLoad.Enabled = hasDb; btnLoadStatsAll.Enabled = hasDb; btnLoadStatsSerial.Enabled = hasDb && !string.IsNullOrWhiteSpace(txtSerial.Text); btnExportOverview.Enabled = dgvOverview.DataSource is DataTable dtO && dtO.Rows.Count > 0; btnExportDetails.Enabled = dgvDetails.DataSource is DataTable dtD && dtD.Rows.Count > 0; btnExportStatsAll.Enabled = dgvStatsAll.DataSource is DataTable dtA && dtA.Rows.Count > 0; btnExportStatsSerial.Enabled = dgvStatsSerial.DataSource is DataTable dtS && dtS.Rows.Count > 0; }
        private struct StatAcc { public long Count; public double Mean; public double M2; }
        private static bool _udfRegistered = false; private static void RegisterCustomFunctions(SqliteConnection conn) { if (_udfRegistered) return; conn.CreateAggregate<double?, StatAcc, double?>("variance", new StatAcc { Count = 0, Mean = 0.0, M2 = 0.0 }, (acc, x) => { if (x.HasValue) { acc.Count++; var delta = x.Value - acc.Mean; acc.Mean += delta / acc.Count; var delta2 = x.Value - acc.Mean; acc.M2 += delta * delta2; } return acc; }, acc => acc.Count > 1 ? (double?)(acc.M2 / (acc.Count - 1)) : null, true); conn.CreateAggregate<double?, StatAcc, double?>("stdev", new StatAcc { Count = 0, Mean = 0.0, M2 = 0.0 }, (acc, x) => { if (x.HasValue) { acc.Count++; var delta = x.Value - acc.Mean; acc.Mean += delta / acc.Count; var delta2 = x.Value - acc.Mean; acc.M2 += delta * delta2; } return acc; }, acc => acc.Count > 1 ? (double?)Math.Sqrt(acc.M2 / (acc.Count - 1)) : null, true); _udfRegistered = true; }
        private async Task<DataTable> ExecuteQueryAsync(string sql, Action<SqliteCommand> configure) { var table = new DataTable(); using var conn = new SqliteConnection(ConnectionString); await conn.OpenAsync(); RegisterCustomFunctions(conn); using var cmd = new SqliteCommand(sql, conn); configure?.Invoke(cmd); using var reader = await cmd.ExecuteReaderAsync(); table.Load(reader); return table; }
        private string BuildFilteredOverviewWrapperSql() { return "SELECT * FROM (" + _overviewSql + ") ov WHERE 1=1\n{FILTERS}\nORDER BY ov.PlanRunNumber, ov.TimeStamp"; }
        private async Task PopulateStatusFilterAsync() { try { var sql = "SELECT DISTINCT Status FROM (" + _overviewSql + ") WHERE Status IS NOT NULL ORDER BY 1"; var dt = await ExecuteQueryAsync(sql, null); cboStatus.Items.Clear(); cboStatus.Items.Add("(Any)"); foreach (DataRow r in dt.Rows) cboStatus.Items.Add(Convert.ToString(r[0]) ?? ""); cboStatus.SelectedIndex = 0; lblStatusFilter.Visible = cboStatus.Visible = true; } catch { lblStatusFilter.Visible = cboStatus.Visible = false; } }
        private void AddOverviewFilterParameters(SqliteCommand cmd) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", " {FILTERS} "); if (!string.IsNullOrWhiteSpace(txtPlanRunName.Text)) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.PlanRunName LIKE @name\n{FILTERS}"); cmd.Parameters.Add(new SqliteParameter("@name", "%" + txtPlanRunName.Text.Trim() + "%")); } if (cboVerdict.SelectedIndex > 0) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.Verdict = @verdict\n{FILTERS}"); cmd.Parameters.Add(new SqliteParameter("@verdict", Convert.ToString(cboVerdict.SelectedItem))); } if (cboStatus.Visible && cboStatus.SelectedIndex > 0) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.Status = @status\n{FILTERS}"); cmd.Parameters.Add(new SqliteParameter("@status", Convert.ToString(cboStatus.SelectedItem))); } if (dtFrom.Checked) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.TimeStamp >= @from\n{FILTERS}"); cmd.Parameters.Add(new SqliteParameter("@from", dtFrom.Value.ToString("yyyy-MM-dd HH:mm:ss"))); } if (dtTo.Checked) { cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.TimeStamp <= @to\n{FILTERS}"); cmd.Parameters.Add(new SqliteParameter("@to", dtTo.Value.ToString("yyyy-MM-dd HH:mm:ss"))); } cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", string.Empty); }
        private async Task LoadOverviewAsync() { if (!File.Exists(_dbPath)) { MessageBox.Show(this, "Please open a SQLite database first.", "No Database", MessageBoxButtons.OK, MessageBoxIcon.Information); return; } ToggleUi(false); lblStatus.Text = "Loading overview..."; try { await PopulateStatusFilterAsync(); var sql = BuildFilteredOverviewWrapperSql(); var table = await ExecuteQueryAsync(sql, cmd => AddOverviewFilterParameters(cmd)); dgvOverview.DataSource = table; lblStatus.Text = $"Loaded {table.Rows.Count} rows."; cmbMapPlanRunNumber.Items.Clear(); foreach (DataColumn c in table.Columns) cmbMapPlanRunNumber.Items.Add(c.ColumnName); var defaultCol = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).FirstOrDefault(n => string.Equals(n, "PlanRunNumber", StringComparison.OrdinalIgnoreCase)); if (!string.IsNullOrEmpty(defaultCol)) cmbMapPlanRunNumber.SelectedItem = defaultCol; else if (cmbMapPlanRunNumber.Items.Count > 0) cmbMapPlanRunNumber.SelectedIndex = 0; if (table.Rows.Count > 0) { dgvOverview.ClearSelection(); dgvOverview.Rows[0].Selected = true; await LoadSelectedDetailsAsync(); } else { dgvDetails.DataSource = null; } } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Error Loading Overview", MessageBoxButtons.OK, MessageBoxIcon.Error); lblStatus.Text = "Failed to load overview."; } finally { ToggleUi(true); UpdateUiState(); } }
        private object? GetSelectedOverviewCell(string columnName) { if (dgvOverview.CurrentRow == null) return null; var view = dgvOverview.CurrentRow.DataBoundItem as DataRowView; if (view == null) return null; var row = view.Row; if (!row.Table.Columns.Contains(columnName)) return null; return row[columnName]; }
        private async Task LoadSelectedDetailsAsync() { if (dgvOverview.CurrentRow == null) { dgvDetails.DataSource = null; return; } if (cmbMapPlanRunNumber.SelectedItem == null) { lblStatus.Text = "Select the PlanRunNumber column mapping first."; return; } var col = cmbMapPlanRunNumber.SelectedItem.ToString(); var rawVal = GetSelectedOverviewCell(col); if (rawVal == null || rawVal == DBNull.Value) { lblStatus.Text = $"No value in column '{col}'."; dgvDetails.DataSource = null; return; } if (!long.TryParse(Convert.ToString(rawVal), out long planRunNumber)) { lblStatus.Text = $"Selected PlanRunNumber ('{col}') is not numeric."; dgvDetails.DataSource = null; return; } ToggleUi(false); lblStatus.Text = $"Loading details for PlanRunNumber {planRunNumber}..."; try { var table = await ExecuteQueryAsync(_detailsSql, cmd => cmd.Parameters.Add(new SqliteParameter(":PlanRunNumber", planRunNumber))); dgvDetails.DataSource = table; lblStatus.Text = $"Loaded {table.Rows.Count} detail rows for PlanRunNumber {planRunNumber}."; } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Error Loading Details", MessageBoxButtons.OK, MessageBoxIcon.Error); lblStatus.Text = "Failed to load details."; } finally { ToggleUi(true); UpdateUiState(); } }
        private async Task LoadStatsAllAsync() { if (!File.Exists(_dbPath)) return; ToggleUi(false); lblStatus.Text = "Loading stats (all runs)..."; try { var dt = await ExecuteQueryAsync(_statsAllSql, null); dgvStatsAll.DataSource = dt; lblStatus.Text = $"Loaded {dt.Rows.Count} rows."; } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Error Loading Stats (All)", MessageBoxButtons.OK, MessageBoxIcon.Error); lblStatus.Text = "Failed to load stats (all)."; } finally { ToggleUi(true); UpdateUiState(); } }
        private async Task LoadStatsBySerialAsync(string serial) { if (!File.Exists(_dbPath)) return; ToggleUi(false); lblStatus.Text = $"Loading stats (serial {serial})..."; try { var dt = await ExecuteQueryAsync(_statsBySerialSql, cmd => cmd.Parameters.Add(new SqliteParameter(":serNum", serial))); dgvStatsSerial.DataSource = dt; lblStatus.Text = $"Loaded {dt.Rows.Count} rows."; } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Error Loading Stats (By Serial)", MessageBoxButtons.OK, MessageBoxIcon.Error); lblStatus.Text = "Failed to load stats (serial)."; } finally { ToggleUi(true); UpdateUiState(); } }
        private void ExportDataTableToCsv(DataTable table, string path) { using var writer = new StreamWriter(path, false, Encoding.UTF8); for (int i = 0; i < table.Columns.Count; i++) { if (i > 0) writer.Write(","); writer.Write('"' + table.Columns[i].ColumnName.Replace("\"", "\"\"") + '"'); } writer.WriteLine(); foreach (DataRow row in table.Rows) { for (int i = 0; i < table.Columns.Count; i++) { if (i > 0) writer.Write(","); var val = row[i]?.ToString() ?? string.Empty; val = val.Replace("\"", "\"\""); writer.Write('"' + val + '"'); } writer.WriteLine(); } }
        private void ExportDataTableToExcel(DataTable table, string path, string sheetName) 
        { 
            using var wb = new XLWorkbook(); 
            var ws = wb.Worksheets.Add(string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName); 
            for (int c = 0; c < table.Columns.Count; c++) 
                ws.Cell(1, c + 1).Value = table.Columns[c].ColumnName; 
            for (int r = 0; r < table.Rows.Count; r++) 
                for (int c = 0; c < table.Columns.Count; c++)
                    //var value = table.Rows[r][c];
                    //ws.Cell(r + 2, c + 1).SetValue(value?.ToString() ?? string.Empty);
                    ws.Cell(r + 2, c + 1).SetValue(table.Rows[r][c]?.ToString() ?? string.Empty); 
                    //ws.Cell(r + 2, c + 1).Value = table.Rows[r][c]; 
                    ws.Columns().AdjustToContents(); wb.SaveAs(path); }
        private void Export(DataTable dt, string defaultFile) { if (dt == null || dt.Rows.Count == 0) return; using var sfd = new SaveFileDialog { Title = "Export", Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All files (*.*)|*.*", FileName = defaultFile }; if (sfd.ShowDialog(this) == DialogResult.OK) { try { var ext = Path.GetExtension(sfd.FileName).ToLowerInvariant(); if (ext == ".xlsx") ExportDataTableToExcel(dt, sfd.FileName, Path.GetFileNameWithoutExtension(defaultFile)); else ExportDataTableToCsv(dt, sfd.FileName); MessageBox.Show(this, "Exported.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error); } } }
        private void ToggleUi(bool enabled) { btnOpenDb.Enabled = enabled; btnLoad.Enabled = enabled && File.Exists(_dbPath); dgvOverview.Enabled = enabled; dgvDetails.Enabled = enabled; dgvStatsAll.Enabled = enabled; dgvStatsSerial.Enabled = enabled; grpFilters.Enabled = enabled; grpMapping.Enabled = enabled; grpStats.Enabled = enabled; btnExportOverview.Enabled = enabled && dgvOverview.DataSource is DataTable dtO && dtO.Rows.Count > 0; btnExportDetails.Enabled = enabled && dgvDetails.DataSource is DataTable dtD && dtD.Rows.Count > 0; btnExportStatsAll.Enabled = enabled && dgvStatsAll.DataSource is DataTable dtA && dtA.Rows.Count > 0; btnExportStatsSerial.Enabled = enabled && dgvStatsSerial.DataSource is DataTable dtS && dtS.Rows.Count > 0; }
        private async void btnLoad_Click(object sender, EventArgs e) => await LoadOverviewAsync(); private async void dgvOverview_SelectionChanged(object sender, EventArgs e) { if (dgvOverview.Focused || dgvOverview.IsHandleCreated) await LoadSelectedDetailsAsync(); }
        private async void btnApplyFilters_Click(object sender, EventArgs e) => await LoadOverviewAsync(); private async void btnLoadStatsAll_Click(object sender, EventArgs e) => await LoadStatsAllAsync(); private async void btnLoadStatsSerial_Click(object sender, EventArgs e) { var s = txtSerial.Text?.Trim(); if (!string.IsNullOrEmpty(s)) await LoadStatsBySerialAsync(s); else MessageBox.Show(this, "Enter a serial number.", "Missing Serial", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        private void btnOpenDb_Click(object sender, EventArgs e) { using var ofd = new OpenFileDialog { Title = "Open Test Results DB", Filter = "Test Results DB (*.db;*.TapResults)|*.db;*.TapResults|All files (*.*)|*.*", CheckFileExists = true, Multiselect = false }; if (ofd.ShowDialog(this) == DialogResult.OK) { _dbPath = ofd.FileName; txtDbPath.Text = _dbPath; SaveSettings(); UpdateUiState(); } }
        private void btnExportOverview_Click(object sender, EventArgs e) { if (dgvOverview.DataSource is DataTable dt) Export(dt, "overview.xlsx"); }
        private void btnExportDetails_Click(object sender, EventArgs e) { if (dgvDetails.DataSource is DataTable dt) Export(dt, "details.xlsx"); }
        private void btnExportStatsAll_Click(object sender, EventArgs e) { if (dgvStatsAll.DataSource is DataTable dt) Export(dt, "stats_all.xlsx"); }
        private void btnExportStatsSerial_Click(object sender, EventArgs e) { if (dgvStatsSerial.DataSource is DataTable dt) Export(dt, "stats_serial.xlsx"); }
    }
}