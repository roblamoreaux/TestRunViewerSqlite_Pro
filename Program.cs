using System;
using System.Windows.Forms;

namespace TestRunViewerSqlite
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Optional first argument: path to a .TapResults / .db file to open on startup
            string? dbPath = args.Length > 0 ? args[0] : null;
            if (dbPath != null && !File.Exists(dbPath))
            {
                MessageBox.Show(
                    $"File not found:\n{dbPath}",
                    "Test Run Viewer (SQLite)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                dbPath = null;
            }

            Application.Run(new MainForm(dbPath));
        }
    }
}
