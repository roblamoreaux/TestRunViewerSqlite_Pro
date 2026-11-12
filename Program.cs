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
                    "Test Run Viewer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                dbPath = null;
            }

            // Splash screen until the main window is up, shown for at least 1.5 seconds
            using var splash = new SplashForm();
            splash.Show();
            splash.Update();
            var splashTime = System.Diagnostics.Stopwatch.StartNew();

            var form = new MainForm(dbPath);
            form.Shown += async (_, _) =>
            {
                var remaining = TimeSpan.FromSeconds(1.5) - splashTime.Elapsed;
                if (remaining > TimeSpan.Zero)
                    await Task.Delay(remaining);
                splash.Close();
                form.Activate();
            };
            Application.Run(form);
        }
    }
}
