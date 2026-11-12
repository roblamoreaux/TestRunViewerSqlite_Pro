using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace TestRunViewerSqlite
{
    // Startup splash screen: just the rounded TAP_Explorer icon, centered on screen while the main window
    // opens. It's a layered window, so everything outside the icon (with its anti-aliased edge) is see-through.
    internal sealed class SplashForm : Form
    {
        private readonly Bitmap _image;

        public SplashForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;

            var size = LogicalToDeviceUnits(new Size(360, 360));
            _image = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
            using (var stream = typeof(SplashForm).Assembly.GetManifestResourceStream("TAP_Explorer_splash.png")!)
            using (var source = Image.FromStream(stream))
            using (var g = Graphics.FromImage(_image))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(source, new Rectangle(Point.Empty, size));
            }

            StartPosition = FormStartPosition.Manual;
            var area = Screen.FromPoint(Cursor.Position).WorkingArea;
            Bounds = new Rectangle(area.Left + (area.Width - size.Width) / 2, area.Top + (area.Height - size.Height) / 2, size.Width, size.Height);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_LAYERED;
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // Give Windows the image with its alpha channel; it draws the window from that.
            IntPtr screenDc = GetDC(IntPtr.Zero);
            IntPtr memoryDc = CreateCompatibleDC(screenDc);
            IntPtr bitmap = _image.GetHbitmap(Color.FromArgb(0)); // premultiplied alpha
            IntPtr previous = SelectObject(memoryDc, bitmap);
            try
            {
                var size = new NativeSize { Width = _image.Width, Height = _image.Height };
                var source = new NativePoint();
                var position = new NativePoint { X = Left, Y = Top };
                var blend = new BlendFunction { SourceConstantAlpha = 255, AlphaFormat = AC_SRC_ALPHA };
                UpdateLayeredWindow(Handle, screenDc, ref position, ref size, memoryDc, ref source, 0, ref blend, ULW_ALPHA);
            }
            finally
            {
                SelectObject(memoryDc, previous);
                DeleteObject(bitmap);
                DeleteDC(memoryDc);
                ReleaseDC(IntPtr.Zero, screenDc);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _image.Dispose();
            base.Dispose(disposing);
        }

        private const int WS_EX_LAYERED = 0x80000;
        private const int ULW_ALPHA = 2;
        private const byte AC_SRC_ALPHA = 1;

        [StructLayout(LayoutKind.Sequential)] private struct NativePoint { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct NativeSize { public int Width, Height; }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct BlendFunction { public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat; }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref NativePoint pptDst, ref NativeSize psize,
            IntPtr hdcSrc, ref NativePoint pptSrc, int crKey, ref BlendFunction pblend, int dwFlags);
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr hDC);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr hObject);
    }
}
