using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace XLify
{
    internal static class DpiHelper
    {
        // Windows 10 (1607+) per-monitor DPI
        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        internal static double GetScaleForWindow(IntPtr hwnd)
        {
            try
            {
                if (hwnd != IntPtr.Zero)
                {
                    try
                    {
                        uint dpi = GetDpiForWindow(hwnd);
                        if (dpi >= 96) return dpi / 96.0;
                    }
                    catch
                    {
                        // API not available; fall back below
                    }
                }

                using (var g = Graphics.FromHwnd(hwnd))
                {
                    return g.DpiX / 96.0;
                }
            }
            catch
            {
                return 1.0; // default scale
            }
        }
    }
}

