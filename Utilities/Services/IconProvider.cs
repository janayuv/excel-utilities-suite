using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace utilities.Services
{
    /// <summary>
    /// Resolves a command's <c>ImageId</c> to a ribbon image. Icons ship as embedded PNG
    /// resources in four sets: Icons (light 32 px), Icons16 (light 16 px),
    /// IconsDark (dark 32 px), IconsDark16 (dark 16 px).
    /// Dark-mode detection reads the Office 16 UI Theme registry key; falls back to
    /// the Windows app colour-mode preference. Result is cached; call
    /// <see cref="InvalidateCache"/> when the Office theme changes.
    /// </summary>
    public static class IconProvider
    {
        // ── Caches ────────────────────────────────────────────────────────────
        private static readonly Dictionary<string, stdole.IPictureDisp> _pictureCache =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, Bitmap> _bitmapCache =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);

        // Cached resource name list so we don't call GetManifestResourceNames() repeatedly.
        private static string[] _resourceNames;

        // ── Dark-mode detection ───────────────────────────────────────────────

        /// <summary>
        /// Returns true when Office / Windows is in dark mode.
        /// Checks the Office 16 "UI Theme" registry value first (3 = Black / Dark),
        /// then falls back to the Windows "AppsUseLightTheme" value.
        /// </summary>
        public static bool IsDarkMode
        {
            get
            {
                try
                {
                    // Office theme: 0=Colorful, 1=Dark Gray, 2=White, 3=Black (dark), 4=Use system
                    object officeTheme = Registry.GetValue(
                        @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common",
                        "UI Theme", null);
                    if (officeTheme is int ot)
                    {
                        if (ot == 3) return true;   // Black = dark
                        if (ot == 1) return false;  // Dark Gray still uses light icons
                        if (ot != 4) return false;  // 0/2 = light themes
                        // ot == 4 -> "Use system setting" -> fall through to Windows check
                    }
                }
                catch { }

                try
                {
                    // Windows: 0 = dark, 1 = light
                    object winLight = Registry.GetValue(
                        @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                        "AppsUseLightTheme", 1);
                    if (winLight is int wl) return wl == 0;
                }
                catch { }

                return false;
            }
        }

        /// <summary>Clear the icon caches (call after an Office theme change).</summary>
        public static void InvalidateCache()
        {
            _pictureCache.Clear();
            _bitmapCache.Clear();
        }

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>Ribbon getImage result. Returns null when no icon is found.</summary>
        public static stdole.IPictureDisp GetPicture(string imageId)
        {
            if (string.IsNullOrEmpty(imageId)) return null;

            // Include dark-mode flag in cache key so theme switches work correctly.
            string cacheKey = imageId + (IsDarkMode ? ":dark" : ":light");
            stdole.IPictureDisp cached;
            if (_pictureCache.TryGetValue(cacheKey, out cached)) return cached;

            Bitmap bmp = GetBitmap(imageId);
            stdole.IPictureDisp pic = bmp != null ? PictureConverter.ToPictureDisp(bmp) : null;
            _pictureCache[cacheKey] = pic;
            return pic;
        }

        public static Bitmap GetBitmap(string imageId)
        {
            if (string.IsNullOrEmpty(imageId)) return null;

            string cacheKey = imageId + (IsDarkMode ? ":dark" : ":light");
            Bitmap cached;
            if (_bitmapCache.TryGetValue(cacheKey, out cached)) return cached;

            Bitmap bmp = LoadIcon(imageId);
            _bitmapCache[cacheKey] = bmp;
            return bmp;
        }

        // ── Private helpers ───────────────────────────────────────────────────

        private static Bitmap LoadIcon(string imageId)
        {
            string fileName = imageId + "Icon.png";
            bool dark = IsDarkMode;

            if (dark)
            {
                // Prefer dark icon, fall back to light.
                Bitmap bmp = LoadEmbeddedFromPrefix("IconsDark.", fileName)
                          ?? LoadEmbeddedFromPrefix("Icons.",     fileName);
                if (bmp != null) return bmp;
            }
            else
            {
                Bitmap bmp = LoadEmbeddedFromPrefix("Icons.", fileName);
                if (bmp != null) return bmp;
            }

            // Last resort: suffix search (handles any legacy resource naming).
            return LoadEmbeddedSuffix(fileName) ?? LoadEmbeddedSuffix(imageId + ".png");
        }

        private static Bitmap LoadEmbeddedFromPrefix(string prefix, string fileName)
        {
            string fullName = prefix + fileName;
            EnsureResourceNames();
            foreach (string name in _resourceNames)
            {
                if (string.Equals(name, fullName, StringComparison.OrdinalIgnoreCase))
                    return LoadStream(name);
            }
            return null;
        }

        private static Bitmap LoadEmbeddedSuffix(string fileSuffix)
        {
            EnsureResourceNames();
            string dotSuffix = "." + fileSuffix;
            foreach (string name in _resourceNames)
            {
                if (name.EndsWith(dotSuffix,  StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith(fileSuffix, StringComparison.OrdinalIgnoreCase))
                    return LoadStream(name);
            }
            return null;
        }

        private static Bitmap LoadStream(string resourceName)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream s = asm.GetManifestResourceStream(resourceName))
                {
                    if (s == null) return null;
                    return new Bitmap(s);
                }
            }
            catch { return null; }
        }

        private static void EnsureResourceNames()
        {
            if (_resourceNames == null)
                _resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
        }

        /// <summary>Converts a managed Bitmap to the COM IPictureDisp the ribbon API requires.</summary>
        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base("00000000-0000-0000-0000-000000000000") { }

            public static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
    }
}
