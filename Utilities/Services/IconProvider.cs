using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace utilities.Services
{
    /// <summary>
    /// Resolves a command's <c>ImageId</c> to a ribbon image. Icons ship as embedded PNG
    /// resources (named "{ImageId}Icon.png"); the bitmaps are cached and converted to the
    /// IPictureDisp the Office ribbon getImage callback expects.
    /// </summary>
    public static class IconProvider
    {
        private static readonly Dictionary<string, stdole.IPictureDisp> _pictureCache =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, Bitmap> _bitmapCache =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Ribbon getImage result. Returns null when no icon is found (button shows text only).</summary>
        public static stdole.IPictureDisp GetPicture(string imageId)
        {
            if (string.IsNullOrEmpty(imageId)) return null;

            stdole.IPictureDisp cached;
            if (_pictureCache.TryGetValue(imageId, out cached)) return cached;

            Bitmap bmp = GetBitmap(imageId);
            stdole.IPictureDisp pic = bmp != null ? PictureConverter.ToPictureDisp(bmp) : null;
            _pictureCache[imageId] = pic;
            return pic;
        }

        public static Bitmap GetBitmap(string imageId)
        {
            if (string.IsNullOrEmpty(imageId)) return null;

            Bitmap cached;
            if (_bitmapCache.TryGetValue(imageId, out cached)) return cached;

            Bitmap bmp = LoadEmbedded(imageId + "Icon.png") ?? LoadEmbedded(imageId + ".png");
            _bitmapCache[imageId] = bmp;
            return bmp;
        }

        private static Bitmap LoadEmbedded(string fileSuffix)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (name.EndsWith("." + fileSuffix, StringComparison.OrdinalIgnoreCase) ||
                        name.EndsWith(fileSuffix, StringComparison.OrdinalIgnoreCase))
                    {
                        using (Stream s = asm.GetManifestResourceStream(name))
                        {
                            if (s == null) continue;
                            return new Bitmap(s);
                        }
                    }
                }
            }
            catch
            {
                // Missing icon is non-fatal — the button simply renders without an image.
            }
            return null;
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
