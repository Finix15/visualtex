using System.Drawing;
using System.Windows.Forms;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class RibbonIconProvider
{
    private static readonly object Gate = new();
    private static readonly Dictionary<string, IconResource> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    internal static object? GetImage(string? key)
    {
        if (key is null) return null;
        var requiredKey = key.Trim();
        if (requiredKey.Length == 0) return null;
        if (!RibbonVectorIconRenderer.Keys.Contains(requiredKey)) return null;

        lock (Gate)
        {
            if (Cache.TryGetValue(requiredKey, out var cached))
                return cached.PictureDisp;

            var bitmap = RibbonVectorIconRenderer.Create(requiredKey);
            var pictureDisp = PictureDispHost.ToPictureDisp(bitmap);
            Cache[requiredKey] = new IconResource(bitmap, pictureDisp);
            return pictureDisp;
        }
    }

    private sealed class IconResource
    {
        internal IconResource(Bitmap bitmap, object pictureDisp)
        {
            Bitmap = bitmap;
            PictureDisp = pictureDisp;
        }

        // Office's IPictureDisp wrapper continues to reference the GDI image.
        internal Bitmap Bitmap { get; }
        internal object PictureDisp { get; }
    }

    private sealed class PictureDispHost : AxHost
    {
        private PictureDispHost() : base(string.Empty) { }

        internal static object ToPictureDisp(Image image) =>
            GetIPictureDispFromPicture(image);
    }
}
