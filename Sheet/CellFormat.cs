using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ClosedXML.Excel;

namespace MOfis.Sheet;

/// Hucrenin ekranda gorunecek metni (Turkce bicimler) ve renk donusumleri.
public static class SheetDisplay
{
    private static readonly CultureInfo Tr = new("tr-TR");

    public static string Text(IXLCell cell)
    {
        XLCellValue v;
        try
        {
            v = cell.Value;
        }
        catch
        {
            return cell.HasFormula ? "#AD?" : "";
        }

        switch (v.Type)
        {
            case XLDataType.Blank: return "";
            case XLDataType.Text: return v.GetText();
            case XLDataType.Boolean: return v.GetBoolean() ? "DOĞRU" : "YANLIŞ";
            case XLDataType.Error: return FormulaLocale.ErrorToTurkish(CellInput.ErrorText(v.GetError()));
        }

        var nf = cell.Style.NumberFormat;
        // Yerel ayara bagli yerlesik tarih bicimleri (14 = kisa tarih, 22 = tarih saat): Turkce goster.
        if (nf.NumberFormatId == 14 && string.IsNullOrEmpty(nf.Format) && v.IsDateTime) return v.GetDateTime().ToString("dd.MM.yyyy", Tr);
        if (nf.NumberFormatId == 22 && string.IsNullOrEmpty(nf.Format) && v.IsDateTime) return v.GetDateTime().ToString("dd.MM.yyyy HH:mm", Tr);

        try
        {
            var s = cell.GetFormattedString(Tr);
            if (v.IsNumber && s.Length > 0 && IsGeneral(nf)) return FormatGeneral(v.GetNumber());
            return s;
        }
        catch
        {
            return v.IsNumber ? FormatGeneral(v.GetNumber()) : v.ToString(Tr);
        }
    }

    public static bool IsGeneral(IXLNumberFormat nf) =>
        (nf.NumberFormatId == 0 && string.IsNullOrEmpty(nf.Format)) || string.Equals(nf.Format, "General", StringComparison.OrdinalIgnoreCase);

    /// "Genel" bicim: Excel gibi en fazla 11 anlamli basamak.
    public static string FormatGeneral(double n)
    {
        if (n == 0) return "0";
        var abs = Math.Abs(n);
        if (abs >= 1e11 || abs < 1e-9) return n.ToString("0.#####E+00", Tr);
        return Math.Round(n, Math.Max(0, 10 - (int)Math.Floor(Math.Log10(abs)))).ToString("0.##########", Tr);
    }

    /// "Genel" bicimli sayi sutuna sigmiyorsa Excel gibi ondaliklari azaltarak sigdirmayi dener.
    public static string? FitGeneral(double n, Func<string, double> measure, double width)
    {
        var full = FormatGeneral(n);
        if (measure(full) <= width) return full;
        if (Math.Abs(n) >= 1e11) return null;
        int intDigits = Math.Abs(n) < 1 ? 1 : (int)Math.Floor(Math.Log10(Math.Abs(n))) + 1;
        for (int dec = Math.Max(0, 10 - intDigits); dec >= 0; dec--)
        {
            var s = Math.Round(n, dec).ToString("0." + new string('#', Math.Max(1, dec)), Tr);
            if (measure(s) <= width) return s;
        }
        return null;
    }

    /// Hizalama belirtilmemisse Excel gibi: sayi/tarih saga, metin sola, mantiksal/hata ortaya.
    public static TextAlignment Alignment(IXLCell cell, XLAlignmentHorizontalValues h, XLDataType type) => h switch
    {
        XLAlignmentHorizontalValues.Left => TextAlignment.Left,
        XLAlignmentHorizontalValues.Right => TextAlignment.Right,
        XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.CenterContinuous => TextAlignment.Center,
        XLAlignmentHorizontalValues.Justify or XLAlignmentHorizontalValues.Distributed => TextAlignment.Left,
        _ => type switch
        {
            XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan => TextAlignment.Right,
            XLDataType.Boolean or XLDataType.Error => TextAlignment.Center,
            _ => TextAlignment.Left,
        },
    };

    /// XLColor -> WPF rengi. Tema renkleri calisma kitabinin temasindan (ton/tint uygulanarak) cozulur.
    public static Color? ToWpf(XLColor? c, IXLWorkbook wb)
    {
        if (c == null || !c.HasValue) return null;
        try
        {
            switch (c.ColorType)
            {
                case XLColorType.Theme:
                {
                    var baseColor = ThemeColor(wb.Theme, c.ThemeColor);
                    if (baseColor == null) return null;
                    return ApplyTint(baseColor.Value, c.ThemeTint);
                }
                default:
                {
                    var d = c.Color;
                    if (d.A == 0 && d.R == 0 && d.G == 0 && d.B == 0) return null;
                    return Color.FromArgb(d.A == 0 ? (byte)255 : d.A, d.R, d.G, d.B);
                }
            }
        }
        catch
        {
            return null;
        }
    }

    private static Color? ThemeColor(IXLTheme t, XLThemeColor tc)
    {
        XLColor x = tc switch
        {
            XLThemeColor.Background1 => t.Background1,
            XLThemeColor.Text1 => t.Text1,
            XLThemeColor.Background2 => t.Background2,
            XLThemeColor.Text2 => t.Text2,
            XLThemeColor.Accent1 => t.Accent1,
            XLThemeColor.Accent2 => t.Accent2,
            XLThemeColor.Accent3 => t.Accent3,
            XLThemeColor.Accent4 => t.Accent4,
            XLThemeColor.Accent5 => t.Accent5,
            XLThemeColor.Accent6 => t.Accent6,
            XLThemeColor.Hyperlink => t.Hyperlink,
            XLThemeColor.FollowedHyperlink => t.FollowedHyperlink,
            _ => t.Text1,
        };
        if (x.ColorType == XLColorType.Theme) return null;
        var d = x.Color;
        return Color.FromRgb(d.R, d.G, d.B);
    }

    private static Color ApplyTint(Color c, double tint)
    {
        if (Math.Abs(tint) < 1e-6) return c;
        byte T(byte v) => (byte)Math.Round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint);
        return Color.FromRgb(T(c.R), T(c.G), T(c.B));
    }

    private static readonly Dictionary<Color, SolidColorBrush> BrushCache = [];

    public static SolidColorBrush Brush(Color c)
    {
        if (!BrushCache.TryGetValue(c, out var b))
        {
            b = new SolidColorBrush(c);
            b.Freeze();
            BrushCache[c] = b;
        }
        return b;
    }
}
