using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MOfis.Common;

namespace MOfis.Writer;

/// Sayfa boyutu ve kenar bosluklari (piksel, 1/96 inc).
public sealed class PageSettings
{
    public double Width { get; set; } = Units.CmToPx(21.0);
    public double Height { get; set; } = Units.CmToPx(29.7);
    public double MarginLeft { get; set; } = Units.CmToPx(2.5);
    public double MarginTop { get; set; } = Units.CmToPx(2.5);
    public double MarginRight { get; set; } = Units.CmToPx(2.5);
    public double MarginBottom { get; set; } = Units.CmToPx(2.5);

    public double ContentWidth => Math.Max(50, Width - MarginLeft - MarginRight);
    public bool IsLandscape => Width > Height;

    public PageSettings Clone() => (PageSettings)MemberwiseClone();

    public void SetMargins(double cm)
    {
        MarginLeft = MarginRight = MarginTop = MarginBottom = Units.CmToPx(cm);
    }
}

/// Belgeye eklenen resmin ozgun baytlari; Image (veya desteklenmeyen bicimde yer tutucu) Tag'inde durur.
public sealed record ImageData(byte[] Bytes, string Extension);

/// Paragraf stilleri (Normal, Baslik 1-3, Konu Basligi). Id'ler DOCX stil kimlikleri ile ayni.
public sealed record ParaStyleDef(string Id, string DocxName, string DisplayName, double SizePt, bool Bold,
    Color Color, double BeforePt, double AfterPt, int OutlineLevel);

public static class StyleCatalog
{
    public const string DefaultFont = "Calibri";
    public const double DefaultSizePt = 11;

    public static readonly ParaStyleDef Normal = new("Normal", "Normal", "Normal", 11, false, Colors.Black, 0, 8, -1);
    public static readonly ParaStyleDef Title = new("Title", "Title", "Konu Başlığı", 26, false, Color.FromRgb(0x11, 0x18, 0x27), 0, 6, -1);
    public static readonly ParaStyleDef Heading1 = new("Heading1", "heading 1", "Başlık 1", 16, true, Color.FromRgb(0x1E, 0x3A, 0x8A), 12, 4, 0);
    public static readonly ParaStyleDef Heading2 = new("Heading2", "heading 2", "Başlık 2", 13, true, Color.FromRgb(0x1E, 0x40, 0xAF), 10, 4, 1);
    public static readonly ParaStyleDef Heading3 = new("Heading3", "heading 3", "Başlık 3", 12, true, Color.FromRgb(0x37, 0x41, 0x51), 8, 4, 2);

    public static readonly ParaStyleDef[] All = [Normal, Title, Heading1, Heading2, Heading3];

    public static ParaStyleDef? ById(string? id) => All.FirstOrDefault(s => s.Id == id);

    /// DOCX stil adindan (her dilde Ingilizce yazilir: "heading 1", "title") bizim stile.
    public static ParaStyleDef? ByDocxName(string? name)
    {
        if (string.IsNullOrEmpty(name)) return null;
        return All.FirstOrDefault(s => string.Equals(s.DocxName, name, StringComparison.OrdinalIgnoreCase));
    }
}

public static class DocHelpers
{
    public static FlowDocument CreateBlank()
    {
        var doc = new FlowDocument();
        Prepare(doc);
        doc.Blocks.Add(new Paragraph());
        return doc;
    }

    /// Her belgeye (yeni veya acilan) uygulanan ortak ayarlar.
    public static void Prepare(FlowDocument doc)
    {
        doc.FontFamily = new FontFamily(StyleCatalog.DefaultFont);
        doc.FontSize = Units.PtToPx(StyleCatalog.DefaultSizePt);
        doc.Foreground = Brushes.Black;
        doc.Background = Brushes.Transparent;
        doc.PagePadding = new Thickness(0);
        doc.TextAlignment = TextAlignment.Left;
        // Yeni paragraflar Word'deki "Normal" gibi alttan 8 pt bosluklu baslar.
        var style = new Style(typeof(Paragraph));
        style.Setters.Add(new Setter(Block.MarginProperty, new Thickness(0, 0, 0, Units.PtToPx(8))));
        doc.Resources[typeof(Paragraph)] = style;
        var tableStyle = new Style(typeof(Table));
        tableStyle.Setters.Add(new Setter(Table.CellSpacingProperty, 0.0));
        tableStyle.Setters.Add(new Setter(Block.MarginProperty, new Thickness(0, 0, 0, Units.PtToPx(8))));
        doc.Resources[typeof(Table)] = tableStyle;
    }

    /// Paragrafa stil uygular (yazi boyutu, kalinlik, renk, bosluklar) ve stil kimligini Tag'e yazar.
    public static void ApplyStyle(Paragraph p, ParaStyleDef s)
    {
        p.Tag = s.Id == StyleCatalog.Normal.Id ? null : s.Id;
        var size = Units.PtToPx(s.SizePt);
        var weight = s.Bold ? FontWeights.Bold : FontWeights.Normal;
        var brush = new SolidColorBrush(s.Color);
        brush.Freeze();
        var range = new TextRange(p.ContentStart, p.ContentEnd);
        range.ApplyPropertyValue(TextElement.FontSizeProperty, size);
        range.ApplyPropertyValue(TextElement.FontWeightProperty, weight);
        range.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        p.FontSize = size;
        p.FontWeight = weight;
        p.Foreground = brush;
        var m = p.Margin;
        p.Margin = new Thickness(double.IsNaN(m.Left) ? 0 : m.Left, Units.PtToPx(s.BeforePt), double.IsNaN(m.Right) ? 0 : m.Right, Units.PtToPx(s.AfterPt));
        p.KeepWithNext = s.OutlineLevel >= 0;
    }

    /// Sayfa sonu olan paragraflari editorde ince bir cizgiyle gosterir (yazdirmada gorunmez).
    public static void MarkPageBreak(Paragraph p, bool on)
    {
        p.BreakPageBefore = on;
        if (on)
        {
            p.BorderBrush = new SolidColorBrush(Color.FromRgb(0xB8, 0xBB, 0xC8));
            p.BorderThickness = new Thickness(0, 1, 0, 0);
            p.Padding = new Thickness(0, 10, 0, 0);
        }
        else
        {
            p.ClearValue(Block.BorderBrushProperty);
            p.ClearValue(Block.BorderThicknessProperty);
            p.ClearValue(Block.PaddingProperty);
        }
    }

    public static void ShowPageBreaks(FlowDocument doc)
    {
        foreach (var p in AllParagraphs(doc.Blocks))
            if (p.BreakPageBefore) MarkPageBreak(p, true);
    }

    public static void HidePageBreakMarks(FlowDocument doc)
    {
        foreach (var p in AllParagraphs(doc.Blocks))
        {
            if (!p.BreakPageBefore) continue;
            p.ClearValue(Block.BorderBrushProperty);
            p.ClearValue(Block.BorderThicknessProperty);
            p.ClearValue(Block.PaddingProperty);
        }
    }

    public static IEnumerable<Paragraph> AllParagraphs(BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            switch (b)
            {
                case Paragraph p:
                    yield return p;
                    break;
                case List l:
                    foreach (var li in l.ListItems)
                        foreach (var p in AllParagraphs(li.Blocks)) yield return p;
                    break;
                case Table t:
                    foreach (var g in t.RowGroups)
                        foreach (var r in g.Rows)
                            foreach (var c in r.Cells)
                                foreach (var p in AllParagraphs(c.Blocks)) yield return p;
                    break;
                case Section s:
                    foreach (var p in AllParagraphs(s.Blocks)) yield return p;
                    break;
            }
        }
    }

    /// Resim baytlarindan editorde gosterilecek ogeyi uretir. Desteklenmeyen bicimlerde (EMF/WMF)
    /// gri bir yer tutucu gosterir ama baytlari saklar; kaydederken resim kaybolmaz.
    public static FrameworkElement CreateImageElement(ImageData data, double? widthPx, double? heightPx, double maxWidth)
    {
        try
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.StreamSource = new MemoryStream(data.Bytes);
            bmp.EndInit();
            bmp.Freeze();

            double w = widthPx ?? bmp.Width;
            double h = heightPx ?? bmp.Height;
            if (w > maxWidth && maxWidth > 0)
            {
                h = h * maxWidth / w;
                w = maxWidth;
            }
            return new Image { Source = bmp, Width = w, Height = h, Stretch = Stretch.Fill, Tag = data };
        }
        catch
        {
            double w = Math.Min(widthPx ?? 200, maxWidth > 0 ? maxWidth : 200);
            double h = heightPx ?? 120;
            return new Border
            {
                Width = w,
                Height = h,
                Background = new SolidColorBrush(Color.FromRgb(0xEE, 0xEF, 0xF3)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xC9, 0xCA, 0xD2)),
                BorderThickness = new Thickness(1),
                Tag = data,
                Child = new TextBlock
                {
                    Text = "Bu resim biçimi gösterilemiyor\n(kaydederken korunur)",
                    TextAlignment = TextAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = Brushes.DimGray,
                    FontSize = 11,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                },
            };
        }
    }

    public static Color? ParseHex(string? hex)
    {
        if (string.IsNullOrEmpty(hex) || hex.Length != 6 || hex.Equals("auto", StringComparison.OrdinalIgnoreCase)) return null;
        try
        {
            return Color.FromRgb(Convert.ToByte(hex[..2], 16), Convert.ToByte(hex.Substring(2, 2), 16), Convert.ToByte(hex.Substring(4, 2), 16));
        }
        catch
        {
            return null;
        }
    }

    public static string ToHex(Color c) => $"{c.R:X2}{c.G:X2}{c.B:X2}";
}
