using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MOfis.Common;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MOfis.Writer;

/// FlowDocument -> DOCX. Desteklenen: paragraf bicimi, yazi bicimi, basliklar, listeler,
/// tablolar (birlesik hucreler dahil), resimler, baglantilar, sayfa sonu, sayfa boyutu/kenar bosluklari.
public sealed class DocxExporter
{
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private readonly PageSettings _page;
    private MainDocumentPart _main = null!;
    private readonly List<W.AbstractNum> _abstracts = [];
    private readonly Dictionary<TextMarkerStyle, int> _abstractByMarker = [];
    private readonly List<W.NumberingInstance> _nums = [];
    private uint _drawingId = 1;

    private DocxExporter(PageSettings page) => _page = page;

    public static void Save(FlowDocument doc, PageSettings page, Stream stream) => new DocxExporter(page).Write(doc, stream);

    private void Write(FlowDocument doc, Stream stream)
    {
        using var pkg = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        _main = pkg.AddMainDocumentPart();
        var body = new W.Body();
        foreach (var el in ConvertBlocks(doc.Blocks)) body.Append(el);
        if (body.LastChild is W.Table) body.Append(new W.Paragraph());
        body.Append(SectionProperties());
        _main.Document = new W.Document(body);
        AddStyles();
        if (_nums.Count > 0) AddNumbering();
        _main.Document.Save();
    }

    // ---------------------------------------------------------------- bloklar

    private IEnumerable<OpenXmlElement> ConvertBlocks(IEnumerable<Block> blocks)
    {
        OpenXmlElement? last = null;
        foreach (var b in blocks)
        {
            foreach (var el in ConvertBlock(b))
            {
                // Arka arkaya iki tablo Word'de tek tabloya birlesir; araya bos paragraf koy.
                if (el is W.Table && last is W.Table) yield return new W.Paragraph();
                yield return el;
                last = el;
            }
        }
    }

    private IEnumerable<OpenXmlElement> ConvertBlock(Block b)
    {
        switch (b)
        {
            case Paragraph p:
                yield return ConvertParagraph(p, null, null);
                break;
            case List l:
                foreach (var el in ConvertList(l, 0)) yield return el;
                break;
            case Table t:
                yield return ConvertTable(t);
                break;
            case Section s:
                foreach (var el in ConvertBlocks(s.Blocks)) yield return el;
                break;
            case BlockUIContainer c when c.Child is FrameworkElement fe && fe.Tag is ImageData d:
                yield return new W.Paragraph(new W.Run(ImageDrawing(d, fe)));
                break;
        }
    }

    private IEnumerable<OpenXmlElement> ConvertList(List list, int level)
    {
        int numId = NewNum(list.MarkerStyle, level, list.StartIndex);
        foreach (var item in list.ListItems)
        {
            bool first = true;
            foreach (var b in item.Blocks)
            {
                if (b is Paragraph p)
                {
                    yield return first ? ConvertParagraph(p, (numId, level), null) : ConvertParagraph(p, null, level);
                    first = false;
                }
                else if (b is List nested)
                {
                    foreach (var el in ConvertList(nested, Math.Min(level + 1, 8))) yield return el;
                }
                else
                {
                    foreach (var el in ConvertBlock(b)) yield return el;
                }
            }
            if (first) yield return ConvertParagraph(new Paragraph(), (numId, level), null);
        }
    }

    private W.Paragraph ConvertParagraph(Paragraph p, (int numId, int level)? num, int? continuationLevel)
    {
        var wp = new W.Paragraph();
        var ppr = new W.ParagraphProperties();

        var style = StyleCatalog.ById(p.Tag as string);
        if (style != null && style != StyleCatalog.Normal) ppr.Append(new W.ParagraphStyleId { Val = style.Id });
        if (p.KeepWithNext) ppr.Append(new W.KeepNext());
        if (p.BreakPageBefore) ppr.Append(new W.PageBreakBefore());
        if (num is { } n)
            ppr.Append(new W.NumberingProperties(new W.NumberingLevelReference { Val = n.level }, new W.NumberingId { Val = n.numId }));

        var m = p.Margin;
        var spacing = new W.SpacingBetweenLines { Before = Tw(Nz(m.Top)), After = Tw(Nz(m.Bottom)) };
        if (!double.IsNaN(p.LineHeight))
        {
            double natural = p.FontFamily.LineSpacing * p.FontSize;
            double factor = natural > 0 ? p.LineHeight / natural : 1;
            spacing.Line = ((int)Math.Round(240 * factor)).ToString(Inv);
            spacing.LineRule = W.LineSpacingRuleValues.Auto;
        }
        ppr.Append(spacing);

        double left = Nz(m.Left), right = Nz(m.Right);
        if (continuationLevel is int cl) left += 48 * (cl + 1);
        if (num == null && (left > 0.5 || right > 0.5 || Math.Abs(p.TextIndent) > 0.5))
        {
            var ind = new W.Indentation();
            if (left > 0.5) ind.Left = Tw(left);
            if (right > 0.5) ind.Right = Tw(right);
            if (p.TextIndent > 0.5) ind.FirstLine = Tw(p.TextIndent);
            else if (p.TextIndent < -0.5) ind.Hanging = Tw(-p.TextIndent);
            ppr.Append(ind);
        }

        switch (p.TextAlignment)
        {
            case TextAlignment.Center: ppr.Append(new W.Justification { Val = W.JustificationValues.Center }); break;
            case TextAlignment.Right: ppr.Append(new W.Justification { Val = W.JustificationValues.Right }); break;
            case TextAlignment.Justify: ppr.Append(new W.Justification { Val = W.JustificationValues.Both }); break;
        }

        // Paragraf isaretinin bicimi: bos satirlarin yuksekligi korunsun.
        var mark = new W.ParagraphMarkRunProperties();
        foreach (var child in RunProps(p).ChildElements) mark.Append(child.CloneNode(true));
        ppr.Append(mark);

        wp.Append(ppr);
        AddInlines(p.Inlines, wp);
        return wp;
    }

    // ---------------------------------------------------------------- satir ici ogeler

    private void AddInlines(InlineCollection inlines, OpenXmlElement parent)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case Run r:
                    parent.Append(ConvertRun(r));
                    break;
                case LineBreak lb:
                    parent.Append(new W.Run(RunProps(lb), new W.Break()));
                    break;
                case Hyperlink h when h.NavigateUri is { IsAbsoluteUri: true } uri:
                    var rel = _main.AddHyperlinkRelationship(uri, true);
                    var wh = new W.Hyperlink { Id = rel.Id };
                    AddInlines(h.Inlines, wh);
                    parent.Append(wh);
                    break;
                case Span s:
                    AddInlines(s.Inlines, parent);
                    break;
                case InlineUIContainer c when c.Child is FrameworkElement fe && fe.Tag is ImageData d:
                    parent.Append(new W.Run(ImageDrawing(d, fe)));
                    break;
            }
        }
    }

    private W.Run ConvertRun(Run r)
    {
        var wr = new W.Run(RunProps(r));
        var sb = new StringBuilder();
        void Flush()
        {
            if (sb.Length == 0) return;
            wr.Append(new W.Text(sb.ToString()) { Space = SpaceProcessingModeValues.Preserve });
            sb.Clear();
        }
        foreach (var ch in r.Text)
        {
            switch (ch)
            {
                case '\t': Flush(); wr.Append(new W.TabChar()); break;
                case '\n': Flush(); wr.Append(new W.Break()); break;
                case '\r': break;
                default:
                    // XML'de gecersiz kontrol karakterlerini atla.
                    if (ch >= 0x20 || ch == '\u0009') sb.Append(ch);
                    break;
            }
        }
        Flush();
        return wr;
    }

    private static W.RunProperties RunProps(TextElement e)
    {
        var rp = new W.RunProperties();
        var font = FirstFamily(e.FontFamily);
        rp.Append(new W.RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font });
        if (e.FontWeight.ToOpenTypeWeight() >= 600) rp.Append(new W.Bold());
        if (e.FontStyle == FontStyles.Italic || e.FontStyle == FontStyles.Oblique) rp.Append(new W.Italic());

        var (underline, strike) = Decorations(e);
        if (strike) rp.Append(new W.Strike());
        if (e.Foreground is SolidColorBrush fg && fg.Color.A > 0 && fg.Color != Colors.Black)
            rp.Append(new W.Color { Val = DocHelpers.ToHex(fg.Color) });
        var halfPoints = Math.Clamp((int)Math.Round(Units.PxToPt(e.FontSize) * 2), 2, 3276).ToString(Inv);
        rp.Append(new W.FontSize { Val = halfPoints });
        rp.Append(new W.FontSizeComplexScript { Val = halfPoints });
        if (underline) rp.Append(new W.Underline { Val = W.UnderlineValues.Single });
        if (BackgroundOf(e) is Color bg)
            rp.Append(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = DocHelpers.ToHex(bg) });
        switch (BaselineOf(e))
        {
            case BaselineAlignment.Superscript: rp.Append(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript }); break;
            case BaselineAlignment.Subscript: rp.Append(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Subscript }); break;
        }
        return rp;
    }

    private static string FirstFamily(FontFamily f)
    {
        var name = f.Source.Split(',')[0].Trim();
        return name.Length == 0 ? StyleCatalog.DefaultFont : name;
    }

    /// Alti cizili / ustu cizili: ogenin kendisi ve ust satir-ici ogeleri (Span, Hyperlink) ile paragrafa bakilir.
    private static (bool underline, bool strike) Decorations(TextElement e)
    {
        bool u = false, s = false;
        for (DependencyObject? cur = e; cur is TextElement te; cur = te.Parent)
        {
            TextDecorationCollection? decos = te switch
            {
                Inline i => i.TextDecorations,
                Paragraph p => p.TextDecorations,
                _ => null,
            };
            if (decos != null)
            {
                foreach (var d in decos)
                {
                    if (d.Location == TextDecorationLocation.Underline) u = true;
                    if (d.Location == TextDecorationLocation.Strikethrough) s = true;
                }
            }
            if (te is Block) break;
        }
        return (u, s);
    }

    private static Color? BackgroundOf(TextElement e)
    {
        for (DependencyObject? cur = e; cur is Inline i; cur = i.Parent)
            if (i.Background is SolidColorBrush b && b.Color.A > 0) return b.Color;
        return null;
    }

    private static BaselineAlignment BaselineOf(TextElement e)
    {
        for (DependencyObject? cur = e; cur is Inline i; cur = i.Parent)
            if (i.BaselineAlignment is BaselineAlignment.Superscript or BaselineAlignment.Subscript) return i.BaselineAlignment;
        return BaselineAlignment.Baseline;
    }

    // ---------------------------------------------------------------- resimler

    private W.Drawing ImageDrawing(ImageData d, FrameworkElement fe)
    {
        var type = d.Extension.ToLowerInvariant() switch
        {
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            ".emf" => ImagePartType.Emf,
            ".wmf" => ImagePartType.Wmf,
            ".ico" => ImagePartType.Icon,
            _ => ImagePartType.Png,
        };
        var part = _main.AddImagePart(type);
        using (var ms = new MemoryStream(d.Bytes)) part.FeedData(ms);
        var relId = _main.GetIdOfPart(part);

        double w = !double.IsNaN(fe.Width) ? fe.Width : (fe.ActualWidth > 0 ? fe.ActualWidth : 100);
        double h = !double.IsNaN(fe.Height) ? fe.Height : (fe.ActualHeight > 0 ? fe.ActualHeight : 100);
        long cx = (long)(w * Units.EmuPerPx), cy = (long)(h * Units.EmuPerPx);
        uint id = _drawingId++;

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = id, Name = $"Resim {id}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"resim{id}{d.Extension}" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };
        return new W.Drawing(inline);
    }

    // ---------------------------------------------------------------- tablolar

    private W.Table ConvertTable(Table t)
    {
        int cols = t.Columns.Count;
        foreach (var g in t.RowGroups)
            foreach (var r in g.Rows)
                cols = Math.Max(cols, r.Cells.Sum(c => Math.Max(1, c.ColumnSpan)));
        cols = Math.Max(cols, 1);

        // Sutun genislikleri: sabit olanlar aynen, yildizli olanlar kalan genisligi paylasir.
        var widths = new double[cols];
        double fixedSum = 0;
        int starCount = 0;
        for (int i = 0; i < cols; i++)
        {
            var col = i < t.Columns.Count ? t.Columns[i] : null;
            if (col != null && col.Width.IsAbsolute) { widths[i] = col.Width.Value; fixedSum += widths[i]; }
            else starCount++;
        }
        if (starCount > 0)
        {
            double each = Math.Max(30, (_page.ContentWidth - fixedSum) / starCount);
            for (int i = 0; i < cols; i++) if (widths[i] == 0) widths[i] = each;
        }

        bool borders = t.BorderThickness.Left > 0 || t.BorderThickness.Top > 0 ||
                       t.RowGroups.SelectMany(g => g.Rows).SelectMany(r => r.Cells).Any(c => c.BorderThickness.Right > 0 || c.BorderThickness.Bottom > 0);

        var wt = new W.Table();
        var tblPr = new W.TableProperties(new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto });
        if (borders)
        {
            tblPr.Append(new W.TableBorders(
                new W.TopBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" },
                new W.LeftBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" },
                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" },
                new W.RightBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" },
                new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" },
                new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 4U, Color = "808080" }));
        }
        tblPr.Append(new W.TableLayout { Type = W.TableLayoutValues.Fixed });
        wt.Append(tblPr);

        var grid = new W.TableGrid();
        foreach (var w in widths) grid.Append(new W.GridColumn { Width = Tw(w) });
        wt.Append(grid);

        // Dikey birlesik hucreler: alttaki satirlara "devam" hucresi yazilir.
        var pending = new (int rows, int span)[cols];
        foreach (var group in t.RowGroups)
        {
            foreach (var row in group.Rows)
            {
                var tr = new W.TableRow();
                int col = 0, ci = 0;
                var cells = row.Cells;
                while (col < cols)
                {
                    if (pending[col].rows > 0)
                    {
                        int pspan = Math.Min(pending[col].span, cols - col);
                        pending[col].rows--;
                        tr.Append(MergeContinueCell(pspan, widths.Skip(col).Take(pspan).Sum()));
                        col += pspan;
                        continue;
                    }
                    if (ci >= cells.Count)
                    {
                        tr.Append(new W.TableCell(new W.TableCellProperties(new W.TableCellWidth { Width = Tw(widths[col]), Type = W.TableWidthUnitValues.Dxa }), new W.Paragraph()));
                        col++;
                        continue;
                    }
                    var c = cells[ci++];
                    int span = Math.Clamp(c.ColumnSpan, 1, cols - col);
                    var tcPr = new W.TableCellProperties(new W.TableCellWidth { Width = Tw(widths.Skip(col).Take(span).Sum()), Type = W.TableWidthUnitValues.Dxa });
                    if (span > 1) tcPr.Append(new W.GridSpan { Val = span });
                    if (c.RowSpan > 1)
                    {
                        tcPr.Append(new W.VerticalMerge { Val = W.MergedCellValues.Restart });
                        pending[col] = (c.RowSpan - 1, span);
                    }
                    if (c.Background is SolidColorBrush bg && bg.Color.A > 0)
                        tcPr.Append(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = DocHelpers.ToHex(bg.Color) });

                    var tc = new W.TableCell(tcPr);
                    var content = ConvertBlocks(c.Blocks).ToList();
                    // Word her hucrenin bir paragrafla bitmesini ister.
                    if (content.Count == 0 || content[^1] is not W.Paragraph) content.Add(new W.Paragraph());
                    tc.Append(content);
                    tr.Append(tc);
                    col += span;
                }
                wt.Append(tr);
            }
        }
        return wt;
    }

    private static W.TableCell MergeContinueCell(int span, double width)
    {
        var tcPr = new W.TableCellProperties(new W.TableCellWidth { Width = Tw(width), Type = W.TableWidthUnitValues.Dxa });
        if (span > 1) tcPr.Append(new W.GridSpan { Val = span });
        tcPr.Append(new W.VerticalMerge());
        return new W.TableCell(tcPr, new W.Paragraph());
    }

    // ---------------------------------------------------------------- numaralandirma

    private int NewNum(TextMarkerStyle marker, int level, int start)
    {
        int abs = AbstractFor(marker);
        int numId = _nums.Count + 1;
        var num = new W.NumberingInstance(
            new W.AbstractNumId { Val = abs },
            new W.LevelOverride(new W.StartOverrideNumberingValue { Val = Math.Max(0, start) }) { LevelIndex = level })
        { NumberID = numId };
        _nums.Add(num);
        return numId;
    }

    private int AbstractFor(TextMarkerStyle marker)
    {
        if (_abstractByMarker.TryGetValue(marker, out var existing)) return existing;
        int id = _abstracts.Count;
        var abs = new W.AbstractNum(new W.MultiLevelType { Val = W.MultiLevelValues.HybridMultilevel }) { AbstractNumberId = id };
        for (int l = 0; l < 9; l++)
        {
            (W.NumberFormatValues fmt, string text, string? font) = marker switch
            {
                TextMarkerStyle.Disc => (W.NumberFormatValues.Bullet, "•", null),
                TextMarkerStyle.Circle => (W.NumberFormatValues.Bullet, "o", "Courier New"),
                TextMarkerStyle.Square or TextMarkerStyle.Box => (W.NumberFormatValues.Bullet, "▪", null),
                TextMarkerStyle.LowerLatin => (W.NumberFormatValues.LowerLetter, $"%{l + 1}.", null),
                TextMarkerStyle.UpperLatin => (W.NumberFormatValues.UpperLetter, $"%{l + 1}.", null),
                TextMarkerStyle.LowerRoman => (W.NumberFormatValues.LowerRoman, $"%{l + 1}.", null),
                TextMarkerStyle.UpperRoman => (W.NumberFormatValues.UpperRoman, $"%{l + 1}.", null),
                TextMarkerStyle.None => (W.NumberFormatValues.None, "", null),
                _ => (W.NumberFormatValues.Decimal, $"%{l + 1}.", null),
            };
            var lvl = new W.Level(
                new W.StartNumberingValue { Val = 1 },
                new W.NumberingFormat { Val = fmt },
                new W.LevelText { Val = text },
                new W.LevelJustification { Val = W.LevelJustificationValues.Left },
                new W.PreviousParagraphProperties(new W.Indentation { Left = (720 * (l + 1)).ToString(Inv), Hanging = "360" }))
            { LevelIndex = l };
            if (font != null) lvl.Append(new W.NumberingSymbolRunProperties(new W.RunFonts { Ascii = font, HighAnsi = font }));
            abs.Append(lvl);
        }
        _abstracts.Add(abs);
        _abstractByMarker[marker] = id;
        return id;
    }

    private void AddNumbering()
    {
        var part = _main.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new W.Numbering();
        foreach (var a in _abstracts) numbering.Append(a);
        foreach (var n in _nums) numbering.Append(n);
        part.Numbering = numbering;
        part.Numbering.Save();
    }

    // ---------------------------------------------------------------- stiller ve sayfa

    private void AddStyles()
    {
        var part = _main.AddNewPart<StyleDefinitionsPart>();
        var font = StyleCatalog.DefaultFont;
        var styles = new W.Styles(
            new W.DocDefaults(
                new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                    new W.RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font },
                    new W.FontSize { Val = "22" },
                    new W.FontSizeComplexScript { Val = "22" },
                    new W.Languages { Val = "tr-TR", EastAsia = "en-US", Bidi = "ar-SA" })),
                new W.ParagraphPropertiesDefault(new W.ParagraphPropertiesBaseStyle(
                    new W.SpacingBetweenLines { After = "160", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }))));

        foreach (var s in StyleCatalog.All)
        {
            var style = new W.Style { Type = W.StyleValues.Paragraph, StyleId = s.Id };
            style.Append(new W.StyleName { Val = s.DocxName });
            if (s == StyleCatalog.Normal)
            {
                style.Default = true;
            }
            else
            {
                style.Append(new W.BasedOn { Val = StyleCatalog.Normal.Id });
                style.Append(new W.NextParagraphStyle { Val = StyleCatalog.Normal.Id });
            }
            style.Append(new W.UIPriority { Val = s.OutlineLevel >= 0 ? 9 : 10 });
            style.Append(new W.PrimaryStyle());
            var ppr = new W.StyleParagraphProperties();
            if (s.OutlineLevel >= 0) ppr.Append(new W.KeepNext());
            ppr.Append(new W.SpacingBetweenLines { Before = Tw(Units.PtToPx(s.BeforePt)), After = Tw(Units.PtToPx(s.AfterPt)) });
            if (s.OutlineLevel >= 0) ppr.Append(new W.OutlineLevel { Val = s.OutlineLevel });
            style.Append(ppr);
            var rpr = new W.StyleRunProperties();
            if (s.Bold) rpr.Append(new W.Bold());
            if (s.Color != Colors.Black) rpr.Append(new W.Color { Val = DocHelpers.ToHex(s.Color) });
            var hp = ((int)Math.Round(s.SizePt * 2)).ToString(Inv);
            rpr.Append(new W.FontSize { Val = hp });
            rpr.Append(new W.FontSizeComplexScript { Val = hp });
            style.Append(rpr);
            styles.Append(style);
        }
        part.Styles = styles;
        part.Styles.Save();
    }

    private W.SectionProperties SectionProperties()
    {
        var size = new W.PageSize { Width = (uint)Units.PxToTwips(_page.Width), Height = (uint)Units.PxToTwips(_page.Height) };
        if (_page.IsLandscape) size.Orient = W.PageOrientationValues.Landscape;
        var margin = new W.PageMargin
        {
            Top = (int)Units.PxToTwips(_page.MarginTop),
            Right = (uint)Units.PxToTwips(_page.MarginRight),
            Bottom = (int)Units.PxToTwips(_page.MarginBottom),
            Left = (uint)Units.PxToTwips(_page.MarginLeft),
            Header = 708U,
            Footer = 708U,
            Gutter = 0U,
        };
        return new W.SectionProperties(size, margin);
    }

    private static double Nz(double v) => double.IsNaN(v) ? 0 : v;
    private static string Tw(double px) => Math.Max(0, Units.PxToTwips(px)).ToString(Inv);
}
