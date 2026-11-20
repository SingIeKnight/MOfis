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
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MOfis.Writer;

/// DOCX -> FlowDocument. Stil kalitimi (belge varsayilanlari -> paragraf stili -> karakter stili ->
/// dogrudan bicim), tema fontlari, listeler, tablolar, resimler ve baglantilar desteklenir.
/// Desteklenmeyen ogeler atlanir ve Warnings listesinde Turkce olarak bildirilir.
public sealed class DocxImporter
{
    public sealed record Result(FlowDocument Document, PageSettings Page, IReadOnlyList<string> Warnings);

    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private readonly List<string> _warnings = [];
    private readonly Dictionary<string, W.Style> _styles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, FontFamily> _fontCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<(int numId, int level), int> _listCounts = [];
    private readonly PageSettings _page = new();
    private MainDocumentPart _main = null!;
    private W.Numbering? _numbering;
    private string? _defaultParaStyle;
    private string _minorFont = "Calibri", _majorFont = "Calibri Light";
    private RunFmt _docRun = new() { Font = "Times New Roman", SizePt = 10 };
    private ParaFmt _docPara = new();
    private bool _pendingPageBreak;
    private bool _paraHasContent;

    public static Result Load(Stream stream)
    {
        using var pkg = WordprocessingDocument.Open(stream, false);
        return new DocxImporter().Read(pkg);
    }

    private Result Read(WordprocessingDocument pkg)
    {
        _main = pkg.MainDocumentPart ?? throw new InvalidDataException("Belgenin ana içeriği bulunamadı.");
        var body = _main.Document?.Body ?? throw new InvalidDataException("Belgenin gövdesi bulunamadı.");

        LoadTheme();
        LoadStyles();
        _numbering = _main.NumberingDefinitionsPart?.Numbering;
        if (body.Elements<W.SectionProperties>().LastOrDefault() is { } sect) ReadPage(sect);
        if (body.Descendants<W.SectionProperties>().Count() > 1) Warn("Birden fazla bölüm (tek sayfa düzeni kullanılır)");

        var doc = new FlowDocument();
        DocHelpers.Prepare(doc);
        var normal = ResolveStyleRun(_defaultParaStyle);
        doc.FontFamily = Family(normal.Font ?? StyleCatalog.DefaultFont);
        doc.FontSize = Units.PtToPx(normal.SizePt ?? 11);

        foreach (var b in ConvertBlocks(body.ChildElements)) doc.Blocks.Add(b);
        if (doc.Blocks.Count == 0) doc.Blocks.Add(new Paragraph());

        if (_main.HeaderParts.Any(h => !string.IsNullOrWhiteSpace(h.Header?.InnerText)) ||
            _main.FooterParts.Any(f => !string.IsNullOrWhiteSpace(f.Footer?.InnerText)))
            Warn("Üst bilgi / alt bilgi");
        if (_main.WordprocessingCommentsPart?.Comments?.HasChildren == true) Warn("Açıklamalar (yorumlar)");

        return new Result(doc, _page, _warnings);
    }

    private void Warn(string what)
    {
        if (!_warnings.Contains(what)) _warnings.Add(what);
    }

    // ---------------------------------------------------------------- tema, stiller, sayfa

    private void LoadTheme()
    {
        var fs = _main.ThemePart?.Theme?.ThemeElements?.FontScheme;
        if (fs?.MinorFont?.LatinFont?.Typeface?.Value is { Length: > 0 } minor) _minorFont = minor;
        if (fs?.MajorFont?.LatinFont?.Typeface?.Value is { Length: > 0 } major) _majorFont = major;
    }

    private void LoadStyles()
    {
        var styles = _main.StyleDefinitionsPart?.Styles;
        if (styles == null) return;

        var defaults = styles.DocDefaults;
        _docRun.Merge(defaults?.RunPropertiesDefault?.RunPropertiesBaseStyle, this);
        _docPara.Merge(defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle);

        foreach (var s in styles.Elements<W.Style>())
        {
            if (s.StyleId?.Value is not { } id || _styles.ContainsKey(id)) continue;
            _styles[id] = s;
            if (s.Type?.InnerText == "paragraph" && s.Default?.Value == true) _defaultParaStyle = id;
        }
    }

    /// Stilin kalitim zinciri, en temelden en ozele dogru.
    private List<W.Style> Chain(string? styleId)
    {
        var chain = new List<W.Style>();
        var guard = new HashSet<string>();
        while (styleId != null && guard.Add(styleId) && _styles.TryGetValue(styleId, out var s))
        {
            chain.Insert(0, s);
            styleId = s.BasedOn?.Val?.Value;
        }
        return chain;
    }

    private RunFmt ResolveStyleRun(string? styleId)
    {
        var f = _docRun.Clone();
        foreach (var s in Chain(styleId)) f.Merge(s.StyleRunProperties, this);
        return f;
    }

    private void ReadPage(W.SectionProperties sect)
    {
        var size = sect.GetFirstChild<W.PageSize>();
        if (size?.Width?.Value is uint w && w > 1440) _page.Width = Units.TwipsToPx(w);
        if (size?.Height?.Value is uint h && h > 1440) _page.Height = Units.TwipsToPx(h);
        var m = sect.GetFirstChild<W.PageMargin>();
        if (m == null) return;
        if (m.Top?.Value is int top) _page.MarginTop = Units.TwipsToPx(Math.Abs(top));
        if (m.Bottom?.Value is int bottom) _page.MarginBottom = Units.TwipsToPx(Math.Abs(bottom));
        if (m.Left?.Value is uint left) _page.MarginLeft = Units.TwipsToPx(left);
        if (m.Right?.Value is uint right) _page.MarginRight = Units.TwipsToPx(right);
    }

    internal string ThemeFont(string themeValue) =>
        themeValue.StartsWith("major", StringComparison.OrdinalIgnoreCase) ? _majorFont : _minorFont;

    // ---------------------------------------------------------------- bloklar

    private List<Block> ConvertBlocks(IEnumerable<OpenXmlElement> elements)
    {
        var output = new List<Block>();
        var stack = new List<(List list, int numId)>(); // indeks = liste duzeyi
        (Paragraph para, ParaFmt fmt)? prev = null;

        foreach (var el in elements)
        {
            if (el is not W.Paragraph) prev = null;
            switch (el)
            {
                case W.Paragraph p:
                {
                    var pf = ResolvePara(p);
                    var para = BuildParagraph(p, pf);
                    // "Ayni stildeki paragraflar arasina bosluk ekleme" (Word liste paragraflarinda varsayilan).
                    if (prev is (var prevPara, var prevFmt) && pf.ContextualSpacing && prevFmt.ContextualSpacing && pf.StyleId == prevFmt.StyleId)
                    {
                        prevPara.Margin = new Thickness(prevPara.Margin.Left, prevPara.Margin.Top, prevPara.Margin.Right, 0);
                        para.Margin = new Thickness(para.Margin.Left, 0, para.Margin.Right, para.Margin.Bottom);
                    }
                    prev = (para, pf);
                    if (pf.NumId is int nid && nid != 0 && _numbering != null)
                    {
                        int lvl = Math.Clamp(pf.Ilvl ?? 0, 0, 8);
                        while (stack.Count > lvl + 1) stack.RemoveAt(stack.Count - 1);
                        if (stack.Count == lvl + 1 && stack[lvl].numId != nid) stack.RemoveAt(lvl);
                        while (stack.Count < lvl + 1)
                        {
                            int level = stack.Count;
                            var (marker, start) = MarkerFor(nid, level);
                            _listCounts.TryGetValue((nid, level), out var already);
                            var list = new List
                            {
                                MarkerStyle = marker,
                                StartIndex = Math.Max(1, start + already),
                                Margin = new Thickness(0),
                                Padding = new Thickness(36, 0, 0, 0),
                            };
                            if (stack.Count == 0)
                            {
                                output.Add(list);
                            }
                            else
                            {
                                var parent = stack[^1].list;
                                if (parent.ListItems.Count == 0) parent.ListItems.Add(new ListItem());
                                parent.ListItems.LastListItem.Blocks.Add(list);
                            }
                            stack.Add((list, nid));
                        }
                        // Liste girintisini WPF listesi verir; paragrafin kendi sol girintisini sifirla.
                        para.Margin = new Thickness(0, para.Margin.Top, para.Margin.Right, para.Margin.Bottom);
                        para.TextIndent = 0;
                        stack[lvl].list.ListItems.Add(new ListItem(para));
                        _listCounts[(nid, lvl)] = _listCounts.GetValueOrDefault((nid, lvl)) + 1;
                        // Ust duzeyde yeni madde gelince alt duzeyler yeniden 1'den baslar.
                        foreach (var key in _listCounts.Keys.Where(k => k.numId == nid && k.level > lvl).ToList())
                            _listCounts.Remove(key);
                    }
                    else
                    {
                        stack.Clear();
                        output.Add(para);
                    }
                    break;
                }
                case W.Table t:
                    stack.Clear();
                    output.Add(BuildTable(t));
                    break;
                case W.SdtBlock sdt:
                    stack.Clear();
                    if (sdt.SdtContentBlock != null) output.AddRange(ConvertBlocks(sdt.SdtContentBlock.ChildElements));
                    break;
                case W.CustomXmlBlock cx:
                    stack.Clear();
                    output.AddRange(ConvertBlocks(cx.ChildElements));
                    break;
                case W.AltChunk:
                    Warn("Gömülü belge parçaları");
                    break;
                default:
                    if (el.NamespaceUri == MathNs) Warn("Denklemler");
                    break;
            }
        }
        return output;
    }

    private const string MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private ParaFmt ResolvePara(W.Paragraph p)
    {
        var ppr = p.ParagraphProperties;
        var styleId = ppr?.ParagraphStyleId?.Val?.Value ?? _defaultParaStyle;
        var pf = _docPara.Clone();
        var run = _docRun.Clone();
        var chain = Chain(styleId);
        foreach (var s in chain)
        {
            pf.Merge(s.StyleParagraphProperties);
            run.Merge(s.StyleRunProperties, this);
        }
        pf.Merge(ppr);
        pf.StyleId = styleId;
        pf.Run = run;
        var mark = run.Clone();
        mark.Merge(ppr?.ParagraphMarkRunProperties, this);
        pf.Mark = mark;
        // En ozel stilden geriye: bilinen bir stil (Baslik 1 vb.) bulursak onu kullan.
        for (int i = chain.Count - 1; i >= 0; i--)
        {
            if (StyleCatalog.ByDocxName(chain[i].StyleName?.Val?.Value) is { } known)
            {
                pf.Style = known;
                break;
            }
        }
        return pf;
    }

    private Paragraph BuildParagraph(W.Paragraph p, ParaFmt pf)
    {
        var para = new Paragraph();
        if (pf.Style != null && pf.Style != StyleCatalog.Normal) para.Tag = pf.Style.Id;
        para.TextAlignment = pf.Jc switch
        {
            "center" => TextAlignment.Center,
            "right" or "end" => TextAlignment.Right,
            "both" or "distribute" => TextAlignment.Justify,
            _ => TextAlignment.Left,
        };
        para.Margin = new Thickness(
            Units.TwipsToPx(Math.Max(0, pf.IndLeft ?? 0)),
            Units.TwipsToPx(Math.Max(0, pf.Before ?? 0)),
            Units.TwipsToPx(Math.Max(0, pf.IndRight ?? 0)),
            Units.TwipsToPx(Math.Max(0, pf.After ?? 0)));
        if (pf.FirstLine is double fl && fl > 0) para.TextIndent = Units.TwipsToPx(fl);
        else if (pf.Hanging is double hg && hg > 0) para.TextIndent = -Units.TwipsToPx(hg);

        ApplyRunFmt(para, pf.Mark ?? pf.Run ?? _docRun);

        if (pf.Line is double line && line > 0)
        {
            if (pf.LineRule is null or "auto")
            {
                double factor = line / 240.0;
                if (Math.Abs(factor - 1) > 0.02)
                    para.LineHeight = Math.Clamp(factor * para.FontFamily.LineSpacing * para.FontSize, 1, 5000);
            }
            else
            {
                para.LineHeight = Math.Clamp(Units.TwipsToPx(line), 1, 5000);
            }
        }

        if (pf.KeepNext) para.KeepWithNext = true;
        if (_pendingPageBreak || pf.PageBreakBefore)
        {
            para.BreakPageBefore = true;
            _pendingPageBreak = false;
        }

        _paraHasContent = false;
        AddInlines(p.ChildElements, para.Inlines, pf, para);
        return para;
    }

    // ---------------------------------------------------------------- satir ici ogeler

    private void AddInlines(IEnumerable<OpenXmlElement> children, InlineCollection target, ParaFmt pf, Paragraph para)
    {
        foreach (var el in children)
        {
            switch (el)
            {
                case W.Run r:
                    AddRun(r, target, pf, para);
                    break;
                case W.Hyperlink h:
                {
                    Uri? uri = null;
                    if (h.Id?.Value is { } id)
                        uri = _main.HyperlinkRelationships.FirstOrDefault(x => x.Id == id)?.Uri;
                    if (uri != null)
                    {
                        var link = new Hyperlink { NavigateUri = uri };
                        AddInlines(h.ChildElements, link.Inlines, pf, para);
                        target.Add(link);
                    }
                    else
                    {
                        AddInlines(h.ChildElements, target, pf, para);
                    }
                    break;
                }
                case W.SimpleField f:
                    AddInlines(f.ChildElements, target, pf, para);
                    break;
                case W.InsertedRun ins:
                    AddInlines(ins.ChildElements, target, pf, para);
                    break;
                case W.DeletedRun:
                    break;
                case W.SdtRun sdt:
                    if (sdt.SdtContentRun != null) AddInlines(sdt.SdtContentRun.ChildElements, target, pf, para);
                    break;
                case W.CustomXmlRun cx:
                    AddInlines(cx.ChildElements, target, pf, para);
                    break;
                default:
                    if (el.NamespaceUri == MathNs) Warn("Denklemler");
                    break;
            }
        }
    }

    private void AddRun(W.Run r, InlineCollection target, ParaFmt pf, Paragraph para)
    {
        var fmt = (pf.Run ?? _docRun).Clone();
        if (r.RunProperties?.RunStyle?.Val?.Value is { } rs)
            foreach (var s in Chain(rs)) fmt.Merge(s.StyleRunProperties, this);
        fmt.Merge(r.RunProperties, this);

        var sb = new StringBuilder();
        void Flush()
        {
            if (sb.Length == 0) return;
            var text = fmt.Caps ? sb.ToString().ToUpper(new CultureInfo("tr-TR")) : sb.ToString();
            var run = new Run(text);
            ApplyRunFmt(run, fmt);
            target.Add(run);
            sb.Clear();
            _paraHasContent = true;
        }

        foreach (var child in r.ChildElements)
        {
            switch (child)
            {
                case W.Text t:
                    sb.Append(t.Text);
                    break;
                case W.TabChar:
                    sb.Append('\t');
                    break;
                case W.Break br:
                    Flush();
                    var type = br.Type?.InnerText;
                    if (type == "page")
                    {
                        if (!_paraHasContent) para.BreakPageBefore = true;
                        else _pendingPageBreak = true;
                    }
                    else if (type != "column")
                    {
                        var lb = new LineBreak();
                        target.Add(lb);
                    }
                    break;
                case W.CarriageReturn:
                    Flush();
                    target.Add(new LineBreak());
                    break;
                case W.NoBreakHyphen:
                    sb.Append('‑');
                    break;
                case W.SymbolChar sym:
                    if (sym.Char?.Value is { } code && int.TryParse(code, NumberStyles.HexNumber, Inv, out var cp))
                    {
                        if (cp >= 0xF000) cp -= 0xF000;
                        sb.Append(cp == 0xB7 || cp == 0xA7 ? '•' : (char)cp);
                    }
                    break;
                case W.Drawing d:
                    Flush();
                    AddDrawing(d, target);
                    break;
                case AlternateContent ac:
                    Flush();
                    if (ac.Descendants<W.Drawing>().FirstOrDefault(x => x.Descendants<A.Blip>().Any()) is { } inner) AddDrawing(inner, target);
                    else Warn("Şekiller ve metin kutuları");
                    break;
                case W.Picture:
                    Warn("Eski biçimli resim ve şekiller");
                    break;
                case W.FootnoteReference:
                    Warn("Dipnotlar");
                    break;
                case W.EndnoteReference:
                    Warn("Son notlar");
                    break;
                case W.EmbeddedObject:
                    Warn("Gömülü nesneler");
                    break;
            }
        }
        Flush();
    }

    private void AddDrawing(W.Drawing d, InlineCollection target)
    {
        var blip = d.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value is { } rid)
        {
            ImagePart? part = null;
            try { part = _main.GetPartById(rid) as ImagePart; } catch (ArgumentOutOfRangeException) { }
            if (part != null)
            {
                using var s = part.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                var extent = d.Descendants<DW.Extent>().FirstOrDefault();
                double? w = extent?.Cx?.Value is long cx ? cx / (double)Units.EmuPerPx : null;
                double? h = extent?.Cy?.Value is long cy ? cy / (double)Units.EmuPerPx : null;
                var el = DocHelpers.CreateImageElement(new ImageData(ms.ToArray(), ExtensionFor(part.ContentType)), w, h, _page.ContentWidth);
                target.Add(new InlineUIContainer(el));
                _paraHasContent = true;
                if (d.Anchor != null) Warn("Yüzen resimler (satır içinde gösterilir)");
                return;
            }
        }
        if (d.Descendants<W.TextBoxContent>().Any()) Warn("Metin kutuları");
        else if (d.Descendants().Any(x => x.LocalName == "chart")) Warn("Grafikler");
        else Warn("Şekiller");
    }

    private static string ExtensionFor(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/jpeg" or "image/jpg" => ".jpg",
        "image/gif" => ".gif",
        "image/bmp" => ".bmp",
        "image/tiff" => ".tiff",
        "image/x-emf" or "image/emf" => ".emf",
        "image/x-wmf" or "image/wmf" => ".wmf",
        "image/x-icon" => ".ico",
        _ => ".png",
    };

    // ---------------------------------------------------------------- tablolar

    private Table BuildTable(W.Table t)
    {
        var line = new SolidColorBrush(Color.FromRgb(0x80, 0x80, 0x80));
        line.Freeze();
        var table = new Table { CellSpacing = 0 };

        var rows = t.Elements<W.TableRow>().ToList();
        var gridWidths = t.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>()
            .Select(g => Units.TwipsToPx(ParseNum(g.Width?.Value) ?? 1440)).ToList() ?? [];
        int cols = Math.Max(gridWidths.Count, rows.Count == 0 ? 1 : rows.Max(r => CellsWithPos(r).Select(c => c.start + c.span).DefaultIfEmpty(0).Max()));
        while (gridWidths.Count < cols) gridWidths.Add(_page.ContentWidth / cols);
        double sum = gridWidths.Sum();
        // Kenarlik cizgileri (sol 1 px + sag 1 px) de yazi alanina sigmali; yoksa son cizgi kirpilir.
        double room = _page.ContentWidth - 2;
        double scale = sum > room && sum > 0 ? room / sum : 1;
        foreach (var w in gridWidths) table.Columns.Add(new TableColumn { Width = new GridLength(Math.Max(10, w * scale)) });

        bool borders = HasBorders(t);
        if (borders)
        {
            table.BorderBrush = line;
            table.BorderThickness = new Thickness(1, 1, 0, 0);
        }

        var group = new TableRowGroup();
        table.RowGroups.Add(group);
        for (int ri = 0; ri < rows.Count; ri++)
        {
            var tr = new TableRow();
            foreach (var (tc, start, span) in CellsWithPos(rows[ri]))
            {
                var tcPr = tc.TableCellProperties;
                var vm = tcPr?.VerticalMerge;
                if (vm != null && vm.Val?.InnerText != "restart") continue; // dikey birlesmenin devami

                var cell = new TableCell { ColumnSpan = span, Padding = new Thickness(7, 2, 7, 0) };
                if (vm != null)
                {
                    int rowSpan = 1;
                    for (int below = ri + 1; below < rows.Count; below++)
                    {
                        var next = CellsWithPos(rows[below]).FirstOrDefault(c => c.start == start).tc;
                        var nvm = next?.TableCellProperties?.VerticalMerge;
                        if (nvm == null || nvm.Val?.InnerText == "restart") break;
                        rowSpan++;
                    }
                    cell.RowSpan = rowSpan;
                }
                if (DocHelpers.ParseHex(tcPr?.Shading?.Fill?.Value) is Color fill) cell.Background = new SolidColorBrush(fill);
                if (borders)
                {
                    cell.BorderBrush = line;
                    cell.BorderThickness = new Thickness(0, 0, 1, 1);
                }
                foreach (var b in ConvertBlocks(tc.ChildElements)) cell.Blocks.Add(b);
                if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
                tr.Cells.Add(cell);
            }
            group.Rows.Add(tr);
        }
        return table;
    }

    private static IEnumerable<(W.TableCell tc, int start, int span)> CellsWithPos(W.TableRow row)
    {
        int col = row.TableRowProperties?.GetFirstChild<W.GridBefore>()?.Val?.Value ?? 0;
        foreach (var tc in row.Elements<W.TableCell>())
        {
            int span = Math.Max(1, tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            yield return (tc, col, span);
            col += span;
        }
    }

    private bool HasBorders(W.Table t)
    {
        var direct = t.GetFirstChild<W.TableProperties>()?.TableBorders;
        if (direct != null) return BordersVisible(direct);
        var styleId = t.GetFirstChild<W.TableProperties>()?.TableStyle?.Val?.Value;
        foreach (var s in Enumerable.Reverse(Chain(styleId)))
        {
            var b = s.StyleTableProperties?.TableBorders;
            if (b != null) return BordersVisible(b);
        }
        // Hucre duzeyinde kenarlik varsa da cizgili goster.
        return t.Descendants<W.TableCellBorders>().Any(BordersVisible);
    }

    private static bool BordersVisible(OpenXmlElement borders) =>
        borders.ChildElements.OfType<W.BorderType>().Any(b => b.Val != null && b.Val.InnerText is not ("none" or "nil"));

    // ---------------------------------------------------------------- listeler

    private (TextMarkerStyle marker, int start) MarkerFor(int numId, int level)
    {
        var num = _numbering?.Elements<W.NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
        if (num == null) return (TextMarkerStyle.Disc, 1);

        var ovr = num.Elements<W.LevelOverride>().FirstOrDefault(o => o.LevelIndex?.Value == level);
        W.Level? lvl = ovr?.Level;
        if (lvl == null)
        {
            var absId = num.AbstractNumId?.Val?.Value;
            var abs = _numbering!.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
            lvl = abs?.Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == level);
        }
        int start = ovr?.StartOverrideNumberingValue?.Val?.Value ?? lvl?.StartNumberingValue?.Val?.Value ?? 1;
        var fmt = lvl?.NumberingFormat?.Val?.InnerText ?? "bullet";
        var text = lvl?.LevelText?.Val?.Value ?? "";
        var marker = fmt switch
        {
            "bullet" => text switch
            {
                "o" or "◦" or "○" => TextMarkerStyle.Circle,
                "▪" or "■" or "§" or "" => TextMarkerStyle.Square,
                _ => TextMarkerStyle.Disc,
            },
            "lowerLetter" => TextMarkerStyle.LowerLatin,
            "upperLetter" => TextMarkerStyle.UpperLatin,
            "lowerRoman" => TextMarkerStyle.LowerRoman,
            "upperRoman" => TextMarkerStyle.UpperRoman,
            "none" => TextMarkerStyle.None,
            _ => TextMarkerStyle.Decimal,
        };
        return (marker, start);
    }

    // ---------------------------------------------------------------- bicim uygulama

    private FontFamily Family(string name)
    {
        if (!_fontCache.TryGetValue(name, out var f))
        {
            f = new FontFamily(name + ", " + StyleCatalog.DefaultFont);
            _fontCache[name] = f;
        }
        return f;
    }

    private void ApplyRunFmt(TextElement el, RunFmt f)
    {
        el.FontFamily = Family(f.Font ?? StyleCatalog.DefaultFont);
        el.FontSize = Units.PtToPx(Math.Clamp(f.SizePt ?? 11, 1, 1638));
        el.FontWeight = f.Bold == true ? FontWeights.Bold : FontWeights.Normal;
        el.FontStyle = f.Italic == true ? FontStyles.Italic : FontStyles.Normal;
        el.Foreground = f.TextColor is Color c ? Frozen(c) : Brushes.Black;
        if (el is Inline inline)
        {
            if (f.Underline == true || f.Strike == true)
            {
                var decos = new TextDecorationCollection();
                if (f.Underline == true) decos.Add(TextDecorations.Underline);
                if (f.Strike == true) decos.Add(TextDecorations.Strikethrough);
                inline.TextDecorations = decos;
            }
            if (f.Fill is Color bg) inline.Background = Frozen(bg);
            if (f.Vert == "superscript") inline.BaselineAlignment = BaselineAlignment.Superscript;
            else if (f.Vert == "subscript") inline.BaselineAlignment = BaselineAlignment.Subscript;
        }
    }

    private static SolidColorBrush Frozen(Color c)
    {
        var b = new SolidColorBrush(c);
        b.Freeze();
        return b;
    }

    internal static double? ParseNum(string? s) =>
        s != null && double.TryParse(s, NumberStyles.Float, Inv, out var d) ? d : null;
}

/// Yazi bicimi (null = bu duzeyde belirtilmemis, ust duzeyden gelir).
internal sealed class RunFmt
{
    public string? Font;
    public double? SizePt;
    public bool? Bold, Italic, Underline, Strike;
    public bool Caps;
    public Color? TextColor;
    public Color? Fill;
    public string? Vert;

    public RunFmt Clone() => (RunFmt)MemberwiseClone();

    public void Merge(OpenXmlElement? rpr, DocxImporter owner)
    {
        if (rpr == null) return;
        if (rpr.GetFirstChild<W.RunFonts>() is { } rf)
        {
            var font = rf.Ascii?.Value ?? rf.HighAnsi?.Value;
            if (font == null && (rf.AsciiTheme ?? rf.HighAnsiTheme) is { } theme) font = owner.ThemeFont(theme.InnerText ?? "");
            if (font != null) Font = font;
        }
        if (rpr.GetFirstChild<W.Bold>() is { } b) Bold = b.Val == null || b.Val.Value;
        if (rpr.GetFirstChild<W.Italic>() is { } i) Italic = i.Val == null || i.Val.Value;
        if (rpr.GetFirstChild<W.Strike>() is { } s) Strike = s.Val == null || s.Val.Value;
        if (rpr.GetFirstChild<W.Caps>() is { } caps) Caps = caps.Val == null || caps.Val.Value;
        if (rpr.GetFirstChild<W.Underline>() is { } u) Underline = u.Val == null || u.Val.InnerText != "none";
        if (rpr.GetFirstChild<W.Color>()?.Val?.Value is { } c) TextColor = DocHelpers.ParseHex(c);
        if (DocxImporter.ParseNum(rpr.GetFirstChild<W.FontSize>()?.Val?.Value) is double hp) SizePt = hp / 2;
        if (rpr.GetFirstChild<W.Shading>()?.Fill?.Value is { } fill) Fill = DocHelpers.ParseHex(fill);
        if (rpr.GetFirstChild<W.Highlight>()?.Val is { } hl) Fill = HighlightColor(hl.InnerText ?? "") ?? Fill;
        if (rpr.GetFirstChild<W.VerticalTextAlignment>()?.Val is { } va) Vert = va.InnerText;
    }

    private static Color? HighlightColor(string name) => name switch
    {
        "yellow" => Colors.Yellow,
        "green" => Color.FromRgb(0x00, 0xFF, 0x00),
        "cyan" => Colors.Cyan,
        "magenta" => Colors.Magenta,
        "blue" => Colors.Blue,
        "red" => Colors.Red,
        "darkBlue" => Color.FromRgb(0x00, 0x00, 0x80),
        "darkCyan" => Color.FromRgb(0x00, 0x80, 0x80),
        "darkGreen" => Color.FromRgb(0x00, 0x80, 0x00),
        "darkMagenta" => Color.FromRgb(0x80, 0x00, 0x80),
        "darkRed" => Color.FromRgb(0x80, 0x00, 0x00),
        "darkYellow" => Color.FromRgb(0x80, 0x80, 0x00),
        "darkGray" => Color.FromRgb(0x80, 0x80, 0x80),
        "lightGray" => Color.FromRgb(0xC0, 0xC0, 0xC0),
        "black" => Colors.Black,
        "white" => Colors.White,
        _ => null,
    };
}

/// Paragraf bicimi (twip cinsinden degerler).
internal sealed class ParaFmt
{
    public string? Jc;
    public double? Before, After, Line;
    public string? LineRule;
    public double? IndLeft, IndRight, FirstLine, Hanging;
    public bool PageBreakBefore, KeepNext, ContextualSpacing;
    public string? StyleId;
    public int? NumId, Ilvl;
    public RunFmt? Run, Mark;
    public ParaStyleDef? Style;

    public ParaFmt Clone() => (ParaFmt)MemberwiseClone();

    public void Merge(OpenXmlElement? ppr)
    {
        if (ppr == null) return;
        if (ppr.GetFirstChild<W.Justification>()?.Val is { } jc) Jc = jc.InnerText;
        if (ppr.GetFirstChild<W.SpacingBetweenLines>() is { } sp)
        {
            Before = DocxImporter.ParseNum(sp.Before?.Value) ?? Before;
            After = DocxImporter.ParseNum(sp.After?.Value) ?? After;
            if (DocxImporter.ParseNum(sp.Line?.Value) is double line)
            {
                Line = line;
                LineRule = sp.LineRule?.InnerText ?? "auto";
            }
        }
        if (ppr.GetFirstChild<W.Indentation>() is { } ind)
        {
            IndLeft = DocxImporter.ParseNum(ind.Left?.Value ?? ind.Start?.Value) ?? IndLeft;
            IndRight = DocxImporter.ParseNum(ind.Right?.Value ?? ind.End?.Value) ?? IndRight;
            if (DocxImporter.ParseNum(ind.FirstLine?.Value) is double fl) { FirstLine = fl; Hanging = null; }
            if (DocxImporter.ParseNum(ind.Hanging?.Value) is double hg) { Hanging = hg; FirstLine = null; }
        }
        if (ppr.GetFirstChild<W.PageBreakBefore>() is { } pb) PageBreakBefore = pb.Val == null || pb.Val.Value;
        if (ppr.GetFirstChild<W.KeepNext>() is { } kn) KeepNext = kn.Val == null || kn.Val.Value;
        if (ppr.GetFirstChild<W.ContextualSpacing>() is { } cs) ContextualSpacing = cs.Val == null || cs.Val.Value;
        if (ppr.GetFirstChild<W.NumberingProperties>() is { } np)
        {
            if (np.NumberingId?.Val?.Value is int id) NumId = id;
            if (np.NumberingLevelReference?.Val?.Value is int lvl) Ilvl = lvl;
        }
    }
}
