using System.Globalization;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ClosedXML.Excel;

namespace MOfis.Sheet;

/// Hucre araligi (1 tabanli, R1<=R2, C1<=C2).
public readonly record struct CellRange(int R1, int C1, int R2, int C2)
{
    public static CellRange Of(int ra, int ca, int rb, int cb) => new(Math.Min(ra, rb), Math.Min(ca, cb), Math.Max(ra, rb), Math.Max(ca, cb));
    public bool Contains(int r, int c) => r >= R1 && r <= R2 && c >= C1 && c <= C2;
    public bool IsSingle => R1 == R2 && C1 == C2;
    public long Count => (long)(R2 - R1 + 1) * (C2 - C1 + 1);
    public bool IsWholeColumns => R1 == 1 && R2 == XLHelper.MaxRowNumber;
    public bool IsWholeRows => C1 == 1 && C2 == XLHelper.MaxColumnNumber;

    public string Address()
    {
        if (IsWholeColumns && IsWholeRows) return "A1:" + XLHelper.GetColumnLetterFromNumber(C2) + R2;
        if (IsWholeColumns) return XLHelper.GetColumnLetterFromNumber(C1) + ":" + XLHelper.GetColumnLetterFromNumber(C2);
        if (IsWholeRows) return R1 + ":" + R2;
        var a = XLHelper.GetColumnLetterFromNumber(C1) + R1;
        return IsSingle ? a : a + ":" + XLHelper.GetColumnLetterFromNumber(C2) + R2;
    }
}

/// Hesap tablosu izgarasi: yalnizca gorunen hucreleri DrawingContext ile cizer (sanallastirilmis).
/// Veri dogrudan ClosedXML calisma sayfasindan okunur.
public sealed class SheetGrid : FrameworkElement
{
    public const int MaxRow = XLHelper.MaxRowNumber;
    public const int MaxCol = XLHelper.MaxColumnNumber;
    private const double HeaderH = 22;

    private static readonly CultureInfo Tr = new("tr-TR");
    private static readonly Brush HeaderBg = Frozen(Color.FromRgb(0xF5, 0xF5, 0xF7));
    private static readonly Brush HeaderSelBg = Frozen(Color.FromRgb(0xE0, 0xE7, 0xFF));
    private static readonly Brush HeaderText = Frozen(Color.FromRgb(0x4B, 0x50, 0x5A));
    private static readonly Brush HeaderSelText = Frozen(Color.FromRgb(0x37, 0x30, 0xA3));
    private static readonly Brush Accent = Frozen(Color.FromRgb(0x43, 0x38, 0xCA));
    private static readonly Brush SelFill = Frozen(Color.FromArgb(0x26, 0x43, 0x38, 0xCA));
    private static readonly Pen GridPen = FrozenPen(Color.FromRgb(0xE1, 0xE1, 0xE7), 1);
    private static readonly Pen HeaderPen = FrozenPen(Color.FromRgb(0xD3, 0xD3, 0xDC), 1);
    private static readonly Pen AccentPen = FrozenPen(Color.FromRgb(0x43, 0x38, 0xCA), 2);
    private static readonly Pen CopyPen;

    static SheetGrid()
    {
        var p = new Pen(Brushes.Black.Clone(), 1.5) { DashStyle = new DashStyle([3, 2], 0) };
        p.Brush = new SolidColorBrush(Color.FromRgb(0x43, 0x38, 0xCA));
        p.Freeze();
        CopyPen = p;
    }

    private IXLWorksheet? _sheet;
    private readonly Dictionary<(string, bool, bool), Typeface> _typefaces = [];
    private readonly DispatcherTimer _autoScroll = new() { Interval = TimeSpan.FromMilliseconds(60) };
    private Point _lastMouse;
    private DragMode _drag;

    private enum DragMode { None, Cells, Rows, Cols }

    public AxisLayout Rows { get; } = new(20);
    public AxisLayout Cols { get; } = new(64);

    public int TopRow { get; private set; } = 1;
    public int LeftCol { get; private set; } = 1;
    public double Zoom { get; private set; } = 1;

    /// Etkin hucre (secimin baslangic kosesi) ve secimin hareketli kosesi.
    public int ActiveRow { get; private set; } = 1;
    public int ActiveCol { get; private set; } = 1;
    private int _extRow = 1, _extCol = 1;

    public CellRange Selection => CellRange.Of(ActiveRow, ActiveCol, _extRow, _extCol);

    /// Kopyalanan aralik (kesikli cerceve ile gosterilir).
    public CellRange? CopyMarquee { get; set; }

    public event Action? SelectionChanged;
    public event Action? ScrollChanged;
    /// Duzenleme baslat: null = mevcut icerikle (F2/cift tik), metin = yazilan ilk karakter(ler).
    public event Action<string?>? EditRequested;
    public event Action? DeleteRequested;
    public event Action<double>? ZoomRequested;

    public SheetGrid()
    {
        Focusable = true;
        FocusVisualStyle = null;
        ClipToBounds = true;
        SnapsToDevicePixels = true;
        _autoScroll.Tick += (_, _) => AutoScrollTick();
    }

    private static SolidColorBrush Frozen(Color c)
    {
        var b = new SolidColorBrush(c);
        b.Freeze();
        return b;
    }

    private static Pen FrozenPen(Color c, double w)
    {
        var p = new Pen(Frozen(c), w);
        p.Freeze();
        return p;
    }

    // ================================================================ veri baglama

    public IXLWorksheet? Sheet => _sheet;

    public void Bind(IXLWorksheet sheet)
    {
        _sheet = sheet;
        LoadLayout();
        TopRow = LeftCol = 1;
        ActiveRow = ActiveCol = _extRow = _extCol = 1;
        CopyMarquee = null;
        InvalidateVisual();
        ScrollChanged?.Invoke();
        SelectionChanged?.Invoke();
    }

    /// Sutun genislikleri ve satir yukseklikleri ClosedXML'den okunur (piksel).
    public void LoadLayout()
    {
        if (_sheet == null) return;
        Rows.Clear();
        Cols.Clear();
        Rows.Default = PtToPx(_sheet.RowHeight);
        Cols.Default = CharsToPx(_sheet.ColumnWidth);
        int lastCol = Math.Max(_sheet.LastColumnUsed(XLCellsUsedOptions.All)?.ColumnNumber() ?? 0, 30);
        foreach (var col in _sheet.Columns(1, lastCol))
        {
            double w = col.IsHidden ? 0 : CharsToPx(col.Width);
            Cols.Set(col.ColumnNumber(), w);
        }
        foreach (var row in _sheet.RowsUsed(XLCellsUsedOptions.All))
        {
            double h = row.IsHidden ? 0 : PtToPx(row.Height);
            Rows.Set(row.RowNumber(), h);
        }
    }

    public static double PtToPx(double pt) => Math.Round(pt * 96.0 / 72.0);

    /// Excel sutun genisligi (karakter) -> piksel (Calibri 11 icin en genis rakam 7 px).
    public static double CharsToPx(double chars) => Math.Round(chars * 7 + 5);

    // ================================================================ konum hesaplari

    private double RowHeaderW
    {
        get
        {
            int digits = Math.Max(3, (BottomRowGuess()).ToString(CultureInfo.InvariantCulture).Length);
            return 14 + digits * 7.5;
        }
    }

    private int BottomRowGuess() => TopRow + (int)(ActualHeight / Zoom / Math.Max(1, Rows.Default)) + 1;

    /// Logical (zoom oncesi) koordinatlar: hucre alani basi (RowHeaderW, HeaderH).
    private double ColX(int c) => RowHeaderW + Cols.Offset(c) - Cols.Offset(LeftCol);
    private double RowY(int r) => HeaderH + Rows.Offset(r) - Rows.Offset(TopRow);

    /// Kontrol koordinatlarinda (zoom uygulanmis) hucre dikdortgeni; birlesik hucreyse tum alan.
    public Rect CellRect(int r, int c)
    {
        var m = MergeAt(r, c);
        if (m != null)
        {
            double x1 = ColX(m.Value.C1), y1 = RowY(m.Value.R1);
            double x2 = ColX(m.Value.C2 + 1), y2 = RowY(m.Value.R2 + 1);
            return new Rect(x1 * Zoom, y1 * Zoom, (x2 - x1) * Zoom, (y2 - y1) * Zoom);
        }
        return new Rect(ColX(c) * Zoom, RowY(r) * Zoom, Cols.Size(c) * Zoom, Rows.Size(r) * Zoom);
    }

    public CellRange? MergeAt(int r, int c)
    {
        if (_sheet == null) return null;
        foreach (var m in _sheet.MergedRanges)
        {
            var a = m.RangeAddress;
            if (r >= a.FirstAddress.RowNumber && r <= a.LastAddress.RowNumber && c >= a.FirstAddress.ColumnNumber && c <= a.LastAddress.ColumnNumber)
                return new CellRange(a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber, a.LastAddress.RowNumber, a.LastAddress.ColumnNumber);
        }
        return null;
    }

    private int VisibleRowCount()
    {
        double h = ActualHeight / Zoom - HeaderH, y = 0;
        int r = TopRow, n = 0;
        while (y < h && r <= MaxRow) { y += Rows.Size(r); r++; n++; }
        return Math.Max(1, n - 1);
    }

    private int VisibleColCount()
    {
        double w = ActualWidth / Zoom - RowHeaderW, x = 0;
        int c = LeftCol, n = 0;
        while (x < w && c <= MaxCol) { x += Cols.Size(c); c++; n++; }
        return Math.Max(1, n - 1);
    }

    public int PageRows => VisibleRowCount();
    public int PageCols => VisibleColCount();

    // ================================================================ kaydirma, yakinlastirma

    public void ScrollTo(int top, int left)
    {
        top = Math.Clamp(top, 1, MaxRow);
        left = Math.Clamp(left, 1, MaxCol);
        if (top == TopRow && left == LeftCol) return;
        TopRow = top;
        LeftCol = left;
        InvalidateVisual();
        ScrollChanged?.Invoke();
    }

    public void SetZoom(double z)
    {
        Zoom = Math.Clamp(z, 0.25, 4);
        InvalidateVisual();
        ScrollChanged?.Invoke();
    }

    public void EnsureVisible(int r, int c)
    {
        int top = TopRow, left = LeftCol;
        if (r < TopRow) top = r;
        else
        {
            // Satirin alt kenari gorunur alanda mi?
            double avail = ActualHeight / Zoom - HeaderH;
            while (Rows.Offset(r + 1) - Rows.Offset(top) > avail && top < r) top++;
        }
        if (c < LeftCol) left = c;
        else
        {
            double avail = ActualWidth / Zoom - RowHeaderW;
            while (Cols.Offset(c + 1) - Cols.Offset(left) > avail && left < c) left++;
        }
        ScrollTo(top, left);
    }

    // ================================================================ secim

    public void Select(int r, int c, bool extend = false)
    {
        r = Math.Clamp(r, 1, MaxRow);
        c = Math.Clamp(c, 1, MaxCol);
        if (extend)
        {
            _extRow = r;
            _extCol = c;
        }
        else
        {
            // Birlesik hucrenin icine gelinirse sol ust hucresi etkin olur.
            var m = MergeAt(r, c);
            if (m != null) { r = m.Value.R1; c = m.Value.C1; }
            ActiveRow = _extRow = r;
            ActiveCol = _extCol = c;
            if (m != null) { _extRow = m.Value.R2; _extCol = m.Value.C2; }
        }
        InvalidateVisual();
        SelectionChanged?.Invoke();
    }

    public void SelectRange(CellRange range)
    {
        ActiveRow = range.R1;
        ActiveCol = range.C1;
        _extRow = range.R2;
        _extCol = range.C2;
        InvalidateVisual();
        SelectionChanged?.Invoke();
    }

    /// Enter/Tab ile secim icinde ilerleme (secim tek hucreyse sadece tasir).
    public void MoveActive(int dr, int dc)
    {
        var sel = Selection;
        if (sel.IsSingle || MergeAt(ActiveRow, ActiveCol) is { } m && m == sel)
        {
            var mm = MergeAt(ActiveRow, ActiveCol);
            int r = ActiveRow, c = ActiveCol;
            if (mm != null) { if (dr > 0) r = mm.Value.R2; if (dc > 0) c = mm.Value.C2; }
            Select(r + dr, c + dc);
            EnsureVisible(ActiveRow, ActiveCol);
            return;
        }
        int nr = ActiveRow + dr, nc = ActiveCol + dc;
        if (dc != 0)
        {
            if (nc > sel.C2) { nc = sel.C1; nr++; }
            if (nc < sel.C1) { nc = sel.C2; nr--; }
            if (nr > sel.R2) nr = sel.R1;
            if (nr < sel.R1) nr = sel.R2;
        }
        else
        {
            if (nr > sel.R2) { nr = sel.R1; nc++; }
            if (nr < sel.R1) { nr = sel.R2; nc--; }
            if (nc > sel.C2) nc = sel.C1;
            if (nc < sel.C1) nc = sel.C2;
        }
        ActiveRow = nr;
        ActiveCol = nc;
        InvalidateVisual();
        SelectionChanged?.Invoke();
        EnsureVisible(ActiveRow, ActiveCol);
    }

    // ================================================================ cizim

    protected override void OnRender(DrawingContext dc)
    {
        double w = ActualWidth, h = ActualHeight;
        dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, w, h));
        if (_sheet == null) return;

        dc.PushTransform(new ScaleTransform(Zoom, Zoom));
        double lw = w / Zoom, lh = h / Zoom;
        double rhw = RowHeaderW;
        double dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;

        // Gorunen satir ve sutunlar
        var vRows = new List<(int I, double Pos, double Size)>();
        {
            double y = HeaderH;
            for (int r = TopRow; r <= MaxRow && y < lh; r++)
            {
                double s = Rows.Size(r);
                if (s > 0) vRows.Add((r, y, s));
                y += s;
            }
        }
        var vCols = new List<(int I, double Pos, double Size)>();
        {
            double x = rhw;
            for (int c = LeftCol; c <= MaxCol && x < lw; c++)
            {
                double s = Cols.Size(c);
                if (s > 0) vCols.Add((c, x, s));
                x += s;
            }
        }
        if (vRows.Count == 0 || vCols.Count == 0) { dc.Pop(); return; }
        int r1 = vRows[0].I, r2 = vRows[^1].I, c1 = vCols[0].I, c2 = vCols[^1].I;
        double right = vCols[^1].Pos + vCols[^1].Size, bottom = vRows[^1].Pos + vRows[^1].Size;

        dc.PushClip(new RectangleGeometry(new Rect(rhw, HeaderH, Math.Max(0, lw - rhw), Math.Max(0, lh - HeaderH))));

        // Izgara cizgileri
        if (_sheet.ShowGridLines)
        {
            foreach (var (_, x, s) in vCols) dc.DrawLine(GridPen, new Point(x + s - 0.5, HeaderH), new Point(x + s - 0.5, bottom));
            foreach (var (_, y, s) in vRows) dc.DrawLine(GridPen, new Point(rhw, y + s - 0.5), new Point(right, y + s - 0.5));
        }

        // Gorunur birlesik araliklar
        var merges = new List<CellRange>();
        foreach (var m in _sheet.MergedRanges)
        {
            var a = m.RangeAddress;
            var cr = new CellRange(a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber, a.LastAddress.RowNumber, a.LastAddress.ColumnNumber);
            if (cr.R2 < r1 || cr.R1 > r2 || cr.C2 < c1 || cr.C1 > c2) continue;
            merges.Add(cr);
        }
        bool Covered(int r, int c)
        {
            foreach (var m in merges) if (m.Contains(r, c) && !(r == m.R1 && c == m.C1)) return true;
            return false;
        }
        CellRange? MergeFor(int r, int c)
        {
            foreach (var m in merges) if (m.R1 == r && m.C1 == c) return m;
            return null;
        }

        var wb = _sheet.Workbook;
        var used = _sheet.Range(r1, c1, r2, c2).CellsUsed(XLCellsUsedOptions.All).ToList();
        // Birlesik araligin sol ust hucresi gorunur alanin disinda kalabilir: onu da ekle.
        foreach (var m in merges)
            if (m.R1 < r1 || m.C1 < c1) used.Add(_sheet.Cell(m.R1, m.C1));

        var filled = new HashSet<(int, int)>();
        foreach (var cell in used)
        {
            int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
            if (Covered(r, c)) continue;
            var rect = AreaRect(r, c, MergeFor(r, c));
            var st = cell.Style;
            Color? fill = null;
            if (st.Fill.PatternType != XLFillPatternValues.None)
                fill = SheetDisplay.ToWpf(st.Fill.PatternType == XLFillPatternValues.Solid ? st.Fill.BackgroundColor : st.Fill.PatternColor, wb);
            if (fill != null) dc.DrawRectangle(SheetDisplay.Brush(fill.Value), null, rect);
            else if (MergeFor(r, c) != null) dc.DrawRectangle(Brushes.White, null, rect);
        }
        // Bos sol ust hucreli birlesik araliklar da ic cizgileri ortsun
        foreach (var m in merges)
        {
            if (used.Any(u => u.Address.RowNumber == m.R1 && u.Address.ColumnNumber == m.C1)) continue;
            dc.DrawRectangle(Brushes.White, null, AreaRect(m.R1, m.C1, m));
        }

        // Metinler: once dolu hucre kumesi (tasma icin)
        var texts = new List<(IXLCell Cell, string Text)>();
        foreach (var cell in used)
        {
            int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
            if (Covered(r, c)) continue;
            if (cell.IsEmpty(XLCellsUsedOptions.Contents)) continue;
            var t = SheetDisplay.Text(cell);
            if (t.Length == 0) continue;
            filled.Add((r, c));
            texts.Add((cell, t));
        }

        foreach (var (cell, text) in texts)
        {
            int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
            var merge = MergeFor(r, c);
            var rect = AreaRect(r, c, merge);
            DrawCellText(dc, cell, text, rect, merge != null, filled, c2, dpi);
        }

        // Kenarliklar
        foreach (var cell in used)
        {
            int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
            if (Covered(r, c)) continue;
            DrawBorders(dc, cell, AreaRect(r, c, MergeFor(r, c)), wb);
        }

        DrawSelection(dc, r1, r2, c1, c2, lw, lh);
        dc.Pop(); // clip

        DrawHeaders(dc, vRows, vCols, lw, lh, rhw, dpi);
        dc.Pop(); // zoom
    }

    private Rect AreaRect(int r, int c, CellRange? merge)
    {
        if (merge is { } m)
        {
            double x1 = ColX(m.C1), y1 = RowY(m.R1);
            return new Rect(x1, y1, ColX(m.C2 + 1) - x1, RowY(m.R2 + 1) - y1);
        }
        return new Rect(ColX(c), RowY(r), Cols.Size(c), Rows.Size(r));
    }

    private Typeface Face(string name, bool bold, bool italic)
    {
        var key = (name, bold, italic);
        if (!_typefaces.TryGetValue(key, out var tf))
        {
            tf = new Typeface(new FontFamily(string.IsNullOrWhiteSpace(name) ? "Calibri" : name),
                italic ? FontStyles.Italic : FontStyles.Normal, bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
            _typefaces[key] = tf;
        }
        return tf;
    }

    /// Hucrenin gorunen metninin piksel genisligi (zoom oncesi) + kenar boslugu.
    public double MeasureCell(IXLCell cell)
    {
        var text = SheetDisplay.Text(cell);
        if (text.Length == 0) return 0;
        var font = cell.Style.Font;
        var ft = new FormattedText(text, Tr, FlowDirection.LeftToRight, Face(font.FontName, font.Bold, font.Italic),
            Math.Max(1, font.FontSize) * 96.0 / 72.0, Brushes.Black, VisualTreeHelper.GetDpi(this).PixelsPerDip);
        return Math.Ceiling(ft.Width) + 6;
    }

    private void DrawCellText(DrawingContext dc, IXLCell cell, string text, Rect rect, bool merged, HashSet<(int, int)> filled, int lastVisCol, double dpi)
    {
        var st = cell.Style;
        var font = st.Font;
        var color = SheetDisplay.ToWpf(font.FontColor, cell.Worksheet.Workbook) ?? Colors.Black;
        double size = Math.Max(1, font.FontSize) * 96.0 / 72.0;
        XLCellValue v;
        try { v = cell.Value; } catch { v = Blank.Value; }
        var align = SheetDisplay.Alignment(cell, st.Alignment.Horizontal, v.Type);
        bool wrap = st.Alignment.WrapText;
        bool isNumber = v.Type is XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan;

        var ft = new FormattedText(text, Tr, FlowDirection.LeftToRight, Face(font.FontName, font.Bold, font.Italic), size,
            SheetDisplay.Brush(color), dpi) { TextAlignment = TextAlignment.Left };
        if (font.Underline != XLFontUnderlineValues.None) ft.SetTextDecorations(TextDecorations.Underline);
        if (font.Strikethrough) ft.SetTextDecorations(TextDecorations.Strikethrough);

        const double pad = 3;
        double indent = st.Alignment.Indent * 9;
        var clip = rect;
        if (wrap)
        {
            ft.MaxTextWidth = Math.Max(1, rect.Width - 2 * pad);
            ft.TextAlignment = align;
        }
        else if (isNumber && ft.Width > rect.Width - 3 && !merged && v.IsNumber && SheetDisplay.IsGeneral(st.NumberFormat) &&
                 SheetDisplay.FitGeneral(v.GetNumber(), t => new FormattedText(t, Tr, FlowDirection.LeftToRight, Face(font.FontName, font.Bold, font.Italic), size, Brushes.Black, dpi).Width, rect.Width - 3) is { } fitted)
        {
            ft = new FormattedText(fitted, Tr, FlowDirection.LeftToRight, Face(font.FontName, font.Bold, font.Italic), size, SheetDisplay.Brush(color), dpi);
        }
        else if (isNumber && ft.Width > rect.Width - 3 && !merged)
        {
            // Sigmayan sayi Excel'deki gibi ### olarak gosterilir.
            ft = new FormattedText(new string('#', Math.Max(1, (int)((rect.Width - 2 * pad) / Math.Max(1, size * 0.55)))), Tr,
                FlowDirection.LeftToRight, Face(font.FontName, font.Bold, font.Italic), size, SheetDisplay.Brush(color), dpi);
        }
        else if (!merged && ft.Width > rect.Width - 2 * pad && align == TextAlignment.Left)
        {
            // Metin sagdaki bos hucrelere tasar.
            int r = cell.Address.RowNumber;
            int c = cell.Address.ColumnNumber + 1;
            double extra = 0;
            while (c <= lastVisCol + 1 && !filled.Contains((r, c)) && ft.Width > rect.Width + extra - 2 * pad)
            {
                extra += Cols.Size(c);
                c++;
            }
            clip = new Rect(rect.X, rect.Y, rect.Width + extra, rect.Height);
            // Tasan metnin altindaki izgara cizgilerini ort.
            if (extra > 0) dc.DrawRectangle(Brushes.White, null, new Rect(rect.Right, rect.Y + 0.5, extra - 1, rect.Height - 1.5));
        }

        double x = align switch
        {
            TextAlignment.Right when !wrap => rect.Right - Math.Min(pad, Math.Max(1.5, (rect.Width - ft.Width) / 2)) - ft.Width - indent,
            TextAlignment.Center when !wrap => rect.X + (rect.Width - ft.Width) / 2,
            _ => rect.X + pad + indent,
        };
        double th = ft.Height;
        double y = st.Alignment.Vertical switch
        {
            XLAlignmentVerticalValues.Top => rect.Y + 1,
            XLAlignmentVerticalValues.Center => rect.Y + (rect.Height - th) / 2,
            _ => rect.Bottom - th - 2,
        };
        if (!wrap && y < rect.Y) y = rect.Y + (rect.Height - th) / 2; // satir yaziya gore alcaksa ortala

        dc.PushClip(new RectangleGeometry(new Rect(clip.X, clip.Y, Math.Max(0, clip.Width - 1), Math.Max(0, clip.Height - 1))));
        dc.DrawText(ft, new Point(x, y));
        dc.Pop();
    }

    private static void DrawBorders(DrawingContext dc, IXLCell cell, Rect rect, IXLWorkbook wb)
    {
        var b = cell.Style.Border;
        void Edge(XLBorderStyleValues style, XLColor color, Point a, Point z)
        {
            if (style == XLBorderStyleValues.None) return;
            double w = style switch
            {
                XLBorderStyleValues.Medium or XLBorderStyleValues.MediumDashed or XLBorderStyleValues.MediumDashDot => 2,
                XLBorderStyleValues.Thick or XLBorderStyleValues.Double => 3,
                _ => 1,
            };
            var pen = new Pen(SheetDisplay.Brush(SheetDisplay.ToWpf(color, wb) ?? Colors.Black), w);
            if (style is XLBorderStyleValues.Dashed or XLBorderStyleValues.MediumDashed) pen.DashStyle = DashStyles.Dash;
            else if (style is XLBorderStyleValues.Dotted or XLBorderStyleValues.Hair) pen.DashStyle = DashStyles.Dot;
            dc.DrawLine(pen, a, z);
        }
        double l = rect.Left - 0.5, t = rect.Top - 0.5, r = rect.Right - 0.5, bo = rect.Bottom - 0.5;
        Edge(b.TopBorder, b.TopBorderColor, new Point(l, t), new Point(r, t));
        Edge(b.BottomBorder, b.BottomBorderColor, new Point(l, bo), new Point(r, bo));
        Edge(b.LeftBorder, b.LeftBorderColor, new Point(l, t), new Point(l, bo));
        Edge(b.RightBorder, b.RightBorderColor, new Point(r, t), new Point(r, bo));
    }

    private void DrawSelection(DrawingContext dc, int r1, int r2, int c1, int c2, double lw, double lh)
    {
        var sel = Selection;
        // Gorunur alana kirpilmis secim dikdortgeni
        Rect SelRect(CellRange s)
        {
            double x1 = s.C1 < LeftCol ? RowHeaderW - 2 : ColX(s.C1);
            double y1 = s.R1 < TopRow ? HeaderH - 2 : RowY(s.R1);
            double x2 = s.C2 > c2 ? lw + 2 : ColX(s.C2 + 1);
            double y2 = s.R2 > r2 ? lh + 2 : RowY(s.R2 + 1);
            return new Rect(x1, y1, Math.Max(0, x2 - x1), Math.Max(0, y2 - y1));
        }
        if (sel.R2 >= TopRow && sel.C2 >= LeftCol)
        {
            var rect = SelRect(sel);
            if (!sel.IsSingle && MergeAt(sel.R1, sel.C1) != sel)
            {
                // Etkin hucre beyaz kalir, geri kalan secim yari saydam boyanir.
                var active = CellRect(ActiveRow, ActiveCol);
                var activeL = new Rect(active.X / Zoom, active.Y / Zoom, active.Width / Zoom, active.Height / Zoom);
                var geo = new CombinedGeometry(GeometryCombineMode.Exclude, new RectangleGeometry(rect), new RectangleGeometry(activeL));
                dc.DrawGeometry(SelFill, null, geo);
            }
            dc.DrawRectangle(null, AccentPen, new Rect(rect.X - 0.5, rect.Y - 0.5, rect.Width, rect.Height));
            // Doldurma tutamaci
            if (sel.R2 <= r2 && sel.C2 <= c2)
                dc.DrawRectangle(Accent, new Pen(Brushes.White, 1), new Rect(rect.Right - 3.5, rect.Bottom - 3.5, 6, 6));
        }
        if (CopyMarquee is { } cm && cm.R2 >= TopRow && cm.C2 >= LeftCol)
        {
            var rect = SelRect(cm);
            dc.DrawRectangle(null, CopyPen, new Rect(rect.X + 0.5, rect.Y + 0.5, Math.Max(0, rect.Width - 2), Math.Max(0, rect.Height - 2)));
        }
    }

    private void DrawHeaders(DrawingContext dc, List<(int I, double Pos, double Size)> vRows, List<(int I, double Pos, double Size)> vCols,
        double lw, double lh, double rhw, double dpi)
    {
        var sel = Selection;
        var face = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var faceSel = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

        dc.DrawRectangle(HeaderBg, null, new Rect(0, 0, lw, HeaderH));
        dc.DrawRectangle(HeaderBg, null, new Rect(0, 0, rhw, lh));

        dc.PushClip(new RectangleGeometry(new Rect(rhw, 0, Math.Max(0, lw - rhw), HeaderH)));
        foreach (var (c, x, s) in vCols)
        {
            bool on = c >= sel.C1 && c <= sel.C2;
            if (on) dc.DrawRectangle(HeaderSelBg, null, new Rect(x, 0, s, HeaderH));
            var ft = new FormattedText(XLHelper.GetColumnLetterFromNumber(c), Tr, FlowDirection.LeftToRight, on ? faceSel : face, 12,
                on ? HeaderSelText : HeaderText, dpi);
            if (ft.Width < s - 2) dc.DrawText(ft, new Point(x + (s - ft.Width) / 2, (HeaderH - ft.Height) / 2));
            dc.DrawLine(HeaderPen, new Point(x + s - 0.5, 0), new Point(x + s - 0.5, HeaderH));
            if (on) dc.DrawLine(AccentPen, new Point(x, HeaderH - 1), new Point(x + s, HeaderH - 1));
        }
        dc.Pop();

        dc.PushClip(new RectangleGeometry(new Rect(0, HeaderH, rhw, Math.Max(0, lh - HeaderH))));
        foreach (var (r, y, s) in vRows)
        {
            bool on = r >= sel.R1 && r <= sel.R2;
            if (on) dc.DrawRectangle(HeaderSelBg, null, new Rect(0, y, rhw, s));
            var ft = new FormattedText(r.ToString(CultureInfo.InvariantCulture), Tr, FlowDirection.LeftToRight, on ? faceSel : face, 12,
                on ? HeaderSelText : HeaderText, dpi);
            if (ft.Height < s + 4) dc.DrawText(ft, new Point(rhw - 6 - ft.Width, y + (s - ft.Height) / 2));
            dc.DrawLine(HeaderPen, new Point(0, y + s - 0.5), new Point(rhw, y + s - 0.5));
            if (on) dc.DrawLine(AccentPen, new Point(rhw - 1, y), new Point(rhw - 1, y + s));
        }
        dc.Pop();

        dc.DrawLine(HeaderPen, new Point(0, HeaderH - 0.5), new Point(lw, HeaderH - 0.5));
        dc.DrawLine(HeaderPen, new Point(rhw - 0.5, 0), new Point(rhw - 0.5, lh));
        // Sol ust kose: tumunu sec ucgeni
        var tri = new StreamGeometry();
        using (var g = tri.Open())
        {
            g.BeginFigure(new Point(rhw - 4, HeaderH - 14), true, true);
            g.LineTo(new Point(rhw - 4, HeaderH - 4), false, false);
            g.LineTo(new Point(rhw - 14, HeaderH - 4), false, false);
        }
        dc.DrawGeometry(Frozen(Color.FromRgb(0xB4, 0xB6, 0xC0)), null, tri);
    }

    // ================================================================ fare

    private enum Zone { None, Corner, ColHeader, RowHeader, Cell }

    private (Zone Zone, int Row, int Col) HitTest(Point p)
    {
        double x = p.X / Zoom, y = p.Y / Zoom, rhw = RowHeaderW;
        int c = Cols.IndexAt(Cols.Offset(LeftCol) + (x - rhw));
        int r = Rows.IndexAt(Rows.Offset(TopRow) + (y - HeaderH));
        c = Math.Clamp(c, 1, MaxCol);
        r = Math.Clamp(r, 1, MaxRow);
        if (x < rhw && y < HeaderH) return (Zone.Corner, r, c);
        if (y < HeaderH) return (Zone.ColHeader, r, c);
        if (x < rhw) return (Zone.RowHeader, r, c);
        return (Zone.Cell, r, c);
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        Focus();
        var p = e.GetPosition(this);
        var (zone, r, c) = HitTest(p);
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        switch (zone)
        {
            case Zone.Corner:
                SelectRange(new CellRange(1, 1, MaxRow, MaxCol));
                return;
            case Zone.ColHeader:
                if (shift) SelectRange(CellRange.Of(1, Math.Min(ActiveCol, c), MaxRow, Math.Max(ActiveCol, c)));
                else { ActiveRow = 1; ActiveCol = c; _extRow = MaxRow; _extCol = c; InvalidateVisual(); SelectionChanged?.Invoke(); }
                _drag = DragMode.Cols;
                break;
            case Zone.RowHeader:
                if (shift) SelectRange(CellRange.Of(Math.Min(ActiveRow, r), 1, Math.Max(ActiveRow, r), MaxCol));
                else { ActiveRow = r; ActiveCol = 1; _extRow = r; _extCol = MaxCol; InvalidateVisual(); SelectionChanged?.Invoke(); }
                _drag = DragMode.Rows;
                break;
            case Zone.Cell:
                if (e.ClickCount == 2)
                {
                    Select(r, c);
                    EditRequested?.Invoke(null);
                    e.Handled = true;
                    return;
                }
                Select(r, c, shift);
                _drag = DragMode.Cells;
                break;
        }
        _lastMouse = p;
        CaptureMouse();
        _autoScroll.Start();
        e.Handled = true;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_drag == DragMode.None || e.LeftButton != MouseButtonState.Pressed) return;
        _lastMouse = e.GetPosition(this);
        DragTo(_lastMouse);
    }

    private void DragTo(Point p)
    {
        var (_, r, c) = HitTest(new Point(Math.Max(p.X, RowHeaderW * Zoom + 1), Math.Max(p.Y, HeaderH * Zoom + 1)));
        switch (_drag)
        {
            case DragMode.Cells:
                if (r != _extRow || c != _extCol) Select(r, c, extend: true);
                break;
            case DragMode.Cols:
                if (c != _extCol) { _extCol = c; _extRow = MaxRow; InvalidateVisual(); SelectionChanged?.Invoke(); }
                break;
            case DragMode.Rows:
                if (r != _extRow) { _extRow = r; _extCol = MaxCol; InvalidateVisual(); SelectionChanged?.Invoke(); }
                break;
        }
    }

    private void AutoScrollTick()
    {
        if (_drag == DragMode.None) { _autoScroll.Stop(); return; }
        int dr = 0, dc = 0;
        if (_lastMouse.Y > ActualHeight) dr = 1;
        else if (_lastMouse.Y < HeaderH * Zoom && _drag != DragMode.Cols) dr = -1;
        if (_lastMouse.X > ActualWidth) dc = 1;
        else if (_lastMouse.X < RowHeaderW * Zoom && _drag != DragMode.Rows) dc = -1;
        if (dr == 0 && dc == 0) return;
        ScrollTo(TopRow + dr, LeftCol + dc);
        DragTo(_lastMouse);
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonUp(e);
        _drag = DragMode.None;
        _autoScroll.Stop();
        ReleaseMouseCapture();
    }

    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        base.OnMouseWheel(e);
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            ZoomRequested?.Invoke(Zoom + (e.Delta > 0 ? 0.1 : -0.1));
        }
        else if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
        {
            ScrollTo(TopRow, LeftCol + (e.Delta > 0 ? -3 : 3));
        }
        else
        {
            ScrollTo(TopRow + (e.Delta > 0 ? -3 : 3), LeftCol);
        }
        e.Handled = true;
    }

    // ================================================================ klavye

    protected override void OnTextInput(TextCompositionEventArgs e)
    {
        base.OnTextInput(e);
        if (string.IsNullOrEmpty(e.Text) || char.IsControl(e.Text[0])) return;
        EditRequested?.Invoke(e.Text);
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
        var key = e.Key == Key.System ? e.SystemKey : e.Key;

        void Go(int r, int c)
        {
            if (shift) Select(r, c, extend: true);
            else Select(r, c);
            EnsureVisible(shift ? _extRow : ActiveRow, shift ? _extCol : ActiveCol);
        }

        int fromR = shift ? _extRow : ActiveRow, fromC = shift ? _extCol : ActiveCol;
        // Birlesik hucreden cikarken araligin kenarindan devam et.
        if (!shift && MergeAt(ActiveRow, ActiveCol) is { } mg)
        {
            if (key == Key.Down) fromR = mg.R2;
            if (key == Key.Right) fromC = mg.C2;
        }

        switch (key)
        {
            case Key.Up: Go(ctrl ? DataEdge(fromR, fromC, -1, 0) : fromR - 1, fromC); break;
            case Key.Down: Go(ctrl ? DataEdge(fromR, fromC, 1, 0) : fromR + 1, fromC); break;
            case Key.Left: Go(fromR, ctrl ? DataEdge(fromR, fromC, 0, -1) : fromC - 1); break;
            case Key.Right: Go(fromR, ctrl ? DataEdge(fromR, fromC, 0, 1) : fromC + 1); break;
            case Key.PageDown: ScrollTo(TopRow + PageRows, LeftCol); Go(fromR + PageRows, fromC); break;
            case Key.PageUp: ScrollTo(TopRow - PageRows, LeftCol); Go(fromR - PageRows, fromC); break;
            case Key.Home:
                if (ctrl) Go(1, 1);
                else Go(fromR, 1);
                break;
            case Key.End when ctrl:
            {
                var last = _sheet?.LastCellUsed(XLCellsUsedOptions.Contents);
                Go(last?.Address.RowNumber ?? 1, last?.Address.ColumnNumber ?? 1);
                break;
            }
            case Key.Enter: MoveActive(shift ? -1 : 1, 0); break;
            case Key.Tab: MoveActive(0, shift ? -1 : 1); break;
            case Key.F2: EditRequested?.Invoke(null); break;
            case Key.Delete: DeleteRequested?.Invoke(); break;
            case Key.Back: EditRequested?.Invoke(""); break;
            case Key.A when ctrl: SelectRange(new CellRange(1, 1, MaxRow, MaxCol)); break;
            case Key.Space when ctrl: SelectRange(new CellRange(1, ActiveCol, MaxRow, ActiveCol)); break;
            case Key.Space when shift: SelectRange(new CellRange(ActiveRow, 1, ActiveRow, MaxCol)); break;
            default: return;
        }
        e.Handled = true;
    }

    /// Ctrl+ok: veri blogunun kenarina atla (Excel davranisi).
    private int DataEdge(int r, int c, int dr, int dc)
    {
        if (_sheet == null) return dr != 0 ? r + dr : c + dc;
        var lastUsed = _sheet.LastCellUsed(XLCellsUsedOptions.Contents);
        int lastR = lastUsed?.Address.RowNumber ?? 0, lastC = lastUsed?.Address.ColumnNumber ?? 0;
        bool Has(int rr, int cc) => rr >= 1 && cc >= 1 && rr <= MaxRow && cc <= MaxCol && !_sheet.Cell(rr, cc).IsEmpty(XLCellsUsedOptions.Contents);
        int edgeR = dr > 0 ? MaxRow : 1, edgeC = dc > 0 ? MaxCol : 1;

        int nr = r + dr, nc = c + dc;
        if (nr < 1 || nc < 1 || nr > MaxRow || nc > MaxCol) return dr != 0 ? r : c;
        if (Has(r, c) && Has(nr, nc))
        {
            while (Has(nr + dr, nc + dc)) { nr += dr; nc += dc; }
            return dr != 0 ? nr : nc;
        }
        // Bos alanda: bir sonraki dolu hucreye ya da sayfa kenarina git.
        while (!Has(nr, nc))
        {
            if ((dr > 0 && nr > lastR) || (dc > 0 && nc > lastC)) return dr != 0 ? edgeR : edgeC;
            if ((dr < 0 && nr <= 1) || (dc < 0 && nc <= 1)) return dr != 0 ? edgeR : edgeC;
            nr += dr;
            nc += dc;
        }
        return dr != 0 ? nr : nc;
    }

    protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
    {
        base.OnRenderSizeChanged(sizeInfo);
        ScrollChanged?.Invoke();
    }
}
