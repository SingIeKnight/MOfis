using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MOfis.Common;

namespace MOfis.Writer;

/// Word'deki gibi cetvel. Yatay: kenar bosluklari, cm olcegi, girinti isaretcileri (ilk satir, asili,
/// sol, sag). Dikey: ust/alt kenar bosluklari. Isaretciler ve kenar boslugu sinirlari surukelenebilir.
/// Tum uzunluklar sayfa koordinatinda (piksel, olceklenmemis); Origin ve Scale ekrana yerlestirir.
public sealed class Ruler : FrameworkElement
{
    public const string FirstLine = "first", Hanging = "hanging", LeftBox = "leftbox", RightIndent = "right";
    public const string MarginStart = "mstart", MarginEnd = "mend";

    private const double BarTop = 3, BarBottom = 19;
    private static readonly double CmPx = Units.CmToPx(1);

    private static readonly Brush MarginBrush = Frozen(Color.FromRgb(0xD3, 0xD5, 0xDD));
    private static readonly Brush TextAreaBrush = Brushes.White;
    private static readonly Brush TickBrush = Frozen(Color.FromRgb(0x6B, 0x70, 0x7A));
    private static readonly Pen TickPen = FrozenPen(TickBrush, 1);
    private static readonly Pen EdgePen = FrozenPen(Frozen(Color.FromRgb(0xB4, 0xB7, 0xC2)), 1);
    private static readonly Pen MarkerPen = FrozenPen(Frozen(Color.FromRgb(0x4B, 0x50, 0x5A)), 1);
    private static readonly Brush MarkerFill = Frozen(Color.FromRgb(0xF4, 0xF5, 0xF8));
    private static readonly Brush ActiveFill = Frozen(Color.FromRgb(0xC7, 0xCB, 0xF5));
    private static readonly Typeface LabelFace = new("Segoe UI");

    public bool Vertical { get; set; }

    /// Sayfa kenarinin bu ogedeki konumu (ekran pikseli).
    public double Origin { get; private set; }
    public double Scale { get; private set; } = 1;
    public double PageLength { get; private set; }
    public double MarginStartLength { get; private set; }
    public double MarginEndLength { get; private set; }

    /// Girintiler, yazi alaninin basindan itibaren (yalnizca yatay cetvelde).
    public bool ShowIndents { get; private set; }
    public double IndentBase { get; private set; }
    public double LeftIndent { get; private set; }
    public double FirstLineIndent { get; private set; }
    public double RightIndentLength { get; private set; }

    private string? _drag;
    private string? _hover;

    // Gelistirme: fare simulasyonu icin sahte konum.
    internal Point? TestPosition;
    private Point Pos(MouseEventArgs e) => TestPosition ?? e.GetPosition(this);

    internal void SimulateMouse(string kind, Point p)
    {
        TestPosition = p;
        switch (kind)
        {
            case "move": OnMouseMove(new MouseEventArgs(Mouse.PrimaryDevice, 0) { RoutedEvent = Mouse.MouseMoveEvent }); break;
            case "down": OnMouseLeftButtonDown(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left) { RoutedEvent = UIElement.MouseLeftButtonDownEvent }); break;
            case "up": OnMouseLeftButtonUp(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left) { RoutedEvent = UIElement.MouseLeftButtonUpEvent }); break;
        }
    }

    public event Action<string>? DragStarted;
    public event Action<string, double>? Dragging;
    public event Action<string>? DragCompleted;

    public Ruler()
    {
        ClipToBounds = true;
        SnapsToDevicePixels = true;
        ToolTipService.SetInitialShowDelay(this, 600);
    }

    public void SetPage(double origin, double scale, double pageLength, double marginStart, double marginEnd)
    {
        if (Origin == origin && Scale == scale && PageLength == pageLength && MarginStartLength == marginStart && MarginEndLength == marginEnd) return;
        Origin = origin;
        Scale = scale;
        PageLength = pageLength;
        MarginStartLength = marginStart;
        MarginEndLength = marginEnd;
        InvalidateVisual();
    }

    public void SetIndents(bool show, double indentBase, double left, double firstLine, double right)
    {
        if (ShowIndents == show && IndentBase == indentBase && LeftIndent == left && FirstLineIndent == firstLine && RightIndentLength == right) return;
        ShowIndents = show;
        IndentBase = indentBase;
        LeftIndent = left;
        FirstLineIndent = firstLine;
        RightIndentLength = right;
        InvalidateVisual();
    }

    // ---------------------------------------------------------------- konumlar (ekran pikseli)

    internal double TextStart => Origin + MarginStartLength * Scale;
    internal double TextEnd => Origin + (PageLength - MarginEndLength) * Scale;
    internal double LeftX => TextStart + (IndentBase + LeftIndent) * Scale;
    internal double FirstX => TextStart + (IndentBase + LeftIndent + FirstLineIndent) * Scale;
    internal double RightX => TextEnd - RightIndentLength * Scale;

    // ---------------------------------------------------------------- cizim

    protected override void OnRender(DrawingContext dc)
    {
        double length = Vertical ? ActualHeight : ActualWidth;
        double thickness = Vertical ? ActualWidth : ActualHeight;
        if (length <= 0 || PageLength <= 0) return;

        Rect Span(double a, double b) => Vertical
            ? new Rect(BarTop, Math.Min(a, b), BarBottom - BarTop, Math.Abs(b - a))
            : new Rect(Math.Min(a, b), BarTop, Math.Abs(b - a), BarBottom - BarTop);

        double pageStart = Origin, pageEnd = Origin + PageLength * Scale;
        dc.DrawRectangle(MarginBrush, null, Span(pageStart, pageEnd));
        dc.DrawRectangle(TextAreaBrush, null, Span(TextStart, TextEnd));
        dc.DrawRectangle(null, EdgePen, Span(pageStart, pageEnd));

        // cm olcegi: yazi alaninin basi 0; her 0,25 cm bir cizgi, her cm bir sayi.
        double quarter = CmPx * Scale / 4;
        if (quarter > 1.5)
        {
            int from = -(int)Math.Ceiling((TextStart - pageStart) / quarter);
            int to = (int)Math.Ceiling((pageEnd - TextStart) / quarter);
            double mid = (BarTop + BarBottom) / 2;
            double dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            for (int k = from; k <= to; k++)
            {
                if (k == 0) continue;
                double pos = TextStart + k * quarter;
                if (pos <= pageStart + 1 || pos >= pageEnd - 1) continue;
                if (k % 4 == 0)
                {
                    var text = new FormattedText(Math.Abs(k / 4).ToString(CultureInfo.CurrentCulture), CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight, LabelFace, 9, TickBrush, dpi);
                    var at = Vertical
                        ? new Point(mid - text.Width / 2, pos - text.Height / 2)
                        : new Point(pos - text.Width / 2, mid - text.Height / 2);
                    dc.DrawText(text, at);
                }
                else
                {
                    double half = k % 2 == 0 ? 3 : 1.2;
                    if (Vertical) dc.DrawLine(TickPen, new Point(mid - half, Snap(pos)), new Point(mid + half, Snap(pos)));
                    else dc.DrawLine(TickPen, new Point(Snap(pos), mid - half), new Point(Snap(pos), mid + half));
                }
            }
        }

        if (!Vertical && ShowIndents)
        {
            double fx = FirstX, lx = LeftX, rx = RightX;
            // Ilk satir girintisi: ustte asagi bakan ucgen
            Triangle(dc, fx, BarTop - 1, BarTop + 5, _hover == FirstLine || _drag == FirstLine);
            // Asili girinti: altta yukari bakan ucgen; altinda sol girinti kutusu
            Triangle(dc, lx, BarBottom + 1, BarBottom - 5, _hover == Hanging || _drag == Hanging);
            var box = new Rect(Math.Round(lx) - 4.5, BarBottom + 1.5, 9, 3.5);
            dc.DrawRectangle(_hover == LeftBox || _drag == LeftBox ? ActiveFill : MarkerFill, MarkerPen, box);
            // Sag girinti
            Triangle(dc, rx, BarBottom + 1, BarBottom - 5, _hover == RightIndent || _drag == RightIndent);
        }
    }

    private static void Triangle(DrawingContext dc, double x, double baseY, double tipY, bool active)
    {
        x = Math.Round(x) + 0.5;
        var g = new StreamGeometry();
        using (var ctx = g.Open())
        {
            ctx.BeginFigure(new Point(x - 4.5, baseY), true, true);
            ctx.LineTo(new Point(x + 4.5, baseY), true, false);
            ctx.LineTo(new Point(x, tipY), true, false);
        }
        g.Freeze();
        dc.DrawGeometry(active ? ActiveFill : MarkerFill, MarkerPen, g);
    }

    private static double Snap(double v) => Math.Round(v) + 0.5;

    // ---------------------------------------------------------------- fare

    private string? HitTest(Point p)
    {
        double along = Vertical ? p.Y : p.X;
        double across = Vertical ? p.X : p.Y;
        if (!Vertical && ShowIndents)
        {
            // Ust serit (<=9): ilk satir; alt serit (>=13): asili / sag; en alt (>=20): sol kutu.
            // Ortadaki serit (9-13) kenar boslugu sinirini yakalamak icin bos birakilir (Word gibi).
            if (Math.Abs(along - FirstX) <= 5 && across <= BarTop + 6) return FirstLine;
            if (Math.Abs(along - LeftX) <= 5 && across >= BarBottom + 1) return LeftBox;
            if (Math.Abs(along - LeftX) <= 5 && across >= BarBottom - 6) return Hanging;
            if (Math.Abs(along - RightX) <= 5 && across >= BarBottom - 6) return RightIndent;
        }
        if (Math.Abs(along - TextStart) <= 4) return MarginStart;
        if (Math.Abs(along - TextEnd) <= 4) return MarginEnd;
        return null;
    }

    private static string? TipFor(string? target, bool vertical) => target switch
    {
        FirstLine => "İlk satır girintisi",
        Hanging => "Asılı girinti",
        LeftBox => "Sol girinti",
        RightIndent => "Sağ girinti",
        MarginStart => vertical ? "Üst kenar boşluğu" : "Sol kenar boşluğu",
        MarginEnd => vertical ? "Alt kenar boşluğu" : "Sağ kenar boşluğu",
        _ => null,
    };

    /// Fare konumunu sayfa koordinatina cevirir; Alt basili degilse 0,25 cm'ye yaslar.
    private double ToPage(Point p)
    {
        double along = Vertical ? p.Y : p.X;
        double page = (along - Origin) / Scale;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) return page;
        double step = CmPx / 4;
        double textStartPage = MarginStartLength;
        return _drag is MarginStart or MarginEnd
            ? Math.Round(page / step) * step
            : textStartPage + Math.Round((page - textStartPage) / step) * step;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        var p = Pos(e);
        if (_drag != null)
        {
            Dragging?.Invoke(_drag, ToPage(p));
            return;
        }
        var hit = HitTest(p);
        if (hit != _hover)
        {
            _hover = hit;
            Cursor = hit == null ? null : (Vertical ? Cursors.SizeNS : Cursors.SizeWE);
            ToolTip = TipFor(hit, Vertical);
            InvalidateVisual();
        }
    }

    protected override void OnMouseLeave(MouseEventArgs e)
    {
        base.OnMouseLeave(e);
        if (_drag == null && _hover != null)
        {
            _hover = null;
            InvalidateVisual();
        }
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        var hit = HitTest(Pos(e));
        if (hit == null) return;
        _drag = hit;
        CaptureMouse();
        DragStarted?.Invoke(hit);
        e.Handled = true;
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonUp(e);
        EndDrag();
    }

    protected override void OnLostMouseCapture(MouseEventArgs e)
    {
        base.OnLostMouseCapture(e);
        EndDrag();
    }

    private void EndDrag()
    {
        if (_drag == null) return;
        var done = _drag;
        _drag = null;
        if (IsMouseCaptured) ReleaseMouseCapture();
        DragCompleted?.Invoke(done);
        InvalidateVisual();
    }

    protected override Size MeasureOverride(Size availableSize) =>
        Vertical ? new Size(22, 0) : new Size(0, 22);

    private static SolidColorBrush Frozen(Color c)
    {
        var b = new SolidColorBrush(c);
        b.Freeze();
        return b;
    }

    private static Pen FrozenPen(Brush b, double w)
    {
        var p = new Pen(b, w);
        p.Freeze();
        return p;
    }
}
