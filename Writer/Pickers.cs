using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MOfis.Writer;

/// Serit dugmesinin altinda acilan acilir paneller (renk paleti, tablo boyutu secici).
public static class Pickers
{
    private static readonly string[] ThemeBase =
        ["000000", "FFFFFF", "374151", "4338CA", "0F766E", "B45309", "BE123C", "6D28D9", "0369A1", "4D7C0F"];

    private static readonly string[] Standard =
        ["C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0"];

    public static readonly string[] Highlights =
        ["FFFF00", "00FF00", "00FFFF", "FF00FF", "0000FF", "FF0000", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0", "000000"];

    private static Popup MakePopup(UIElement anchor, UIElement content)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xC9, 0xCA, 0xD2)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8),
            Child = content,
            Effect = new System.Windows.Media.Effects.DropShadowEffect { BlurRadius = 10, ShadowDepth = 2, Opacity = 0.2, Direction = 270 },
            Margin = new Thickness(0, 0, 8, 8),
        };
        return new Popup
        {
            PlacementTarget = anchor,
            Placement = PlacementMode.Bottom,
            StaysOpen = false,
            AllowsTransparency = true,
            PopupAnimation = PopupAnimation.Fade,
            Child = border,
        };
    }

    /// Renk paleti. noneLabel verilirse en ustte "Otomatik"/"Renk yok" secenegi olur (null dondurur).
    public static void ShowColors(UIElement anchor, string? noneLabel, bool highlightSet, Action<Color?> picked)
    {
        var root = new StackPanel { Width = 10 * 22 };
        Popup? popup = null;

        void Pick(Color? c)
        {
            popup!.IsOpen = false;
            picked(c);
        }

        if (noneLabel != null)
        {
            var none = new Button { Content = noneLabel, Style = (Style)Application.Current.Resources["TileButton"], Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 0, 0, 6) };
            none.Click += (_, _) => Pick(null);
            root.Children.Add(none);
        }

        if (highlightSet)
        {
            root.Children.Add(Swatches(Highlights.Select(Parse), Pick));
        }
        else
        {
            root.Children.Add(Label("Tema renkleri"));
            var bases = ThemeBase.Select(Parse).ToArray();
            root.Children.Add(Swatches(bases, Pick));
            var space = new Border { Height = 4 };
            root.Children.Add(space);
            foreach (var amount in new[] { 0.8, 0.6, 0.4 })
                root.Children.Add(Swatches(bases.Select(c => Mix(c, Colors.White, amount)), Pick));
            foreach (var amount in new[] { 0.25, 0.5 })
                root.Children.Add(Swatches(bases.Select(c => Mix(c, Colors.Black, amount)), Pick));
            root.Children.Add(Label("Standart renkler"));
            root.Children.Add(Swatches(Standard.Select(Parse), Pick));
        }

        popup = MakePopup(anchor, root);
        popup.IsOpen = true;
    }

    private static TextBlock Label(string text) => new()
    {
        Text = text,
        FontSize = 11.5,
        FontWeight = FontWeights.SemiBold,
        Margin = new Thickness(0, 6, 0, 4),
        Foreground = (Brush)Application.Current.Resources["SubtleTextBrush"],
    };

    private static WrapPanel Swatches(IEnumerable<Color> colors, Action<Color?> pick)
    {
        var panel = new WrapPanel();
        foreach (var c in colors)
        {
            var sw = new Border
            {
                Width = 18,
                Height = 18,
                Margin = new Thickness(2),
                Background = new SolidColorBrush(c),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xC0, 0xC0, 0xC8)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                ToolTip = "#" + DocHelpers.ToHex(c),
            };
            sw.MouseEnter += (_, _) => sw.BorderBrush = Brushes.Black;
            sw.MouseLeave += (_, _) => sw.BorderBrush = new SolidColorBrush(Color.FromRgb(0xC0, 0xC0, 0xC8));
            sw.MouseLeftButtonUp += (_, _) => pick(c);
            panel.Children.Add(sw);
        }
        return panel;
    }

    private static Color Parse(string hex) => DocHelpers.ParseHex(hex) ?? Colors.Black;

    private static Color Mix(Color a, Color b, double t) => Color.FromRgb(
        (byte)Math.Round(a.R + (b.R - a.R) * t),
        (byte)Math.Round(a.G + (b.G - a.G) * t),
        (byte)Math.Round(a.B + (b.B - a.B) * t));

    /// Office'teki gibi kareler uzerinde gezinerek tablo boyutu secme.
    public static void ShowTableGrid(UIElement anchor, Action<int, int> picked)
    {
        const int cols = 10, rows = 8;
        var title = new TextBlock { Text = "Tablo ekle", Margin = new Thickness(0, 0, 0, 6), FontWeight = FontWeights.SemiBold };
        var grid = new UniformGrid { Columns = cols, Rows = rows };
        var cells = new Border[rows, cols];
        Popup? popup = null;
        var accent = (Brush)Application.Current.Resources["AccentLightBrush"];
        var accentLine = (Brush)Application.Current.Resources["AccentBrush"];
        var line = new SolidColorBrush(Color.FromRgb(0xC0, 0xC0, 0xC8));

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                int rr = r, cc = c;
                var b = new Border { Width = 18, Height = 18, Margin = new Thickness(1.5), BorderBrush = line, BorderThickness = new Thickness(1), Background = Brushes.White };
                b.MouseEnter += (_, _) =>
                {
                    title.Text = $"{cc + 1} × {rr + 1} Tablo";
                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            bool on = i <= rr && j <= cc;
                            cells[i, j].Background = on ? accent : Brushes.White;
                            cells[i, j].BorderBrush = on ? accentLine : line;
                        }
                    }
                };
                b.MouseLeftButtonUp += (_, _) =>
                {
                    popup!.IsOpen = false;
                    picked(rr + 1, cc + 1);
                };
                cells[r, c] = b;
                grid.Children.Add(b);
            }
        }

        var root = new StackPanel();
        root.Children.Add(title);
        root.Children.Add(grid);
        popup = MakePopup(anchor, root);
        popup.IsOpen = true;
    }
}

/// Paragraf dugmeleri icin kucuk cizgi ikonlari (hizalama, madde, numara, satir araligi).
public static class LineIcons
{
    private static Brush Ink => (Brush)Application.Current.Resources["TextBrush"];

    private static Canvas Canvas16() => new() { Width = 16, Height = 16, SnapsToDevicePixels = true };

    private static void Bar(Canvas c, double x, double y, double w)
    {
        var r = new Rectangle { Width = w, Height = 1.6, Fill = Ink };
        System.Windows.Controls.Canvas.SetLeft(r, x);
        System.Windows.Controls.Canvas.SetTop(r, y);
        c.Children.Add(r);
    }

    public static UIElement Align(TextAlignment a)
    {
        var c = Canvas16();
        double[] widths = a == TextAlignment.Justify ? [14, 14, 14, 14] : [14, 9, 14, 9];
        for (int i = 0; i < 4; i++)
        {
            double w = widths[i];
            double x = a switch
            {
                TextAlignment.Center => 1 + (14 - w) / 2,
                TextAlignment.Right => 15 - w,
                _ => 1,
            };
            Bar(c, x, 2 + i * 3.8, w);
        }
        return c;
    }

    public static UIElement Bullets()
    {
        var c = Canvas16();
        for (int i = 0; i < 3; i++)
        {
            var dot = new Ellipse { Width = 3, Height = 3, Fill = Ink };
            System.Windows.Controls.Canvas.SetLeft(dot, 0.5);
            System.Windows.Controls.Canvas.SetTop(dot, 2 + i * 5);
            c.Children.Add(dot);
            Bar(c, 5.5, 2.7 + i * 5, 10);
        }
        return c;
    }

    public static UIElement Numbers()
    {
        var c = Canvas16();
        for (int i = 0; i < 3; i++)
        {
            var t = new TextBlock { Text = (i + 1).ToString(), FontSize = 6.5, Foreground = Ink, FontWeight = FontWeights.SemiBold };
            System.Windows.Controls.Canvas.SetLeft(t, 0);
            System.Windows.Controls.Canvas.SetTop(t, -0.5 + i * 5);
            c.Children.Add(t);
            Bar(c, 5.5, 2.7 + i * 5, 10);
        }
        return c;
    }

    public static UIElement LineSpacing()
    {
        var c = Canvas16();
        var arrow = new System.Windows.Shapes.Path
        {
            Data = Geometry.Parse("M3,1 L5.5,4 L0.5,4 Z M3,15 L5.5,12 L0.5,12 Z M2.4,4 L3.6,4 L3.6,12 L2.4,12 Z"),
            Fill = Ink,
        };
        c.Children.Add(arrow);
        for (int i = 0; i < 4; i++) Bar(c, 7.5, 2 + i * 3.8, 8);
        return c;
    }
}

/// Imlecin bulundugu tabloda satir/sutun ekleme ve silme (birlesik hucresi olmayan tablolar icin).
public static class TableOps
{
    public static (Table table, TableRow row, TableCell cell)? Locate(TextPointer pos)
    {
        DependencyObject? cur = pos.Parent;
        TableCell? cell = null;
        while (cur != null)
        {
            if (cur is TableCell tc && cell == null) cell = tc;
            if (cur is Table t && cell != null && cell.Parent is TableRow row) return (t, row, cell);
            cur = cur is TextElement te ? te.Parent : null;
        }
        return null;
    }

    private static TableCell NewCellLike(TableCell model) => new(new Paragraph())
    {
        BorderBrush = model.BorderBrush,
        BorderThickness = model.BorderThickness,
        Padding = model.Padding,
    };

    public static void InsertRow(TextPointer pos, bool below)
    {
        if (Locate(pos) is not (_, var row, _)) return;
        var group = (TableRowGroup)row.Parent;
        var newRow = new TableRow();
        foreach (var c in row.Cells) newRow.Cells.Add(NewCellLike(c));
        int index = group.Rows.IndexOf(row);
        group.Rows.Insert(below ? index + 1 : index, newRow);
    }

    public static void InsertColumn(TextPointer pos, bool right)
    {
        if (Locate(pos) is not (var table, var row, var cell)) return;
        int index = row.Cells.IndexOf(cell) + (right ? 1 : 0);
        foreach (var g in table.RowGroups)
            foreach (var r in g.Rows)
                r.Cells.Insert(Math.Min(index, r.Cells.Count), NewCellLike(r.Cells.Count > 0 ? r.Cells[Math.Min(index, r.Cells.Count - 1)] : cell));
        if (table.Columns.Count > 0)
        {
            foreach (var col in table.Columns) col.Width = new GridLength(1, GridUnitType.Star);
            table.Columns.Insert(Math.Min(index, table.Columns.Count), new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        }
    }

    public static void DeleteRow(TextPointer pos)
    {
        if (Locate(pos) is not (var table, var row, _)) return;
        var group = (TableRowGroup)row.Parent;
        group.Rows.Remove(row);
        if (table.RowGroups.Sum(g => g.Rows.Count) == 0) DeleteTable(table);
    }

    public static void DeleteColumn(TextPointer pos)
    {
        if (Locate(pos) is not (var table, var row, var cell)) return;
        int index = row.Cells.IndexOf(cell);
        foreach (var g in table.RowGroups)
            foreach (var r in g.Rows)
                if (index < r.Cells.Count) r.Cells.RemoveAt(index);
        if (index < table.Columns.Count) table.Columns.RemoveAt(index);
        if (table.RowGroups.SelectMany(g => g.Rows).All(r => r.Cells.Count == 0)) DeleteTable(table);
    }

    public static void DeleteTable(Table table)
    {
        var siblings = table.SiblingBlocks;
        if (siblings == null) return;
        if (siblings.Count == 1) siblings.Add(new Paragraph());
        siblings.Remove(table);
    }
}
