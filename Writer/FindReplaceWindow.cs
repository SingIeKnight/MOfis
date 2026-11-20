using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using MOfis.Common;

namespace MOfis.Writer;

/// Bul / Degistir penceresi (modsuz; belge acikken acik kalabilir). Turkce buyuk/kucuk harf kurallarini bilir (i/İ, ı/I).
public sealed class FindReplaceWindow : Window
{
    private static readonly CompareInfo TrCompare = new CultureInfo("tr-TR").CompareInfo;

    private readonly RichTextBox _editor;
    private readonly TextBox _find = new() { Padding = new Thickness(4, 3, 4, 3) };
    private readonly TextBox _replace = new() { Padding = new Thickness(4, 3, 4, 3) };
    private readonly CheckBox _matchCase = new() { Content = "Büyük/küçük harf duyarlı", Margin = new Thickness(0, 8, 0, 0) };
    private readonly TextBlock _status = new() { Margin = new Thickness(0, 8, 0, 0), Foreground = (System.Windows.Media.Brush)Application.Current.Resources["SubtleTextBrush"] };
    private readonly FrameworkElement[] _replaceRow;

    public FindReplaceWindow(Window owner, RichTextBox editor)
    {
        _editor = editor;
        Owner = owner;
        Title = "Bul ve Değiştir";
        Width = 420;
        SizeToContent = SizeToContent.Height;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        FontFamily = (System.Windows.Media.FontFamily)Application.Current.Resources["UiFont"];
        FontSize = 13;

        var findNext = Btn("Sonrakini bul", true);
        var replaceOne = Btn("Değiştir", false);
        var replaceAll = Btn("Tümünü değiştir", false);
        findNext.IsDefault = true;
        findNext.Click += (_, _) => FindNext(true);
        replaceOne.Click += (_, _) => ReplaceOne();
        replaceAll.Click += (_, _) => ReplaceAll();

        var replaceLabel = new TextBlock { Text = "Yeni metin:", Margin = new Thickness(0, 10, 0, 4) };
        _replaceRow = [replaceLabel, _replace, replaceOne, replaceAll];

        var buttons = new WrapPanel { Margin = new Thickness(0, 14, 0, 0), HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(findNext);
        buttons.Children.Add(replaceOne);
        buttons.Children.Add(replaceAll);

        var panel = new StackPanel { Margin = new Thickness(16) };
        panel.Children.Add(new TextBlock { Text = "Aranan:", Margin = new Thickness(0, 0, 0, 4) });
        panel.Children.Add(_find);
        panel.Children.Add(replaceLabel);
        panel.Children.Add(_replace);
        panel.Children.Add(_matchCase);
        panel.Children.Add(buttons);
        panel.Children.Add(_status);
        Content = panel;

        InputBindings.Add(new KeyBinding(new Shell.RelayCommand(Close), Key.Escape, ModifierKeys.None));
    }

    private Button Btn(string text, bool accent) => new()
    {
        Content = text,
        Style = (Style)Application.Current.Resources[accent ? "AccentButton" : "PlainButton"],
        Margin = new Thickness(6, 0, 0, 6),
        Padding = new Thickness(12, 5, 12, 5),
    };

    /// Pencereyi "bul" veya "degistir" modunda gosterir; secili metin aranan kutusuna gelir.
    public void ShowMode(bool replace)
    {
        foreach (var el in _replaceRow) el.Visibility = replace ? Visibility.Visible : Visibility.Collapsed;
        Title = replace ? "Bul ve Değiştir" : "Bul";
        var sel = _editor.Selection.Text;
        if (!string.IsNullOrEmpty(sel) && !sel.Contains('\n') && sel.Length < 100) _find.Text = sel;
        _status.Text = "";
        if (!IsVisible) Show();
        Activate();
        _find.Focus();
        _find.SelectAll();
    }

    private CompareOptions Options => _matchCase.IsChecked == true ? CompareOptions.None : CompareOptions.IgnoreCase;

    /// Verilen konumdan ileriye dogru arar. Eslesme tek bir metin parcasi icinde olmalidir
    /// (ayni bicimdeki metin); farkli bicimlere bolunmus kelimeler bulunmayabilir.
    private TextRange? Search(TextPointer from, TextPointer until)
    {
        var query = _find.Text;
        if (query.Length == 0) return null;
        for (var p = from; p != null && p.CompareTo(until) < 0; p = p.GetNextContextPosition(LogicalDirection.Forward))
        {
            if (p.GetPointerContext(LogicalDirection.Forward) != TextPointerContext.Text) continue;
            var text = p.GetTextInRun(LogicalDirection.Forward);
            int idx = TrCompare.IndexOf(text, query, Options);
            if (idx < 0) continue;
            var start = p.GetPositionAtOffset(idx);
            var end = start?.GetPositionAtOffset(query.Length);
            if (start != null && end != null && end.CompareTo(until) <= 0) return new TextRange(start, end);
        }
        return null;
    }

    private bool FindNext(bool wrap)
    {
        var doc = _editor.Document;
        var found = Search(_editor.Selection.End, doc.ContentEnd);
        if (found == null && wrap) found = Search(doc.ContentStart, doc.ContentEnd);
        if (found == null)
        {
            _status.Text = $"\"{_find.Text}\" bulunamadı.";
            return false;
        }
        _status.Text = "";
        _editor.Selection.Select(found.Start, found.End);
        var rect = found.Start.GetCharacterRect(LogicalDirection.Forward);
        if (!rect.IsEmpty) _editor.BringIntoView(rect);
        return true;
    }

    private bool SelectionMatches() =>
        _find.Text.Length > 0 && TrCompare.Compare(_editor.Selection.Text, _find.Text, Options) == 0;

    private void ReplaceOne()
    {
        if (SelectionMatches()) _editor.Selection.Text = _replace.Text;
        FindNext(true);
    }

    private void ReplaceAll()
    {
        if (_find.Text.Length == 0) return;
        int count = 0;
        _editor.BeginChange(); // tek "geri al" adimi
        try
        {
            var pos = _editor.Document.ContentStart;
            while (Search(pos, _editor.Document.ContentEnd) is { } r)
            {
                r.Text = _replace.Text;
                pos = r.End; // yeni metnin sonundan devam: "a" -> "aa" sonsuz donmez
                count++;
                if (count > 100000) break;
            }
        }
        finally
        {
            _editor.EndChange();
        }
        _status.Text = count == 0 ? $"\"{_find.Text}\" bulunamadı." : $"{count} yerde değiştirildi.";
    }
}
