using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using MOfis.Common;
using MOfis.Shell;

namespace MOfis.Writer;

public partial class WriterWindow : Window, IDocumentWindow
{
    private static int _untitledCounter;
    private static readonly double[] Sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

    private PageSettings _page = new();
    private bool _dirty;
    private bool _loading;
    private bool _syncing;
    private IReadOnlyList<string> _warnings = [];
    private bool _warningsAccepted;
    private readonly string _untitledName;
    private readonly RecoveryEntry _recovery = new() { Id = Guid.NewGuid().ToString("N") };
    private readonly DispatcherTimer _autosave = new() { Interval = TimeSpan.FromMinutes(2) };
    private readonly DispatcherTimer _countTimer = new() { Interval = TimeSpan.FromMilliseconds(400) };
    private FindReplaceWindow? _findWindow;
    private Color _lastFontColor = Color.FromRgb(0xC0, 0x00, 0x00);
    private Color? _lastHighlight = Colors.Yellow;

    public string? FilePath { get; private set; }

    public WriterWindow()
    {
        InitializeComponent();
        _untitledName = "Belge" + (++_untitledCounter);

        BulletsBtn.Content = LineIcons.Bullets();
        NumbersBtn.Content = LineIcons.Numbers();
        AlignLeftBtn.Content = LineIcons.Align(TextAlignment.Left);
        AlignCenterBtn.Content = LineIcons.Align(TextAlignment.Center);
        AlignRightBtn.Content = LineIcons.Align(TextAlignment.Right);
        AlignJustifyBtn.Content = LineIcons.Align(TextAlignment.Justify);
        LineSpacingBtn.Content = LineIcons.LineSpacing();

        FontBox.ItemsSource = Fonts.SystemFontFamilies.Select(f => f.Source).Distinct().OrderBy(s => s, StringComparer.CurrentCulture).ToList();
        SizeBox.ItemsSource = Sizes.Select(FormatSize).ToList();
        FontBox.SelectionChanged += (_, _) => { if (!_syncing && FontBox.IsDropDownOpen) ApplyFont(FontBox.SelectedItem as string); };
        FontBox.DropDownClosed += (_, _) => { if (!_syncing) ApplyFont(FontBox.SelectedItem as string ?? FontBox.Text); };
        FontBox.PreviewKeyDown += (_, e) => { if (e.Key == Key.Enter) { ApplyFont(FontBox.Text); e.Handled = true; } };
        SizeBox.DropDownClosed += (_, _) => { if (!_syncing && SizeBox.SelectedItem is string s) ApplySize(s); };
        SizeBox.PreviewKeyDown += (_, e) => { if (e.Key == Key.Enter) { ApplySize(SizeBox.Text); e.Handled = true; } };

        BuildStyleGallery();
        SetDocument(DocHelpers.CreateBlank(), new PageSettings());

        Editor.TextChanged += Editor_TextChanged;
        Editor.SelectionChanged += Editor_SelectionChanged;
        Editor.PreviewKeyDown += Editor_PreviewKeyDown;
        Editor.ContextMenuOpening += Editor_ContextMenuOpening;
        Editor.ContextMenu = new ContextMenu();
        Editor.AddHandler(Hyperlink.RequestNavigateEvent, new System.Windows.Navigation.RequestNavigateEventHandler(Link_RequestNavigate));
        CommandManager.AddPreviewExecutedHandler(Editor, Editor_PreviewExecuted);
        Scroller.PreviewMouseWheel += Scroller_PreviewMouseWheel;
        HookRulers();
        ZoomSlider.ValueChanged += (_, e) => SetZoom(e.NewValue / 100.0);
        SetZoom(1);

        _autosave.Tick += (_, _) => { if (_dirty) AutoRecovery.Save(Recovery(), Editor.Document, _page); };
        _autosave.Start();
        _countTimer.Tick += (_, _) => { _countTimer.Stop(); UpdateWordCount(); };

        AddShortcuts();
        Loaded += (_, _) => { Editor.Focus(); UpdateWordCount(); SyncToolbar(); };
        Closing += Window_Closing;
        Closed += (_, _) =>
        {
            _autosave.Stop();
            AutoRecovery.Delete(_recovery.Id);
        };
        UpdateTitle();
    }

    // ================================================================ belge yukleme / kaydetme

    private void SetDocument(FlowDocument doc, PageSettings page)
    {
        _loading = true;
        Editor.Document = doc;
        // RichTextBox ilk cizimde PagePadding'i kendi degerine (5,0,5,0) ceker; her degistiginde
        // kenar boslugu telafisini yeniden hesapla ki yazi alani Word ile ayni genislikte kalsin.
        DependencyPropertyDescriptor.FromProperty(FlowDocument.PagePaddingProperty, typeof(FlowDocument))
            .AddValueChanged(doc, (_, _) => { if (Editor.Document == doc) ApplyPage(); });
        DocHelpers.ShowPageBreaks(doc);
        _page = page;
        ApplyPage();
        _loading = false;
        _dirty = false;
        UpdateTitle();
        UpdateWordCount();
    }

    private void ApplyPage()
    {
        PageHost.Width = _page.Width;
        Editor.MinHeight = _page.Height;
        // RichTextBox belgeye kendi ic boslugunu (PagePadding, varsayilan 5 px) ekler; yazi alani
        // Word ile birebir ayni genislikte olsun diye bu bosluk kenar boslugundan dusulur.
        var inner = Editor.Document.PagePadding;
        Editor.Padding = new Thickness(
            Math.Max(0, _page.MarginLeft - inner.Left), Math.Max(0, _page.MarginTop - inner.Top),
            Math.Max(0, _page.MarginRight - inner.Right), Math.Max(0, _page.MarginBottom - inner.Bottom));
        Dispatcher.BeginInvoke(UpdateRulers, DispatcherPriority.Loaded);
    }

    /// Gelistirme: yerlesim olculeri (--shot ile birlikte .txt olarak yazilir).
    internal string DescribeLayout()
    {
        var d = Editor.Document;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Editor.ActualWidth={Editor.ActualWidth} Padding={Editor.Padding} PageHost.Width={PageHost.Width}");
        sb.AppendLine($"Doc.PagePadding={d.PagePadding} ColumnWidth={d.ColumnWidth} PageWidth={d.PageWidth} MaxPageWidth={d.MaxPageWidth} ColumnGap={d.ColumnGap} RuleWidth={d.ColumnRuleWidth}");
        sb.AppendLine($"ContentWidth(hesap)={_page.ContentWidth}");
        foreach (var t in d.Blocks.OfType<Table>())
        {
            sb.AppendLine($"Table margin={t.Margin} padding={t.Padding} border={t.BorderThickness} cols=" + string.Join(",", t.Columns.Select(c => c.Width.ToString())));
            var r = t.ContentStart.GetCharacterRect(LogicalDirection.Forward);
            var e = t.ContentEnd.GetCharacterRect(LogicalDirection.Backward);
            sb.AppendLine($"  start={r} end={e}");
        }
        var p = d.Blocks.FirstBlock;
        sb.AppendLine($"FirstBlock start rect={p.ContentStart.GetCharacterRect(LogicalDirection.Forward)}");
        return sb.ToString();
    }

    public bool LoadFile(string path)
    {
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            var loaded = DocumentIO.Load(path);
            SetDocument(loaded.Document, loaded.Page);
            FilePath = path;
            _warnings = loaded.Warnings;
            _warningsAccepted = false;
            RecentFiles.Add(path);
            UpdateTitle();
            if (_warnings.Count > 0)
                ShowInfo("Bu belgedeki bazı öğeler henüz desteklenmiyor ve gösterilmiyor: " + string.Join(", ", _warnings) +
                         ". Belgeyi aynı dosyanın üzerine kaydederseniz bu öğeler kaybolur; \"Farklı kaydet\" önerilir.");
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Dosya açılamadı:\n{path}\n\n{ex.Message}", AppInfo.WriterName, MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    public bool LoadRecovery(RecoveryEntry entry)
    {
        try
        {
            var loaded = DocumentIO.Load(entry.DocPath);
            SetDocument(loaded.Document, loaded.Page);
            _recovery.Id = entry.Id; // ayni kurtarma kaydi uzerinden devam
            FilePath = entry.OriginalPath != null && File.Exists(entry.OriginalPath) ? entry.OriginalPath : null;
            _dirty = true;
            UpdateTitle();
            ShowInfo("Bu belge otomatik kurtarma kopyasından açıldı. Saklamak için kaydedin.");
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Kurtarma kopyası açılamadı:\n\n" + ex.Message, AppInfo.WriterName, MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private RecoveryEntry Recovery()
    {
        _recovery.Title = DisplayName;
        _recovery.OriginalPath = FilePath;
        return _recovery;
    }

    private string DisplayName => FilePath != null ? Path.GetFileName(FilePath) : _untitledName;

    private bool Save()
    {
        if (FilePath == null) return SaveAs();
        if (_warnings.Count > 0 && !_warningsAccepted)
        {
            var answer = MessageBox.Show(
                "Bu belgede MOfis'in henüz desteklemediği öğeler vardı:\n\n• " + string.Join("\n• ", _warnings) +
                "\n\nAynı dosyanın üzerine kaydederseniz bu öğeler dosyadan çıkarılır.\n\n" +
                "Evet: yine de üzerine kaydet\nHayır: farklı bir dosyaya kaydet\nİptal: kaydetme",
                AppInfo.WriterName, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (answer == MessageBoxResult.Cancel) return false;
            if (answer == MessageBoxResult.No) return SaveAs();
            _warningsAccepted = true;
        }
        return SaveTo(FilePath);
    }

    private bool SaveAs()
    {
        var dlg = new SaveFileDialog
        {
            Filter = DocumentIO.SaveFilter,
            FileName = Path.GetFileNameWithoutExtension(DisplayName),
            Title = "Farklı kaydet",
            AddExtension = true,
        };
        if (FilePath != null) dlg.InitialDirectory = Path.GetDirectoryName(FilePath);
        if (dlg.ShowDialog(this) != true) return false;
        if (SaveTo(dlg.FileName))
        {
            _warnings = [];
            return true;
        }
        return false;
    }

    private bool SaveTo(string path)
    {
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            DocumentIO.Save(Editor.Document, _page, path);
            FilePath = path;
            _dirty = false;
            RecentFiles.Add(path);
            AutoRecovery.Delete(_recovery.Id);
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Dosya kaydedilemedi. Başka bir programda açık olabilir.\n\n{ex.Message}", AppInfo.WriterName,
                MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (!_dirty) return;
        var answer = MessageBox.Show($"\"{DisplayName}\" belgesindeki değişiklikler kaydedilsin mi?", AppInfo.WriterName,
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (answer == MessageBoxResult.Cancel || (answer == MessageBoxResult.Yes && !Save())) e.Cancel = true;
    }

    private void UpdateTitle() => Title = $"{DisplayName}{(_dirty ? " •" : "")} - {AppInfo.WriterName}";

    private void MarkDirty()
    {
        if (_loading || _dirty) return;
        _dirty = true;
        UpdateTitle();
    }

    // ================================================================ bilgi cubugu, durum cubugu

    private void ShowInfo(string text)
    {
        InfoText.Text = text;
        InfoBar.Visibility = Visibility.Visible;
    }

    private void CloseInfo_Click(object sender, RoutedEventArgs e) => InfoBar.Visibility = Visibility.Collapsed;

    private void UpdateWordCount()
    {
        var text = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
        int words = Regex.Matches(text, @"[\p{L}\p{N}][\p{L}\p{N}'’\-]*").Count;
        int chars = text.Count(c => c != '\r' && c != '\n');
        WordCountText.Text = $"{words:N0} sözcük   ·   {chars:N0} karakter";
    }

    // ================================================================ editor olaylari

    private void Editor_TextChanged(object sender, TextChangedEventArgs e)
    {
        MarkDirty();
        _countTimer.Stop();
        _countTimer.Start();
        BringCaretIntoView();
    }

    private void Editor_SelectionChanged(object sender, RoutedEventArgs e)
    {
        SyncToolbar();
        if (Mouse.LeftButton != MouseButtonState.Pressed) BringCaretIntoView();
    }

    /// Ic kaydirma kapali oldugundan imlecin gorunur kalmasini dis kaydirma cubugu saglar.
    private void BringCaretIntoView()
    {
        var rect = Editor.CaretPosition.GetCharacterRect(LogicalDirection.Forward);
        if (rect.IsEmpty) return;
        rect.Inflate(0, 24);
        Editor.BringIntoView(rect);
    }

    private void Editor_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        var mods = Keyboard.Modifiers;
        if (e.Key == Key.Enter && mods == ModifierKeys.Control)
        {
            InsertPageBreak();
            e.Handled = true;
            return;
        }
        // Basligin sonunda Enter: yeni satir Word'deki gibi "Normal" stille baslar.
        if (e.Key == Key.Enter && mods == ModifierKeys.None)
        {
            var p = Editor.CaretPosition.Paragraph;
            if (p?.Tag is string id && StyleCatalog.ById(id) != null &&
                new TextRange(Editor.CaretPosition, p.ContentEnd).Text.Length == 0)
            {
                Dispatcher.BeginInvoke(() =>
                {
                    var np = Editor.CaretPosition.Paragraph;
                    if (np != null && np != p) DocHelpers.ApplyStyle(np, StyleCatalog.Normal);
                    SyncToolbar();
                }, DispatcherPriority.Input);
            }
        }
    }

    /// Panodaki resim yapistirilirken PNG olarak eklenir (WPF bunu kendiliginden yapmaz).
    private void Editor_PreviewExecuted(object sender, ExecutedRoutedEventArgs e)
    {
        if (e.Command != ApplicationCommands.Paste) return;
        if (Clipboard.ContainsImage() && !Clipboard.ContainsText() && !Clipboard.ContainsData(DataFormats.Rtf))
        {
            var bmp = Clipboard.GetImage();
            if (bmp == null) return;
            var enc = new PngBitmapEncoder();
            enc.Frames.Add(BitmapFrame.Create(bmp));
            using var ms = new MemoryStream();
            enc.Save(ms);
            InsertImageBytes(ms.ToArray(), ".png");
            e.Handled = true;
        }
    }

    private void Link_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        e.Handled = true;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) OpenLink(e.Uri);
    }

    /// Sadece web ve e-posta baglantilari acilir (belgedeki baglanti ile program calistirilamaz).
    private static void OpenLink(Uri? uri)
    {
        if (uri is not { IsAbsoluteUri: true } || uri.Scheme is not ("http" or "https" or "mailto")) return;
        try { Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true }); } catch { }
    }

    private void Scroller_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) return;
        SetZoom(ZoomTransform.ScaleX + (e.Delta > 0 ? 0.1 : -0.1));
        e.Handled = true;
    }

    // ================================================================ arac cubugunu secime gore guncelle

    private void SyncToolbar()
    {
        _syncing = true;
        try
        {
            var sel = Editor.Selection;
            BoldBtn.IsChecked = sel.GetPropertyValue(TextElement.FontWeightProperty) is FontWeight w && w.ToOpenTypeWeight() >= 600;
            ItalicBtn.IsChecked = sel.GetPropertyValue(TextElement.FontStyleProperty) is FontStyle fs && fs != FontStyles.Normal;
            var (u, s) = CurrentDecorations();
            UnderlineBtn.IsChecked = u;
            StrikeBtn.IsChecked = s;
            var baseline = sel.GetPropertyValue(Inline.BaselineAlignmentProperty);
            SupBtn.IsChecked = baseline is BaselineAlignment.Superscript;
            SubBtn.IsChecked = baseline is BaselineAlignment.Subscript;

            FontBox.Text = sel.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily ff ? ff.Source.Split(',')[0].Trim() : "";
            SizeBox.Text = sel.GetPropertyValue(TextElement.FontSizeProperty) is double px ? FormatSize(Math.Round(Units.PxToPt(px) * 2) / 2) : "";

            var align = sel.GetPropertyValue(Paragraph.TextAlignmentProperty);
            AlignLeftBtn.IsChecked = align is TextAlignment.Left;
            AlignCenterBtn.IsChecked = align is TextAlignment.Center;
            AlignRightBtn.IsChecked = align is TextAlignment.Right;
            AlignJustifyBtn.IsChecked = align is TextAlignment.Justify;
        }
        finally
        {
            _syncing = false;
        }
        UpdateRulers();
    }

    // ================================================================ cetveller

    private List<(Paragraph p, double left, double indent)> _dragParas = [];

    private void HookRulers()
    {
        Scroller.ScrollChanged += (_, _) => UpdateRulers();
        SizeChanged += (_, _) => UpdateRulers();
        HRuler.DragStarted += Ruler_DragStarted;
        VRuler.DragStarted += Ruler_DragStarted;
        HRuler.Dragging += (target, v) => Ruler_Dragging(false, target, v);
        VRuler.Dragging += (target, v) => Ruler_Dragging(true, target, v);
        HRuler.DragCompleted += Ruler_DragCompleted;
        VRuler.DragCompleted += Ruler_DragCompleted;
    }

    private void RulerCheck_Click(object sender, RoutedEventArgs e)
    {
        var v = RulerCheck.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        HRuler.Visibility = VRuler.Visibility = RulerCorner.Visibility = v;
        Dispatcher.BeginInvoke(UpdateRulers, DispatcherPriority.Loaded);
    }

    /// Cetvelleri sayfanin ekrandaki yerine ve imlecin bulundugu paragrafin girintilerine gore gunceller.
    private void UpdateRulers()
    {
        if (!IsLoaded || HRuler.Visibility != Visibility.Visible) return;
        double scale = ZoomTransform.ScaleX;
        var h = PageHost.TranslatePoint(new Point(0, 0), HRuler);
        HRuler.SetPage(h.X, scale, _page.Width, _page.MarginLeft, _page.MarginRight);
        var v = PageHost.TranslatePoint(new Point(0, 0), VRuler);
        VRuler.SetPage(v.Y, scale, _page.Height, _page.MarginTop, _page.MarginBottom);
        if (_indentDrag != null) return; // surukleme onizlemesi suruyor; isaretcileri paragraftan okuma

        var p = Editor.CaretPosition.Paragraph;
        if (p == null || TableOps.Locate(Editor.CaretPosition) != null)
        {
            HRuler.SetIndents(false, 0, 0, 0, 0); // tablo hucresinde girinti isaretcileri gosterilmez
            return;
        }
        double left = Nz(p.Margin.Left), right = Nz(p.Margin.Right), indent = Nz(p.TextIndent);
        HRuler.SetIndents(true, ListOffset(p, left, indent), left, indent, right);
    }

    /// Liste maddesindeki paragrafin, liste girintisi kadar ice kaymis baslangici (yazi alanina gore).
    private double ListOffset(Paragraph p, double left, double indent)
    {
        if (FindAncestor<ListItem>(p) == null) return 0;
        var rect = p.ContentStart.GetCharacterRect(LogicalDirection.Forward);
        if (rect.IsEmpty) return HRuler.IndentBase;
        double textLeft = Editor.Padding.Left + Editor.Document.PagePadding.Left;
        return Math.Max(0, rect.X - textLeft - left - indent);
    }

    private static double Nz(double v) => double.IsNaN(v) ? 0 : v;

    /// Math.Clamp, alt sinir ust siniri gecince (kayan nokta yuvarlamasi) hata verir; bu vermez, alt sinir kazanir.
    private static double Clamp(double v, double lo, double hi) => Math.Max(lo, Math.Min(v, hi));

    private string? _indentDrag;       // surukleme suren girinti isaretcisi (kenar bosluklari degil)
    private double _indentDragValue;   // son fare konumu, sayfa koordinati

    /// Surukleme sirasinda belgeye dokunulmaz: editorde degisiklik blogu acik kalirsa WPF cizim sirasinda
    /// "Dispatcher islemi askiya alindi" hatasi verir. Sadece cetvel ve kilavuz cizgisi hareket eder.
    private void Ruler_DragStarted(string target)
    {
        _dragParas = SelectedParagraphs().Select(p => (p, Nz(p.Margin.Left), Nz(p.TextIndent))).ToList();
        if (target is Ruler.MarginStart or Ruler.MarginEnd) return;
        _indentDrag = target;
        _indentDragValue = double.NaN;
    }

    private void Ruler_Dragging(bool vertical, string target, double v)
    {
        double minText = Units.CmToPx(2);
        if (vertical)
        {
            if (target == Ruler.MarginStart) _page.MarginTop = Clamp(v, 0, _page.Height - _page.MarginBottom - minText);
            else if (target == Ruler.MarginEnd) _page.MarginBottom = Clamp(_page.Height - v, 0, _page.Height - _page.MarginTop - minText);
            PageChanged();
            return;
        }

        switch (target)
        {
            case Ruler.MarginStart:
                _page.MarginLeft = Clamp(v, 0, _page.Width - _page.MarginRight - minText);
                PageChanged();
                return;
            case Ruler.MarginEnd:
                _page.MarginRight = Clamp(_page.Width - v, 0, _page.Width - _page.MarginLeft - minText);
                PageChanged();
                return;
        }

        // Girinti: onizleme (cetvel + kilavuz cizgisi); uygulama birakinca.
        _indentDragValue = v;
        var caret = Editor.CaretPosition.Paragraph;
        var item = _dragParas.FirstOrDefault(x => x.p == caret);
        if (item.p == null && _dragParas.Count > 0) item = _dragParas[0];
        if (item.p == null) return;
        var (left, indent, right) = ComputeIndents(target, v, item.p, item.left, item.indent);
        HRuler.SetIndents(true, HRuler.IndentBase, left, indent, right);
        ShowGuide(target switch
        {
            Ruler.FirstLine => HRuler.FirstX,
            Ruler.RightIndent => HRuler.RightX,
            _ => HRuler.LeftX,
        });
    }

    /// Surukleme konumuna (sayfa koordinati) gore paragrafin yeni sol girinti / ilk satir / sag girintisi.
    private (double left, double indent, double right) ComputeIndents(string target, double v, Paragraph p, double left0, double indent0)
    {
        double rel = v - _page.MarginLeft - HRuler.IndentBase;
        double right = Nz(p.Margin.Right);
        double maxLeft = Math.Max(0, _page.ContentWidth - HRuler.IndentBase - right - 40);
        switch (target)
        {
            case Ruler.FirstLine:
                // ilk satir, sol kenar boslugunun disina tasamaz
                return (left0, Clamp(rel, 0, maxLeft) - left0, right);
            case Ruler.Hanging:
            {
                // asili girinti: sol girinti degisir, ilk satir yerinde kalir
                double newLeft = Clamp(rel, 0, maxLeft);
                return (newLeft, Math.Max(left0 + indent0, 0) - newLeft, right);
            }
            case Ruler.LeftBox:
                // sol girinti kutusu: ilk satirla birlikte tasinir
                return (Clamp(rel, Math.Max(0, -indent0), maxLeft), indent0, right);
            case Ruler.RightIndent:
                return (left0, indent0, Clamp(_page.Width - _page.MarginRight - v, 0, Math.Max(0, _page.ContentWidth - left0 - 40)));
            default:
                return (left0, indent0, right);
        }
    }

    private void ShowGuide(double x)
    {
        GuideLine.X1 = GuideLine.X2 = Math.Round(x) + 0.5;
        GuideLine.Y1 = 0;
        GuideLine.Y2 = GuideLayer.ActualHeight;
        GuideLine.Visibility = Visibility.Visible;
    }

    private void HideGuide() => GuideLine.Visibility = Visibility.Collapsed;


    /// Gelistirme: her cetvel isaretcisini farkli paragraflarda surukleme simulasyonu; hatalari dondurur.
    internal string SimulateRulerDrags()
    {
        var log = new System.Text.StringBuilder();
        var paragraphs = DocHelpers.AllParagraphs(Editor.Document.Blocks).Take(8).ToList();
        foreach (var p in paragraphs)
        {
            Editor.CaretPosition = p.ContentStart;
            UpdateRulers();
            foreach (var (vertical, target) in new[] { (false, Ruler.FirstLine), (false, Ruler.Hanging), (false, Ruler.LeftBox), (false, Ruler.RightIndent), (false, Ruler.MarginStart), (false, Ruler.MarginEnd), (true, Ruler.MarginStart), (true, Ruler.MarginEnd) })
            {
                try
                {
                    Ruler_DragStarted(target);
                    foreach (var v in new[] { 60.0, 150, 300, 10, 700, 500 })
                    {
                        Ruler_Dragging(vertical, target, v);
                        if (Environment.GetEnvironmentVariable("MOFIS_RULER_LAYOUT") == "1") UpdateLayout();
                    }
                    Ruler_DragCompleted(target);
                    log.AppendLine($"OK   {(vertical ? "V" : "H")} {target,-8} '{new TextRange(p.ContentStart, p.ContentEnd).Text}'");
                }
                catch (Exception ex)
                {
                    log.AppendLine($"HATA {(vertical ? "V" : "H")} {target,-8} '{new TextRange(p.ContentStart, p.ContentEnd).Text}': {ex}");
                    try { Editor.EndChange(); } catch { }
                }
            }
        }
        // 2. asama: cetvelin kendi fare kodu (uzerine gelme, basma, surukleme, birakma).
        foreach (var p in paragraphs.Take(5))
        {
            Editor.CaretPosition = p.ContentStart;
            UpdateRulers();
            UpdateLayout();
            var points = new (Ruler ruler, string name, Func<Point> at)[]
            {
                (HRuler, Ruler.FirstLine, () => new Point(HRuler.FirstX, 5)),
                (HRuler, Ruler.Hanging, () => new Point(HRuler.LeftX, 15)),
                (HRuler, Ruler.LeftBox, () => new Point(HRuler.LeftX, 20.5)),
                (HRuler, Ruler.RightIndent, () => new Point(HRuler.RightX, 15)),
                (HRuler, Ruler.MarginStart, () => new Point(HRuler.TextStart, 11)),
                (HRuler, Ruler.MarginEnd, () => new Point(HRuler.TextEnd, 11)),
                (VRuler, Ruler.MarginStart, () => new Point(11, VRuler.TextStart)),
                (VRuler, Ruler.MarginEnd, () => new Point(11, VRuler.TextEnd)),
            };
            foreach (var (ruler, name, at) in points)
            {
                string step = "move";
                try
                {
                    var start = at();
                    ruler.SimulateMouse("move", start);
                    step = "down";
                    ruler.SimulateMouse("down", start);
                    foreach (var d in new[] { 10.0, 40, -30, 80 })
                    {
                        step = "drag" + d;
                        var to = ruler.Vertical ? new Point(start.X, start.Y + d) : new Point(start.X + d, start.Y);
                        ruler.SimulateMouse("move", to);
                    }
                    step = "up";
                    ruler.SimulateMouse("up", start);
                    log.AppendLine($"FARE OK   {(ruler.Vertical ? "V" : "H")} {name,-8}");
                }
                catch (Exception ex)
                {
                    log.AppendLine($"FARE HATA {(ruler.Vertical ? "V" : "H")} {name,-8} adim={step}: {ex}");
                }
                finally
                {
                    ruler.TestPosition = null;
                }
            }
        }
        _dirty = false; // simulasyon kapanista "kaydedilsin mi" sormasin
        return log.ToString();
    }

    private void Ruler_DragCompleted(string target)
    {
        HideGuide();
        var drag = _indentDrag;
        _indentDrag = null;
        if (drag != null && !double.IsNaN(_indentDragValue))
        {
            // Girinti, fare birakilinca tek seferde ve tek "geri al" adimi olarak uygulanir.
            Editor.BeginChange();
            try
            {
                foreach (var (p, left0, indent0) in _dragParas)
                {
                    var (left, indent, right) = ComputeIndents(drag, _indentDragValue, p, left0, indent0);
                    var m = p.Margin;
                    p.Margin = new Thickness(left, m.Top, right, m.Bottom);
                    p.TextIndent = indent;
                }
            }
            finally
            {
                Editor.EndChange();
            }
            MarkDirty();
        }
        _dragParas = [];
        // Odak degisikligi fare birakma isleminin icinde yapilmasin ("Dispatcher askida" hatasi).
        Dispatcher.BeginInvoke(() => { Editor.Focus(); UpdateRulers(); }, DispatcherPriority.Input);
    }

    private void PageChanged()
    {
        ApplyPage();
        _dirty = true;
        UpdateTitle();
    }

    private static string FormatSize(double pt) => pt.ToString("0.#", CultureInfo.CurrentCulture);

    // ================================================================ yazi bicimi

    private void Bold_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.ToggleBold);
    private void Italic_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.ToggleItalic);

    private void Exec(RoutedUICommand cmd)
    {
        cmd.Execute(null, Editor);
        Editor.Focus();
        SyncToolbar();
    }

    private (bool underline, bool strike) CurrentDecorations()
    {
        if (Editor.Selection.GetPropertyValue(Inline.TextDecorationsProperty) is not TextDecorationCollection decos) return (false, false);
        return (decos.Any(d => d.Location == TextDecorationLocation.Underline), decos.Any(d => d.Location == TextDecorationLocation.Strikethrough));
    }

    /// Alti cizili ve ustu cizili birbirini silmeden ayri ayri acilip kapanir.
    private void SetDecorations(bool underline, bool strike)
    {
        var decos = new TextDecorationCollection();
        if (underline) decos.Add(TextDecorations.Underline);
        if (strike) decos.Add(TextDecorations.Strikethrough);
        Editor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, decos);
        Editor.Focus();
        SyncToolbar();
    }

    private void Underline_Click(object sender, RoutedEventArgs e)
    {
        var (u, s) = CurrentDecorations();
        SetDecorations(!u, s);
    }

    private void Strike_Click(object sender, RoutedEventArgs e)
    {
        var (u, s) = CurrentDecorations();
        SetDecorations(u, !s);
    }

    private void ToggleBaseline(BaselineAlignment target)
    {
        var current = Editor.Selection.GetPropertyValue(Inline.BaselineAlignmentProperty);
        var value = current is BaselineAlignment b && b == target ? BaselineAlignment.Baseline : target;
        Editor.Selection.ApplyPropertyValue(Inline.BaselineAlignmentProperty, value);
        Editor.Focus();
        SyncToolbar();
    }

    private void Subscript_Click(object sender, RoutedEventArgs e) => ToggleBaseline(BaselineAlignment.Subscript);
    private void Superscript_Click(object sender, RoutedEventArgs e) => ToggleBaseline(BaselineAlignment.Superscript);

    private void ApplyFont(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily(name.Trim()));
        Editor.Focus();
    }

    private void ApplySize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        var normalized = text.Trim().Replace(',', '.');
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt) || pt < 1 || pt > 1638)
        {
            SyncToolbar();
            return;
        }
        Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, Units.PtToPx(pt));
        Editor.Focus();
        SyncToolbar();
    }

    private void StepSize(int direction)
    {
        double current = Editor.Selection.GetPropertyValue(TextElement.FontSizeProperty) is double px ? Units.PxToPt(px) : StyleCatalog.DefaultSizePt;
        double next = direction > 0
            ? Sizes.FirstOrDefault(s => s > current + 0.01, Math.Min(current + 10, 1638))
            : Sizes.LastOrDefault(s => s < current - 0.01, Math.Max(current - 1, 1));
        ApplySize(next.ToString(CultureInfo.InvariantCulture));
    }

    private void Grow_Click(object sender, RoutedEventArgs e) => StepSize(1);
    private void Shrink_Click(object sender, RoutedEventArgs e) => StepSize(-1);

    private void ClearFormat_Click(object sender, RoutedEventArgs e)
    {
        Editor.Selection.ClearAllProperties();
        Editor.Focus();
        SyncToolbar();
    }

    private void FontColor_Click(object sender, RoutedEventArgs e) =>
        Pickers.ShowColors(FontColorBtn, "Otomatik (siyah)", false, c =>
        {
            var color = c ?? Colors.Black;
            if (c != null)
            {
                _lastFontColor = color;
                FontColorBar.Background = new SolidColorBrush(color);
            }
            Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color));
            Editor.Focus();
        });

    private void Highlight_Click(object sender, RoutedEventArgs e) =>
        Pickers.ShowColors(HighlightBtn, "Renk yok", true, c =>
        {
            _lastHighlight = c;
            HighlightBar.Background = c is Color col ? new SolidColorBrush(col) : Brushes.Transparent;
            Editor.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, c is Color hc ? new SolidColorBrush(hc) : null);
            Editor.Focus();
        });

    // ================================================================ paragraf

    private void Bullets_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.ToggleBullets);
    private void Numbers_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.ToggleNumbering);
    private void Indent_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.IncreaseIndentation);
    private void Outdent_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.DecreaseIndentation);
    private void AlignLeft_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.AlignLeft);
    private void AlignCenter_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.AlignCenter);
    private void AlignRight_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.AlignRight);
    private void AlignJustify_Click(object sender, RoutedEventArgs e) => Exec(EditingCommands.AlignJustify);

    private IEnumerable<Paragraph> SelectedParagraphs()
    {
        var sel = Editor.Selection;
        var list = DocHelpers.AllParagraphs(Editor.Document.Blocks)
            .Where(p => p.ContentEnd.CompareTo(sel.Start) >= 0 && p.ContentStart.CompareTo(sel.End) <= 0)
            .ToList();
        if (list.Count == 0 && Editor.CaretPosition.Paragraph is { } cp) list.Add(cp);
        return list;
    }

    private void SetLineSpacing(double factor)
    {
        Editor.BeginChange();
        foreach (var p in SelectedParagraphs())
        {
            if (Math.Abs(factor - 1) < 0.01) p.ClearValue(Paragraph.LineHeightProperty);
            else p.LineHeight = factor * p.FontFamily.LineSpacing * p.FontSize;
        }
        Editor.EndChange();
        MarkDirty();
        Editor.Focus();
    }

    private void LineSpacing_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = LineSpacingBtn, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        foreach (var f in new[] { 1.0, 1.15, 1.5, 2.0, 2.5, 3.0 })
        {
            var item = new MenuItem { Header = f.ToString("0.0#", CultureInfo.CurrentCulture) };
            item.Click += (_, _) => SetLineSpacing(f);
            menu.Items.Add(item);
        }
        menu.Items.Add(new Separator());
        var addBefore = new MenuItem { Header = "Paragraftan önce boşluk ekle (12 nk)" };
        addBefore.Click += (_, _) => SetParagraphSpace(before: Units.PtToPx(12), after: null);
        var noAfter = new MenuItem { Header = "Paragraftan sonraki boşluğu kaldır" };
        noAfter.Click += (_, _) => SetParagraphSpace(before: null, after: 0);
        menu.Items.Add(addBefore);
        menu.Items.Add(noAfter);
        menu.IsOpen = true;
    }

    private void SetParagraphSpace(double? before, double? after)
    {
        Editor.BeginChange();
        foreach (var p in SelectedParagraphs())
        {
            var m = p.Margin;
            p.Margin = new Thickness(m.Left, before ?? m.Top, m.Right, after ?? m.Bottom);
        }
        Editor.EndChange();
        MarkDirty();
        Editor.Focus();
    }

    private void BuildStyleGallery()
    {
        foreach (var s in StyleCatalog.All)
        {
            var sample = new TextBlock
            {
                Text = "AaBbÇç",
                FontFamily = new FontFamily(StyleCatalog.DefaultFont),
                FontSize = Math.Min(Units.PtToPx(s.SizePt), 21),
                FontWeight = s.Bold ? FontWeights.Bold : FontWeights.Normal,
                Foreground = new SolidColorBrush(s.Color),
                HorizontalAlignment = HorizontalAlignment.Center,
                Height = 26,
            };
            var name = new TextBlock { Text = s.DisplayName, FontSize = 11, HorizontalAlignment = HorizontalAlignment.Center, Foreground = (Brush)FindResource("SubtleTextBrush") };
            var stack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            stack.Children.Add(sample);
            stack.Children.Add(name);
            var card = new Border
            {
                Width = 86,
                Height = 58,
                Background = Brushes.White,
                BorderBrush = (Brush)FindResource("LineBrush"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Child = stack,
            };
            var btn = new Button { Content = card, Style = (Style)FindResource("TileButton"), Padding = new Thickness(2), Focusable = false, ToolTip = s.DisplayName + StyleShortcut(s) };
            var style = s;
            btn.Click += (_, _) => ApplyParagraphStyle(style);
            StyleGallery.Children.Add(btn);
        }
    }

    private static string StyleShortcut(ParaStyleDef s) => s.Id switch
    {
        "Heading1" => " (Ctrl+Alt+1)",
        "Heading2" => " (Ctrl+Alt+2)",
        "Heading3" => " (Ctrl+Alt+3)",
        "Normal" => " (Ctrl+Shift+N)",
        _ => "",
    };

    private void ApplyParagraphStyle(ParaStyleDef s)
    {
        Editor.BeginChange();
        foreach (var p in SelectedParagraphs()) DocHelpers.ApplyStyle(p, s);
        Editor.EndChange();
        MarkDirty();
        Editor.Focus();
        SyncToolbar();
    }

    // ================================================================ ekle

    private void InsertTable_Click(object sender, RoutedEventArgs e) =>
        Pickers.ShowTableGrid(TableBtn, InsertTable);

    private void InsertTable(int rows, int cols)
    {
        var line = new SolidColorBrush(Color.FromRgb(0x80, 0x80, 0x80));
        var table = new Table { CellSpacing = 0, BorderBrush = line, BorderThickness = new Thickness(1, 1, 0, 0) };
        for (int c = 0; c < cols; c++) table.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var group = new TableRowGroup();
        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
                row.Cells.Add(new TableCell(new Paragraph { Margin = new Thickness(0) }) { BorderBrush = line, BorderThickness = new Thickness(0, 0, 1, 1), Padding = new Thickness(7, 3, 7, 3) });
            group.Rows.Add(row);
        }
        table.RowGroups.Add(group);

        Editor.BeginChange();
        var current = Editor.CaretPosition.Paragraph;
        if (current?.SiblingBlocks is { } siblings)
        {
            bool empty = new TextRange(current.ContentStart, current.ContentEnd).Text.Length == 0;
            if (empty) siblings.InsertBefore(current, table);
            else
            {
                siblings.InsertAfter(current, table);
                if (table.NextBlock == null) siblings.InsertAfter(table, new Paragraph());
            }
        }
        else
        {
            Editor.Document.Blocks.Add(table);
            Editor.Document.Blocks.Add(new Paragraph());
        }
        Editor.EndChange();
        Editor.CaretPosition = group.Rows[0].Cells[0].ContentStart;
        Editor.Focus();
    }

    private void InsertImage_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Resim ekle",
            Filter = "Resimler|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff|Tüm dosyalar|*.*",
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            InsertImageBytes(File.ReadAllBytes(dlg.FileName), Path.GetExtension(dlg.FileName).ToLowerInvariant());
        }
        catch (Exception ex)
        {
            MessageBox.Show("Resim eklenemedi:\n\n" + ex.Message, AppInfo.WriterName, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void InsertImageBytes(byte[] bytes, string ext)
    {
        var el = DocHelpers.CreateImageElement(new ImageData(bytes, ext), null, null, _page.ContentWidth);
        if (el is not Image)
        {
            MessageBox.Show("Bu resim biçimi desteklenmiyor.", AppInfo.WriterName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        Editor.BeginChange();
        if (!Editor.Selection.IsEmpty) Editor.Selection.Text = "";
        var pos = Editor.CaretPosition.GetInsertionPosition(LogicalDirection.Forward);
        var container = new InlineUIContainer(el, pos);
        Editor.EndChange();
        Editor.CaretPosition = container.ElementEnd;
        Editor.Focus();
    }

    private void InsertLink_Click(object sender, RoutedEventArgs e) => InsertLink();

    private void InsertLink()
    {
        var input = InputDialog.Ask(this, "Bağlantı ekle", "Web adresi (örn. https://ornek.com):", "https://");
        if (string.IsNullOrWhiteSpace(input) || input.Trim() == "https://") return;
        var text = input.Trim();
        if (!text.Contains("://") && !text.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            text = (text.Contains('@') ? "mailto:" : "https://") + text;
        if (!Uri.TryCreate(text, UriKind.Absolute, out var uri))
        {
            MessageBox.Show("Geçerli bir adres değil.", AppInfo.WriterName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Editor.BeginChange();
        if (Editor.Selection.IsEmpty)
        {
            var pos = Editor.CaretPosition.GetInsertionPosition(LogicalDirection.Forward);
            var link = new Hyperlink(new Run(uri.OriginalString), pos) { NavigateUri = uri };
            Editor.CaretPosition = link.ElementEnd;
        }
        else
        {
            _ = new Hyperlink(Editor.Selection.Start, Editor.Selection.End) { NavigateUri = uri };
        }
        Editor.EndChange();
        Editor.Focus();
    }

    private void PageBreak_Click(object sender, RoutedEventArgs e) => InsertPageBreak();

    private void InsertPageBreak()
    {
        Editor.BeginChange();
        if (!Editor.Selection.IsEmpty) Editor.Selection.Text = "";
        var pos = Editor.CaretPosition.InsertParagraphBreak();
        if (pos.Paragraph is { } p) DocHelpers.MarkPageBreak(p, true);
        Editor.CaretPosition = pos;
        Editor.EndChange();
        Editor.Focus();
    }

    // ================================================================ sag tik menusu

    private void Editor_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        var menu = Editor.ContextMenu!;
        menu.Items.Clear();

        // Yazim onerileri (Windows yazim denetimi, Turkce sozluk yukluyse)
        if (Editor.GetSpellingError(Editor.CaretPosition) is { } error)
        {
            int n = 0;
            foreach (var s in error.Suggestions)
            {
                var item = new MenuItem { Header = s, FontWeight = FontWeights.SemiBold };
                var suggestion = s;
                item.Click += (_, _) => error.Correct(suggestion);
                menu.Items.Add(item);
                if (++n == 5) break;
            }
            if (n == 0) menu.Items.Add(new MenuItem { Header = "(öneri yok)", IsEnabled = false });
            var ignore = new MenuItem { Header = "Tümünü yoksay" };
            ignore.Click += (_, _) => error.IgnoreAll();
            menu.Items.Add(ignore);
            menu.Items.Add(new Separator());
        }

        menu.Items.Add(new MenuItem { Header = "Kes", Command = ApplicationCommands.Cut, CommandTarget = Editor, InputGestureText = "Ctrl+X" });
        menu.Items.Add(new MenuItem { Header = "Kopyala", Command = ApplicationCommands.Copy, CommandTarget = Editor, InputGestureText = "Ctrl+C" });
        menu.Items.Add(new MenuItem { Header = "Yapıştır", Command = ApplicationCommands.Paste, CommandTarget = Editor, InputGestureText = "Ctrl+V" });

        // Baglanti
        var linkEl = FindAncestor<Hyperlink>(Editor.CaretPosition.Parent);
        if (linkEl != null)
        {
            menu.Items.Add(new Separator());
            var open = new MenuItem { Header = "Bağlantıyı aç" };
            open.Click += (_, _) => OpenLink(linkEl.NavigateUri);
            var remove = new MenuItem { Header = "Bağlantıyı kaldır" };
            remove.Click += (_, _) => RemoveLink(linkEl);
            menu.Items.Add(open);
            menu.Items.Add(remove);
        }

        // Tablo
        if (TableOps.Locate(Editor.CaretPosition) is { } loc)
        {
            menu.Items.Add(new Separator());
            var pos = Editor.CaretPosition;
            AddTableItem(menu, "Üste satır ekle", () => TableOps.InsertRow(pos, false));
            AddTableItem(menu, "Alta satır ekle", () => TableOps.InsertRow(pos, true));
            AddTableItem(menu, "Sola sütun ekle", () => TableOps.InsertColumn(pos, false));
            AddTableItem(menu, "Sağa sütun ekle", () => TableOps.InsertColumn(pos, true));
            AddTableItem(menu, "Satırı sil", () => TableOps.DeleteRow(pos));
            AddTableItem(menu, "Sütunu sil", () => TableOps.DeleteColumn(pos));
            AddTableItem(menu, "Tabloyu sil", () => TableOps.DeleteTable(loc.table));
        }
    }

    private void AddTableItem(ContextMenu menu, string header, Action action)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) =>
        {
            Editor.BeginChange();
            action();
            Editor.EndChange();
            MarkDirty();
        };
        menu.Items.Add(item);
    }

    private void RemoveLink(Hyperlink link)
    {
        Editor.BeginChange();
        var inlines = link.Inlines.ToList();
        if (link.Parent is Paragraph para)
        {
            foreach (var i in inlines) { link.Inlines.Remove(i); para.Inlines.InsertBefore(link, i); }
            para.Inlines.Remove(link);
        }
        else if (link.Parent is Span span)
        {
            foreach (var i in inlines) { link.Inlines.Remove(i); span.Inlines.InsertBefore(link, i); }
            span.Inlines.Remove(link);
        }
        Editor.EndChange();
    }

    private static T? FindAncestor<T>(DependencyObject? d) where T : class
    {
        while (d != null)
        {
            if (d is T t) return t;
            d = d is TextElement te ? te.Parent : null;
        }
        return null;
    }

    // ================================================================ sayfa duzeni

    private void Margins_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = MarginsBtn, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        AddPageItem(menu, "Normal (2,5 cm)", p => p.SetMargins(2.5));
        AddPageItem(menu, "Dar (1,27 cm)", p => p.SetMargins(1.27));
        AddPageItem(menu, "Orta (üst/alt 2,54, sol/sağ 1,91 cm)", p =>
        {
            p.MarginTop = p.MarginBottom = Units.CmToPx(2.54);
            p.MarginLeft = p.MarginRight = Units.CmToPx(1.91);
        });
        AddPageItem(menu, "Geniş (üst/alt 2,54, sol/sağ 5,08 cm)", p =>
        {
            p.MarginTop = p.MarginBottom = Units.CmToPx(2.54);
            p.MarginLeft = p.MarginRight = Units.CmToPx(5.08);
        });
        menu.IsOpen = true;
    }

    private void Orientation_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = OrientationBtn, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        AddPageItem(menu, "Dikey", p => { if (p.IsLandscape) (p.Width, p.Height) = (p.Height, p.Width); });
        AddPageItem(menu, "Yatay", p => { if (!p.IsLandscape) (p.Width, p.Height) = (p.Height, p.Width); });
        menu.IsOpen = true;
    }

    private void PageSize_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = SizeBtn, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        void Size(string name, double wCm, double hCm) => AddPageItem(menu, name, p =>
        {
            bool land = p.IsLandscape;
            p.Width = Units.CmToPx(land ? hCm : wCm);
            p.Height = Units.CmToPx(land ? wCm : hCm);
        });
        Size("A4 (21 × 29,7 cm)", 21, 29.7);
        Size("A5 (14,8 × 21 cm)", 14.8, 21);
        Size("A3 (29,7 × 42 cm)", 29.7, 42);
        Size("Letter (21,59 × 27,94 cm)", 21.59, 27.94);
        Size("Legal (21,59 × 35,56 cm)", 21.59, 35.56);
        menu.IsOpen = true;
    }

    private void AddPageItem(ContextMenu menu, string header, Action<PageSettings> change)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) =>
        {
            change(_page);
            ApplyPage();
            _dirty = true;
            UpdateTitle();
        };
        menu.Items.Add(item);
    }

    // ================================================================ yakinlastirma

    private void SetZoom(double scale)
    {
        scale = Clamp(scale, 0.25, 3.0);
        ZoomTransform.ScaleX = ZoomTransform.ScaleY = scale;
        ZoomText.Text = $"%{Math.Round(scale * 100)}";
        if (Math.Abs(ZoomSlider.Value - scale * 100) > 0.5) ZoomSlider.Value = scale * 100;
    }

    private void Zoom100_Click(object sender, RoutedEventArgs e) => SetZoom(1);
    private void ZoomIn_Click(object sender, RoutedEventArgs e) => SetZoom(ZoomTransform.ScaleX + 0.1);
    private void ZoomOut_Click(object sender, RoutedEventArgs e) => SetZoom(ZoomTransform.ScaleX - 0.1);
    private void ZoomWidth_Click(object sender, RoutedEventArgs e) => SetZoom((Scroller.ViewportWidth - 60) / _page.Width);

    // ================================================================ bul / degistir

    private void Find_Click(object sender, RoutedEventArgs e) => ShowFind(false);
    private void Replace_Click(object sender, RoutedEventArgs e) => ShowFind(true);

    private void ShowFind(bool replace)
    {
        if (_findWindow == null || !_findWindow.IsLoaded)
        {
            _findWindow = new FindReplaceWindow(this, Editor);
            _findWindow.Closed += (_, _) => _findWindow = null;
        }
        _findWindow.ShowMode(replace);
    }

    // ================================================================ Dosya menusu ve kisayollar

    private void Save_Click(object sender, RoutedEventArgs e) => Save();

    private void FileButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = FileButton, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom, MinWidth = 240 };
        MenuItem Item(string header, string? gesture, Action action, string? icon = null)
        {
            var mi = new MenuItem { Header = header, InputGestureText = gesture ?? "" };
            if (icon != null) mi.Icon = new TextBlock { Text = icon, FontFamily = (FontFamily)FindResource("IconFont"), FontSize = 14 };
            mi.Click += (_, _) => action();
            menu.Items.Add(mi);
            return mi;
        }

        Item("Yeni belge", "Ctrl+N", AppActions.NewDocument, "");
        Item("Aç...", "Ctrl+O", () => AppActions.ShowOpenDialog(this), "");
        var recent = new MenuItem { Header = "Son kullanılanlar" };
        foreach (var r in RecentFiles.Load().Take(10))
        {
            var mi = new MenuItem { Header = r.FileName.Replace("_", "__"), ToolTip = r.Path };
            var path = r.Path;
            mi.Click += (_, _) => AppActions.OpenFile(path, this);
            recent.Items.Add(mi);
        }
        recent.IsEnabled = recent.Items.Count > 0;
        menu.Items.Add(recent);
        menu.Items.Add(new Separator());
        Item("Kaydet", "Ctrl+S", () => Save(), "");
        Item("Farklı kaydet...", "F12", () => SaveAs());
        menu.Items.Add(new Separator());
        Item("Yazdır...", "Ctrl+P", Print, "");
        Item("PDF olarak dışa aktar...", null, () => Printing.Print(Editor.Document, _page, DisplayName, true), "");
        menu.Items.Add(new Separator());
        Item("Başlangıç ekranı", null, () => new StartWindow().Show(), "");
        Item("Hakkında", null, () => MessageBox.Show(
            $"{AppInfo.Name} {AppInfo.Version}\n\nÜcretsiz ve açık kaynak ofis programı (MIT lisansı).\nGeliştiren: SingleKnight",
            "Hakkında", MessageBoxButton.OK, MessageBoxImage.Information));
        Item("Kapat", "Alt+F4", Close);
        menu.IsOpen = true;
    }

    private void Print() => Printing.Print(Editor.Document, _page, DisplayName, false);

    private void AddShortcuts()
    {
        void Bind(Key key, ModifierKeys mods, Action action) => InputBindings.Add(new KeyBinding(new RelayCommand(action), key, mods));
        Bind(Key.N, ModifierKeys.Control, AppActions.NewDocument);
        Bind(Key.O, ModifierKeys.Control, () => AppActions.ShowOpenDialog(this));
        Bind(Key.S, ModifierKeys.Control, () => Save());
        Bind(Key.S, ModifierKeys.Control | ModifierKeys.Shift, () => SaveAs());
        Bind(Key.F12, ModifierKeys.None, () => SaveAs());
        Bind(Key.P, ModifierKeys.Control, Print);
        Bind(Key.F, ModifierKeys.Control, () => ShowFind(false));
        Bind(Key.H, ModifierKeys.Control, () => ShowFind(true));
        Bind(Key.K, ModifierKeys.Control, InsertLink);
        Bind(Key.D1, ModifierKeys.Control, () => SetLineSpacing(1));
        Bind(Key.D2, ModifierKeys.Control, () => SetLineSpacing(2));
        Bind(Key.D5, ModifierKeys.Control, () => SetLineSpacing(1.5));
        Bind(Key.D1, ModifierKeys.Control | ModifierKeys.Alt, () => ApplyParagraphStyle(StyleCatalog.Heading1));
        Bind(Key.D2, ModifierKeys.Control | ModifierKeys.Alt, () => ApplyParagraphStyle(StyleCatalog.Heading2));
        Bind(Key.D3, ModifierKeys.Control | ModifierKeys.Alt, () => ApplyParagraphStyle(StyleCatalog.Heading3));
        Bind(Key.N, ModifierKeys.Control | ModifierKeys.Shift, () => ApplyParagraphStyle(StyleCatalog.Normal));
        Bind(Key.OemPlus, ModifierKeys.Control, () => ToggleBaseline(BaselineAlignment.Subscript));
        Bind(Key.OemPlus, ModifierKeys.Control | ModifierKeys.Shift, () => ToggleBaseline(BaselineAlignment.Superscript));
    }
}
