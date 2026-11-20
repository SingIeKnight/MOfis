using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ClosedXML.Excel;
using Microsoft.Win32;
using MOfis.Common;
using MOfis.Shell;
using MOfis.Writer;

namespace MOfis.Sheet;

/// Hesap tablosu penceresi (MOfis Hesap).
public partial class SheetWindow : Window, IDocumentWindow
{
    private static int _untitledCounter;
    private static readonly CultureInfo Tr = new("tr-TR");

    private XLWorkbook _wb;
    private bool _dirty;
    private readonly UndoStack _undo = new();
    private IReadOnlyList<string> _warnings = [];
    private bool _warningsAccepted;
    private readonly string _untitledName;
    private readonly RecoveryEntry _recovery = new() { Id = Guid.NewGuid().ToString("N"), Kind = "sheet" };
    private readonly DispatcherTimer _autosave = new() { Interval = TimeSpan.FromMinutes(2) };
    private readonly Dictionary<string, (int Row, int Col, int Top, int Left)> _sheetViews = [];

    // Duzenleme durumu
    private bool _editing;
    private bool _editInBar;
    private bool _editByTyping;
    private string _editOriginal = "";
    private bool _syncText;
    private bool _syncScroll;

    // Pano: son kopyalanan aralik (ayni metin yapistirilirsa formuller ve bicimlerle birlikte kopyalanir).
    private (string Sheet, CellRange Range, bool Cut, string Text)? _clip;

    public string? FilePath { get; private set; }

    private IXLWorksheet Sheet => View.Sheet!;

    public SheetWindow()
    {
        InitializeComponent();
        _untitledName = "Kitap" + (++_untitledCounter);
        _wb = NewWorkbook();

        View.SelectionChanged += View_SelectionChanged;
        View.ScrollChanged += SyncScrollBars;
        View.EditRequested += text => BeginEdit(text);
        View.DeleteRequested += ClearContents;
        View.ZoomRequested += z => ZoomSlider.Value = Math.Round(z * 100);
        View.PreviewMouseLeftButtonDown += View_PreviewMouseLeftButtonDown;
        View.SizeChanged += (_, _) => PositionEditor();

        VScroll.Scroll += (_, _) => { if (!_syncScroll) View.ScrollTo((int)Math.Round(VScroll.Value), View.LeftCol); };
        HScroll.Scroll += (_, _) => { if (!_syncScroll) View.ScrollTo(View.TopRow, (int)Math.Round(HScroll.Value)); };
        ZoomSlider.ValueChanged += (_, e) =>
        {
            View.SetZoom(e.NewValue / 100.0);
            ZoomText.Text = $"%{e.NewValue:0}";
            PositionEditor();
        };
        ZoomText.Text = "%100";

        CellEditor.PreviewKeyDown += CellEditor_PreviewKeyDown;
        CellEditor.TextChanged += (_, _) =>
        {
            if (_syncText) return;
            _syncText = true;
            FormulaBox.Text = CellEditor.Text;
            _syncText = false;
            PositionEditor();
        };
        FormulaBox.PreviewKeyDown += FormulaBox_PreviewKeyDown;
        FormulaBox.GotKeyboardFocus += (_, _) =>
        {
            if (!_editing) BeginEdit(null, inBar: true);
        };
        FormulaBox.TextChanged += (_, _) =>
        {
            if (_syncText || !_editing) return;
            _syncText = true;
            CellEditor.Text = FormulaBox.Text;
            _syncText = false;
        };
        NameBox.PreviewKeyDown += NameBox_PreviewKeyDown;
        NameBox.GotKeyboardFocus += (_, _) => NameBox.SelectAll();

        BindWorkbook(0);

        _autosave.Tick += (_, _) => { if (_dirty && !_editing) AutoRecovery.SaveSheet(Recovery(), _wb); };
        _autosave.Start();

        AddShortcuts();
        Loaded += (_, _) => View.Focus();
        Closing += Window_Closing;
        Closed += (_, _) =>
        {
            _autosave.Stop();
            AutoRecovery.Delete(_recovery.Id);
            _wb.Dispose();
        };
        UpdateTitle();
    }

    private static XLWorkbook NewWorkbook()
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sayfa1");
        return wb;
    }

    // ================================================================ kitap / sayfa baglama

    private void BindWorkbook(int sheetIndex)
    {
        _sheetViews.Clear();
        sheetIndex = Math.Clamp(sheetIndex, 0, _wb.Worksheets.Count - 1);
        ShowSheet(_wb.Worksheet(sheetIndex + 1));
    }

    private int ActiveSheetIndex => Sheet.Position - 1;

    private void ShowSheet(IXLWorksheet ws)
    {
        if (View.Sheet != null && View.Sheet != ws)
            _sheetViews[View.Sheet.Name] = (View.ActiveRow, View.ActiveCol, View.TopRow, View.LeftCol);
        View.Bind(ws);
        foreach (var w in _wb.Worksheets)
        {
            w.TabActive = w == ws;
            w.TabSelected = w == ws;
        }
        if (_sheetViews.TryGetValue(ws.Name, out var v))
        {
            View.ScrollTo(v.Top, v.Left);
            View.Select(v.Row, v.Col);
        }
        RebuildTabs();
    }

    private void RebuildTabs()
    {
        SheetTabs.Children.Clear();
        foreach (var ws in _wb.Worksheets.OrderBy(w => w.Position))
        {
            bool active = ws == View.Sheet;
            var text = new TextBlock
            {
                Text = ws.Name,
                FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = active ? (Brush)FindResource("AccentBrush") : (Brush)FindResource("TextBrush"),
            };
            var tab = new Border
            {
                Child = text,
                Padding = new Thickness(14, 4, 14, 5),
                Background = active ? Brushes.White : Brushes.Transparent,
                BorderBrush = active ? (Brush)FindResource("AccentBrush") : (Brush)FindResource("LineBrush"),
                BorderThickness = active ? new Thickness(0, 0, 0, 2) : new Thickness(0, 0, 1, 0),
                Cursor = Cursors.Hand,
                ToolTip = "Çift tıklayarak yeniden adlandırın; sağ tıklayarak diğer işlemler",
            };
            var sheet = ws;
            tab.MouseLeftButtonDown += (_, e) =>
            {
                if (!CommitIfEditing()) return;
                if (e.ClickCount == 2) RenameSheet(sheet);
                else if (sheet != View.Sheet) ShowSheet(sheet);
                e.Handled = true;
            };
            var menu = new ContextMenu();
            void Item(string header, Action a, bool enabled = true)
            {
                var mi = new MenuItem { Header = header, IsEnabled = enabled };
                mi.Click += (_, _) => a();
                menu.Items.Add(mi);
            }
            Item("Yeniden adlandır...", () => RenameSheet(sheet));
            Item("Sola taşı", () => MoveSheet(sheet, -1), sheet.Position > 1);
            Item("Sağa taşı", () => MoveSheet(sheet, 1), sheet.Position < _wb.Worksheets.Count);
            menu.Items.Add(new Separator());
            Item("Sil", () => DeleteSheet(sheet), _wb.Worksheets.Count > 1);
            tab.ContextMenu = menu;
            SheetTabs.Children.Add(tab);
        }
    }

    private void AddSheet_Click(object sender, RoutedEventArgs e) => AddSheet();

    private void AddSheet()
    {
        if (!CommitIfEditing()) return;
        int n = _wb.Worksheets.Count + 1;
        while (_wb.Worksheets.Any(w => string.Equals(w.Name, "Sayfa" + n, StringComparison.OrdinalIgnoreCase))) n++;
        var ws = _wb.AddWorksheet("Sayfa" + n, ActiveSheetIndex + 2);
        MarkDirty();
        ShowSheet(ws);
        View.Focus();
    }

    private void RenameSheet(IXLWorksheet ws)
    {
        var name = InputDialog.Ask(this, "Sayfayı yeniden adlandır", "Yeni sayfa adı:", ws.Name);
        if (name == null) return;
        name = SheetIO.SafeSheetName(name);
        if (name == ws.Name) return;
        if (_wb.Worksheets.Any(w => w != ws && string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show($"\"{name}\" adında bir sayfa zaten var.", AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        try
        {
            ws.Name = name;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Sayfa adı kullanılamıyor:\n\n" + ex.Message, AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        _undo.Clear(); // geri alma adimlari sayfa adiyla tutuluyor
        MarkDirty();
        RebuildTabs();
    }

    private void MoveSheet(IXLWorksheet ws, int delta)
    {
        int p = ws.Position + delta;
        if (p < 1 || p > _wb.Worksheets.Count) return;
        ws.Position = p;
        MarkDirty();
        RebuildTabs();
    }

    private void DeleteSheet(IXLWorksheet ws)
    {
        if (_wb.Worksheets.Count <= 1) return;
        if (MessageBox.Show($"\"{ws.Name}\" sayfası ve içindeki her şey silinsin mi? Bu işlem geri alınamaz.", AppInfo.SheetName,
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
        int pos = ws.Position;
        bool wasActive = ws == View.Sheet;
        _sheetViews.Remove(ws.Name);
        ws.Delete();
        _undo.Clear();
        _clip = null;
        MarkDirty();
        if (wasActive) ShowSheet(_wb.Worksheet(Math.Min(pos, _wb.Worksheets.Count)));
        else RebuildTabs();
    }

    private void SwitchSheetBy(int delta)
    {
        if (!CommitIfEditing()) return;
        int p = Sheet.Position + delta;
        if (p < 1 || p > _wb.Worksheets.Count) return;
        ShowSheet(_wb.Worksheet(p));
    }

    // ================================================================ dosya

    public bool LoadFile(string path)
    {
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            var loaded = SheetIO.Load(path);
            ReplaceWorkbook(loaded.Workbook);
            FilePath = path;
            _warnings = loaded.Warnings;
            _warningsAccepted = false;
            RecentFiles.Add(path);
            UpdateTitle();
            if (_warnings.Count > 0)
                ShowInfo("Bu dosyadaki bazı öğeler henüz desteklenmiyor: " + string.Join(", ", _warnings) +
                         ". Aynı dosyanın üzerine kaydederseniz bu öğeler kaybolabilir; \"Farklı kaydet\" önerilir.");
            return true;
        }
        catch (Exception ex)
        {
            ErrorLog.Write(ex);
            MessageBox.Show($"Dosya açılamadı:\n{path}\n\n{ex.Message}", AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Error);
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
            var loaded = SheetIO.Load(entry.DocPath);
            ReplaceWorkbook(loaded.Workbook);
            _recovery.Id = entry.Id;
            FilePath = entry.OriginalPath != null && File.Exists(entry.OriginalPath) ? entry.OriginalPath : null;
            _dirty = true;
            UpdateTitle();
            ShowInfo("Bu tablo otomatik kurtarma kopyasından açıldı. Saklamak için kaydedin.");
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Kurtarma kopyası açılamadı:\n\n" + ex.Message, AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private void ReplaceWorkbook(XLWorkbook wb)
    {
        var old = _wb;
        _wb = wb;
        _undo.Clear();
        _clip = null;
        View.CopyMarquee = null;
        var active = wb.Worksheets.FirstOrDefault(w => w.TabActive) ?? wb.Worksheet(1);
        _sheetViews.Clear();
        ShowSheet(active);
        if (!ReferenceEquals(old, wb)) old.Dispose();
        _dirty = false;
        UpdateTitle();
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
        if (!CommitIfEditing()) return false;
        if (FilePath == null) return SaveAs();
        if (_warnings.Count > 0 && !_warningsAccepted)
        {
            var answer = MessageBox.Show(
                "Bu dosyada MOfis'in henüz desteklemediği öğeler vardı:\n\n• " + string.Join("\n• ", _warnings) +
                "\n\nAynı dosyanın üzerine kaydederseniz bu öğeler kaybolabilir.\n\n" +
                "Evet: yine de üzerine kaydet\nHayır: farklı bir dosyaya kaydet\nİptal: kaydetme",
                AppInfo.SheetName, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (answer == MessageBoxResult.Cancel) return false;
            if (answer == MessageBoxResult.No) return SaveAs();
            _warningsAccepted = true;
        }
        return SaveTo(FilePath);
    }

    private bool SaveAs()
    {
        if (!CommitIfEditing()) return false;
        var dlg = new SaveFileDialog
        {
            Filter = SheetIO.SaveFilter,
            FileName = Path.GetFileNameWithoutExtension(DisplayName),
            Title = "Farklı kaydet",
            AddExtension = true,
        };
        if (FilePath != null)
        {
            dlg.InitialDirectory = Path.GetDirectoryName(FilePath);
            dlg.FilterIndex = SheetIO.IsCsv(FilePath) ? 2 : 1;
        }
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
        if (SheetIO.IsCsv(path))
        {
            bool formulas = Sheet.CellsUsed(XLCellsUsedOptions.Contents).Any(c => c.HasFormula);
            if (_wb.Worksheets.Count > 1 || formulas)
            {
                var msg = "CSV dosyası yalnızca etkin sayfanın (\"" + Sheet.Name + "\") değerlerini saklar." +
                          (formulas ? " Formüller yerine sonuçları yazılır." : "") +
                          (_wb.Worksheets.Count > 1 ? " Diğer sayfalar kaydedilmez." : "") +
                          " Biçimlendirme de kaydedilmez.\n\nDevam edilsin mi?";
                if (MessageBox.Show(msg, AppInfo.SheetName, MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return false;
            }
        }
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            SheetIO.Save(_wb, path, ActiveSheetIndex);
            FilePath = path;
            _dirty = false;
            RecentFiles.Add(path);
            AutoRecovery.Delete(_recovery.Id);
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            ErrorLog.Write(ex);
            MessageBox.Show($"Dosya kaydedilemedi. Başka bir programda açık olabilir.\n\n{ex.Message}", AppInfo.SheetName,
                MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e) => Save();

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_editing && !CommitIfEditing()) { e.Cancel = true; return; }
        if (!_dirty) return;
        var answer = MessageBox.Show($"\"{DisplayName}\" tablosundaki değişiklikler kaydedilsin mi?", AppInfo.SheetName,
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (answer == MessageBoxResult.Cancel || (answer == MessageBoxResult.Yes && !Save())) e.Cancel = true;
    }

    private void UpdateTitle() => Title = $"{DisplayName}{(_dirty ? " •" : "")} - {AppInfo.SheetName}";

    private void MarkDirty()
    {
        if (_dirty) return;
        _dirty = true;
        UpdateTitle();
    }

    /// Gelistirme simulasyonunda (Silent) iletiler kutu acmadan gunluge yazilir.
    internal bool Silent { get; set; }

    private readonly List<string> _silentLog = [];

    private void Warn(string message)
    {
        if (Silent)
        {
            _silentLog.Add(message.Split('\n')[0]);
            return;
        }
        MessageBox.Show(message, AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private void ShowInfo(string text)
    {
        InfoText.Text = text;
        InfoBar.Visibility = Visibility.Visible;
    }

    private void CloseInfo_Click(object sender, RoutedEventArgs e) => InfoBar.Visibility = Visibility.Collapsed;

    // ================================================================ secim, kaydirma, durum cubugu

    private void View_SelectionChanged()
    {
        if (!NameBox.IsKeyboardFocusWithin)
        {
            var sel = View.Selection;
            var merge = View.MergeAt(View.ActiveRow, View.ActiveCol);
            NameBox.Text = sel.IsSingle || merge == sel
                ? XLHelper.GetColumnLetterFromNumber(View.ActiveCol) + View.ActiveRow
                : sel.Address();
        }
        if (!_editing)
        {
            _syncText = true;
            FormulaBox.Text = View.Sheet == null ? "" : CellInput.EditText(Sheet.Cell(View.ActiveRow, View.ActiveCol));
            _syncText = false;
        }
        UpdateStats();
    }

    private void UpdateStats()
    {
        var sel = View.Selection;
        if (View.Sheet == null || sel.IsSingle)
        {
            SumText.Text = "";
            return;
        }
        var (sum, nums, filled) = SheetOps.Stats(Sheet, sel);
        if (filled < 2) { SumText.Text = ""; return; }
        SumText.Text = nums > 0
            ? $"Ortalama: {SheetDisplay.FormatGeneral(sum / nums)}     Sayım: {filled}     Toplam: {SheetDisplay.FormatGeneral(sum)}"
            : $"Sayım: {filled}";
    }

    private void SyncScrollBars()
    {
        if (View.Sheet == null) return;
        _syncScroll = true;
        var last = Sheet.LastCellUsed(XLCellsUsedOptions.All);
        int rows = View.PageRows, cols = View.PageCols;
        VScroll.Maximum = Math.Max(last?.Address.RowNumber ?? 1, View.TopRow) + rows;
        VScroll.ViewportSize = rows;
        VScroll.LargeChange = rows;
        VScroll.Value = View.TopRow;
        HScroll.Maximum = Math.Max(last?.Address.ColumnNumber ?? 1, View.LeftCol) + cols;
        HScroll.ViewportSize = cols;
        HScroll.LargeChange = cols;
        HScroll.Value = View.LeftCol;
        _syncScroll = false;
        PositionEditor();
    }

    private void NameBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            var r = SheetOps.ParseAddress(NameBox.Text);
            if (r == null)
            {
                MessageBox.Show("Geçerli bir hücre adresi yazın (ör. B12, A1:C5, C:D veya 3:7).", AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            View.SelectRange(r.Value);
            View.EnsureVisible(r.Value.R1, r.Value.C1);
            View.Focus();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            View.Focus();
            View_SelectionChanged();
            e.Handled = true;
        }
    }

    // ================================================================ hucre duzenleme

    private void BeginEdit(string? initial, bool inBar = false)
    {
        if (View.Sheet == null) return;
        var cell = Sheet.Cell(View.ActiveRow, View.ActiveCol);
        _editOriginal = CellInput.EditText(cell);
        _editing = true;
        _editInBar = inBar;
        _editByTyping = initial != null;
        var text = initial ?? _editOriginal;

        _syncText = true;
        CellEditor.Text = text;
        if (!inBar) FormulaBox.Text = text;
        _syncText = false;

        var font = cell.Style.Font;
        CellEditor.FontFamily = new FontFamily(string.IsNullOrWhiteSpace(font.FontName) ? "Calibri" : font.FontName);
        CellEditor.FontWeight = font.Bold ? FontWeights.Bold : FontWeights.Normal;
        CellEditor.FontStyle = font.Italic ? FontStyles.Italic : FontStyles.Normal;
        CellEditor.Tag = Math.Max(1, font.FontSize) * 96.0 / 72.0;
        CellEditor.Visibility = Visibility.Visible;
        PositionEditor();
        if (!inBar)
        {
            CellEditor.Focus();
            CellEditor.CaretIndex = CellEditor.Text.Length;
        }
    }

    private void PositionEditor()
    {
        if (!_editing || View.Sheet == null) return;
        var rect = View.CellRect(View.ActiveRow, View.ActiveCol);
        double size = (CellEditor.Tag as double? ?? 15) * View.Zoom;
        CellEditor.FontSize = size;
        Canvas.SetLeft(CellEditor, rect.X - 2);
        Canvas.SetTop(CellEditor, rect.Y - 2);
        CellEditor.MinWidth = rect.Width + 3;
        CellEditor.MaxWidth = Math.Max(rect.Width + 3, EditLayer.ActualWidth - rect.X);
        CellEditor.MinHeight = rect.Height + 3;
    }

    private void EndEdit()
    {
        _editing = false;
        _editInBar = false;
        CellEditor.Visibility = Visibility.Collapsed;
        View.Focus();
        View_SelectionChanged();
    }

    private string EditText => _editInBar ? FormulaBox.Text : CellEditor.Text;

    /// Duzenlemeyi hucreye yazar; formul hataliysa duzenleme acik kalir ve false doner.
    private bool CommitEdit(int dr = 0, int dc = 0)
    {
        if (!_editing) return true;
        var text = EditText;
        if (text == _editOriginal)
        {
            EndEdit();
            if (dr != 0 || dc != 0) View.MoveActive(dr, dc);
            return true;
        }

        var parsed = CellInput.Parse(text);
        if (parsed.Kind == InputKind.Formula)
        {
            var error = SheetEngine.Validate(parsed.FormulaEn!);
            if (error != null)
            {
                Warn("Bu formülde bir sorun var:\n\n" + error +
                     "\n\nTürkçe yazımda bağımsız değişkenler noktalı virgülle ayrılır, ondalık ayırıcı virgüldür: =TOPLA(A1;B1;1,5)");
                FocusEditor();
                return false;
            }
        }

        int row = View.ActiveRow, col = View.ActiveCol;
        var range = new CellRange(row, col, row, col);
        if (!ApplyChange([range], () =>
            {
                var cell = Sheet.Cell(row, col);
                CellInput.Apply(cell, parsed);
                if (text.Contains('\n')) cell.Style.Alignment.WrapText = true;
            }))
        {
            FocusEditor();
            return false;
        }

        AutoWidenForNumber(row, col, parsed);
        EndEdit();
        if (dr != 0 || dc != 0) View.MoveActive(dr, dc);
        return true;
    }

    /// Excel gibi: elle genisligi ayarlanmamis sutuna sigmayan sayi/tarih girilince sutun genisler.
    private void AutoWidenForNumber(int row, int col, ParsedInput parsed)
    {
        if (parsed.Kind is not (InputKind.Number or InputKind.Date)) return;
        var column = Sheet.Column(col);
        if (column.Width != Sheet.ColumnWidth && Math.Abs(SheetGrid.CharsToPx(column.Width) - View.Cols.Default) > 0.5) return;
        double need = View.MeasureCell(Sheet.Cell(row, col));
        if (need <= View.Cols.Size(col)) return;
        column.Width = Math.Round((need - 5) / 7.0, 2);
        View.Cols.Set(col, SheetGrid.CharsToPx(column.Width));
        View.InvalidateVisual();
    }

    private void FocusEditor()
    {
        if (_editInBar) FormulaBox.Focus();
        else CellEditor.Focus();
    }

    private void CancelEdit()
    {
        if (!_editing) return;
        EndEdit();
    }

    /// Duzenleme aciksa once yazar. Yazilamadiysa (hatali formul) false.
    private bool CommitIfEditing() => !_editing || CommitEdit();

    private void CellEditor_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        bool alt = (Keyboard.Modifiers & ModifierKeys.Alt) != 0 || e.Key == Key.System;
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        switch (key)
        {
            case Key.Enter when alt:
            {
                // Alt+Enter: hucre icinde yeni satir
                int i = CellEditor.CaretIndex;
                CellEditor.Text = CellEditor.Text.Insert(i, "\n");
                CellEditor.CaretIndex = i + 1;
                e.Handled = true;
                return;
            }
            case Key.Enter: CommitEdit(shift ? -1 : 1, 0); break;
            case Key.Tab: CommitEdit(0, shift ? -1 : 1); break;
            case Key.Escape: CancelEdit(); break;
            case Key.Up when _editByTyping: CommitEdit(-1, 0); break;
            case Key.Down when _editByTyping: CommitEdit(1, 0); break;
            case Key.Left when _editByTyping && !IsFormulaText(CellEditor.Text): CommitEdit(0, -1); break;
            case Key.Right when _editByTyping && !IsFormulaText(CellEditor.Text): CommitEdit(0, 1); break;
            case Key.F2:
                _editByTyping = !_editByTyping;
                break;
            default: return;
        }
        e.Handled = true;
    }

    private static bool IsFormulaText(string s) => s.StartsWith('=');

    private void FormulaBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        switch (e.Key)
        {
            case Key.Enter: CommitEdit(shift ? -1 : 1, 0); break;
            case Key.Tab: CommitEdit(0, shift ? -1 : 1); break;
            case Key.Escape: CancelEdit(); break;
            default: return;
        }
        e.Handled = true;
    }

    private void View_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_editing && !CommitEdit()) e.Handled = true;
    }

    // ================================================================ degisiklik + geri al

    /// Araliklarin icerigini anlik goruntuyle saklayip degisikligi uygular. Dongusel basvuru olursa
    /// degisiklik geri alinir, motor sifirlanir ve false doner.
    private bool ApplyChange(IReadOnlyList<CellRange> ranges, Action change)
    {
        var sheet = Sheet;
        var addresses = new List<(int R, int C)>();
        foreach (var r in ranges)
        {
            if (r.Count <= 20000)
            {
                for (int i = r.R1; i <= r.R2; i++)
                for (int j = r.C1; j <= r.C2; j++) addresses.Add((i, j));
            }
            else if (SheetOps.ClipToUsed(sheet, r) is { } u)
            {
                foreach (var c in sheet.Range(u.R1, u.C1, u.R2, u.C2).CellsUsed(XLCellsUsedOptions.All))
                    addresses.Add((c.Address.RowNumber, c.Address.ColumnNumber));
            }
        }
        addresses = addresses.Distinct().ToList();
        var before = addresses.Select(a => CellSnap.Take(sheet.Cell(a.R, a.C))).ToList();

        try
        {
            change();
        }
        catch (Exception ex)
        {
            ErrorLog.Write(ex);
            RestoreAfterFailure(sheet.Name, before);
            MessageBox.Show("Değişiklik uygulanamadı:\n\n" + ex.Message, AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        // Yeni formul dongusel basvuru olusturduysa geri al.
        foreach (var (r, c) in addresses)
        {
            var cell = Sheet.Cell(r, c);
            if (cell.HasFormula && SheetEngine.HasCycle(cell))
            {
                RestoreAfterFailure(sheet.Name, before);
                Warn($"Döngüsel başvuru: {XLHelper.GetColumnLetterFromNumber(c)}{r} hücresindeki formül, doğrudan ya da dolaylı olarak kendi sonucuna başvuruyor.\n\nFormülü düzeltin ya da Esc ile vazgeçin.");
                return false;
            }
        }

        var after = addresses.Select(a => CellSnap.Take(Sheet.Cell(a.R, a.C))).ToList();
        _undo.Push(new UndoStep(sheet.Name, before, after));
        MarkDirty();
        AfterContentChange();
        return true;
    }

    /// Hata sonrasi: motoru sifirla (kitabi bellekte yeniden ac) ve eski icerigi geri yaz.
    private void RestoreAfterFailure(string sheetName, List<CellSnap> before)
    {
        int idx = ActiveSheetIndex;
        var (ar, ac, top, left) = (View.ActiveRow, View.ActiveCol, View.TopRow, View.LeftCol);
        var fresh = SheetEngine.Rebuild(_wb);
        _wb = fresh;
        var ws = _wb.Worksheet(sheetName);
        foreach (var s in before) s.Restore(ws);
        _sheetViews.Clear();
        View.Bind(_wb.Worksheet(idx + 1));
        View.ScrollTo(top, left);
        View.Select(ar, ac);
        RebuildTabs();
        _clip = null;
    }

    private void AfterContentChange()
    {
        View.InvalidateVisual();
        UpdateStats();
        SyncScrollBars();
    }

    private void Undo_Click(object sender, RoutedEventArgs e) => UndoRedo(undo: true);

    private void Redo_Click(object sender, RoutedEventArgs e) => UndoRedo(undo: false);

    private void UndoRedo(bool undo)
    {
        if (_editing) { CancelEdit(); return; }
        var step = undo ? _undo.Undo() : _undo.Redo();
        if (step == null) return;
        if (!_wb.TryGetWorksheet(step.Sheet, out var ws)) return;
        if (ws != View.Sheet) ShowSheet(ws);
        var snaps = undo ? step.Before : step.After;
        try
        {
            foreach (var s in snaps) s.Restore(ws);
        }
        catch (Exception ex)
        {
            ErrorLog.Write(ex);
        }
        if (snaps.Count > 0)
        {
            var r1 = snaps.Min(s => s.Row); var r2 = snaps.Max(s => s.Row);
            var c1 = snaps.Min(s => s.Col); var c2 = snaps.Max(s => s.Col);
            View.SelectRange(new CellRange(r1, c1, r2, c2));
            View.EnsureVisible(r1, c1);
        }
        MarkDirty();
        AfterContentChange();
    }

    // ================================================================ pano, temizle, otomatik toplam

    private void ClearContents()
    {
        if (!CommitIfEditing()) return;
        var sel = View.Selection;
        var target = sel.Count > 20000 ? SheetOps.ClipToUsed(Sheet, sel) : sel;
        if (target == null) return;
        ApplyChange([target.Value], () =>
        {
            if (target.Value.Count > 20000)
            {
                foreach (var c in Sheet.Range(target.Value.R1, target.Value.C1, target.Value.R2, target.Value.C2).CellsUsed(XLCellsUsedOptions.Contents))
                    c.Clear(XLClearOptions.Contents);
            }
            else
            {
                Sheet.Range(target.Value.R1, target.Value.C1, target.Value.R2, target.Value.C2).Clear(XLClearOptions.Contents);
            }
        });
    }

    private void Clear_Click(object sender, RoutedEventArgs e) => ClearContents();

    private void Copy_Click(object sender, RoutedEventArgs e) => CopySelection(cut: false);

    private void Cut_Click(object sender, RoutedEventArgs e) => CopySelection(cut: true);

    private void Paste_Click(object sender, RoutedEventArgs e) => PasteClipboard();

    private void CopySelection(bool cut)
    {
        if (_editing) return;
        var range = SheetOps.ClipToUsed(Sheet, View.Selection) ?? new CellRange(View.ActiveRow, View.ActiveCol, View.ActiveRow, View.ActiveCol);
        if (range.Count > 1_000_000)
        {
            MessageBox.Show("Seçim panoya kopyalanamayacak kadar büyük.", AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var text = SheetOps.ToClipboardText(Sheet, range);
        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Pano kullanılamadı:\n\n" + ex.Message, AppInfo.SheetName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        _clip = (Sheet.Name, range, cut, text);
        View.CopyMarquee = range;
        View.InvalidateVisual();
        StatusText.Text = cut ? "Taşımak için hedefi seçip Ctrl+V'ye basın" : "Yapıştırmak için hedefi seçip Ctrl+V'ye basın";
    }

    private void PasteClipboard()
    {
        if (_editing) return;
        string text;
        try
        {
            if (!Clipboard.ContainsText()) return;
            text = Clipboard.GetText();
        }
        catch
        {
            return;
        }
        int row = View.ActiveRow, col = View.ActiveCol;

        if (_clip is { } clip && clip.Text == text && _wb.TryGetWorksheet(clip.Sheet, out var src))
        {
            var from = clip.Range;
            var to = new CellRange(row, col, Math.Min(SheetGrid.MaxRow, row + from.R2 - from.R1), Math.Min(SheetGrid.MaxCol, col + from.C2 - from.C1));
            var ranges = new List<CellRange> { to };
            bool sameSheet = src == Sheet;
            if (clip.Cut && sameSheet) ranges.Add(from);
            bool ok = ApplyChange(ranges, () =>
            {
                SheetOps.CopyCells(src, from, Sheet, row, col);
                if (clip.Cut && sameSheet)
                {
                    for (int r = from.R1; r <= from.R2; r++)
                    for (int c = from.C1; c <= from.C2; c++)
                        if (!to.Contains(r, c)) Sheet.Cell(r, c).Clear();
                }
            });
            if (clip.Cut && !sameSheet && ok)
            {
                // Baska sayfadan tasima: kaynak ayri bir adimda temizlenir.
                var s = Sheet;
                ShowSheet(src);
                ApplyChange([from], () => src.Range(from.R1, from.C1, from.R2, from.C2).Clear());
                ShowSheet(s);
            }
            if (ok)
            {
                View.SelectRange(to);
                if (clip.Cut) { _clip = null; View.CopyMarquee = null; }
            }
        }
        else
        {
            var rows = SheetIO.ParseClipboard(text);
            if (rows.Count == 0) return;
            int width = rows.Max(r => r.Count);
            var to = new CellRange(row, col, Math.Min(SheetGrid.MaxRow, row + rows.Count - 1), Math.Min(SheetGrid.MaxCol, col + width - 1));
            if (ApplyChange([to], () => SheetOps.PasteText(Sheet, rows, row, col))) View.SelectRange(to);
        }
        StatusText.Text = "Hazır";
    }

    private void AutoSum_Click(object sender, RoutedEventArgs e)
    {
        if (!CommitIfEditing()) return;
        var f = SheetOps.AutoSumFormula(Sheet, View.ActiveRow, View.ActiveCol) ?? "=TOPLA()";
        BeginEdit(f);
        if (f == "=TOPLA()") CellEditor.CaretIndex = f.Length - 1;
    }

    // ================================================================ menu, kisayollar

    private void FileButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = FileButton, Placement = PlacementMode.Bottom, MinWidth = 240 };
        void Item(string header, string? gesture, Action action, string? icon = null)
        {
            var mi = new MenuItem { Header = header, InputGestureText = gesture ?? "" };
            if (icon != null) mi.Icon = new TextBlock { Text = icon, FontFamily = (FontFamily)FindResource("IconFont"), FontSize = 14 };
            mi.Click += (_, _) => action();
            menu.Items.Add(mi);
        }

        Item("Yeni tablo", "Ctrl+N", AppActions.NewSheet, "");
        Item("Yeni belge", null, AppActions.NewDocument, "");
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
        Item("Başlangıç ekranı", null, () => new StartWindow().Show(), "");
        Item("Hakkında", null, () => MessageBox.Show(
            $"{AppInfo.Name} {AppInfo.Version}\n\nÜcretsiz ve açık kaynak ofis programı (MIT lisansı).\nGeliştiren: SingleKnight",
            "Hakkında", MessageBoxButton.OK, MessageBoxImage.Information));
        Item("Kapat", "Alt+F4", Close);
        menu.IsOpen = true;
    }

    private void AddShortcuts()
    {
        void Bind(Key key, ModifierKeys mods, Action action) => InputBindings.Add(new KeyBinding(new RelayCommand(action), key, mods));
        Bind(Key.N, ModifierKeys.Control, AppActions.NewSheet);
        Bind(Key.O, ModifierKeys.Control, () => AppActions.ShowOpenDialog(this));
        Bind(Key.S, ModifierKeys.Control, () => Save());
        Bind(Key.S, ModifierKeys.Control | ModifierKeys.Shift, () => SaveAs());
        Bind(Key.F12, ModifierKeys.None, () => SaveAs());
        Bind(Key.F11, ModifierKeys.Shift, AddSheet);
        Bind(Key.PageDown, ModifierKeys.Control, () => SwitchSheetBy(1));
        Bind(Key.PageUp, ModifierKeys.Control, () => SwitchSheetBy(-1));

        // Pano ve geri al yalnizca izgara odaktayken (metin kutularinda kendi davranislari gecerli).
        void GridBind(Key key, ModifierKeys mods, Action action) => View.InputBindings.Add(new KeyBinding(new RelayCommand(action), key, mods));
        GridBind(Key.C, ModifierKeys.Control, () => CopySelection(cut: false));
        GridBind(Key.Insert, ModifierKeys.Control, () => CopySelection(cut: false));
        GridBind(Key.X, ModifierKeys.Control, () => CopySelection(cut: true));
        GridBind(Key.V, ModifierKeys.Control, PasteClipboard);
        GridBind(Key.Insert, ModifierKeys.Shift, PasteClipboard);
        GridBind(Key.Z, ModifierKeys.Control, () => UndoRedo(undo: true));
        GridBind(Key.Y, ModifierKeys.Control, () => UndoRedo(undo: false));
        GridBind(Key.Escape, ModifierKeys.None, () =>
        {
            _clip = null;
            View.CopyMarquee = null;
            View.InvalidateVisual();
            StatusText.Text = "Hazır";
        });
    }

    // ================================================================ gelistirme araci

    /// --shot ile: yuklenen tabloyu ve ilk hucrelerin gorunen metnini yazar.
    internal string DescribeLayout()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Sayfalar: {string.Join(", ", _wb.Worksheets.Select(w => w.Name))}  etkin={Sheet.Name}");
        sb.AppendLine($"Secim={View.Selection.Address()} Top={View.TopRow} Left={View.LeftCol} Zoom={View.Zoom}");
        var used = Sheet.RangeUsed();
        if (used != null)
        {
            int r2 = Math.Min(used.LastRow().RowNumber(), 25), c2 = Math.Min(used.LastColumn().ColumnNumber(), 10);
            for (int r = 1; r <= r2; r++)
                sb.AppendLine(string.Join(" | ", Enumerable.Range(1, c2).Select(c =>
                {
                    var cell = Sheet.Cell(r, c);
                    return cell.HasFormula ? $"{SheetDisplay.Text(cell)} [{CellInput.EditText(cell)}]" : SheetDisplay.Text(cell);
                })));
        }
        return sb.ToString();
    }

    /// Gelistirme: klavye girisini taklit eder (MOFIS_SHEETTEST=1 ile --shot).
    internal string SimulateTyping()
    {
        Silent = true;
        var log = new System.Text.StringBuilder();
        void Type(string addr, string text)
        {
            var r = SheetOps.ParseAddress(addr)!.Value;
            View.Select(r.R1, r.C1);
            BeginEdit(text);
            bool ok = CommitEdit(1, 0);
            log.AppendLine($"{addr} <- {text} : {(ok ? "OK" : "RED")} => {SheetDisplay.Text(Sheet.Cell(r.R1, r.C1))}");
        }
        Type("A1", "Ürün");
        Type("B1", "Tutar");
        Type("A2", "Elma");
        Type("B2", "12,5");
        Type("A3", "Armut");
        Type("B3", "7,25");
        Type("A4", "Toplam");
        Type("B4", "=TOPLA(B2:B3)");
        Type("C4", "=EĞER(B4>15;\"yüksek\";\"düşük\")");
        Type("D1", "%18");
        Type("D2", "21.11.2026");
        Type("D3", "1.250 TL");
        Type("E1", "=TOPLA(B2:B3");      // hatali: reddedilmeli
        if (_editing) CancelEdit();
        Type("E2", "=E2+1");             // dongusel: reddedilmeli
        if (_editing) CancelEdit();
        Type("E3", "=B4*2");
        log.AppendLine($"E3 after cycle check = {SheetDisplay.Text(Sheet.Cell("E3"))}");
        UndoRedo(undo: true);
        log.AppendLine($"undo E3 => '{SheetDisplay.Text(Sheet.Cell("E3"))}'");
        UndoRedo(undo: false);
        log.AppendLine($"redo E3 => '{SheetDisplay.Text(Sheet.Cell("E3"))}'");
        View.SelectRange(new CellRange(2, 2, 3, 2));
        log.AppendLine("Durum: " + SumText.Text);
        foreach (var m in _silentLog) log.AppendLine("Uyarı: " + m);
        return log.ToString();
    }
}
