using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MOfis.Common;
using MOfis.Sheet;
using MOfis.Writer;

namespace MOfis.Shell;

/// Baslangic ekrani: yeni belge, ac, son kullanilanlar ve kurtarilan belgeler.
public partial class StartWindow : Window
{
    public StartWindow()
    {
        InitializeComponent();
        VersionText.Text = "Sürüm " + AppInfo.Version;
        Greeting.Text = GreetingFor(DateTime.Now.Hour);
        BuildSheetThumbnail();
        InputBindings.Add(new KeyBinding(new RelayCommand(() => NewDocument_Click(this, new RoutedEventArgs())), Key.N, ModifierKeys.Control));
        InputBindings.Add(new KeyBinding(new RelayCommand(() => Open_Click(this, new RoutedEventArgs())), Key.O, ModifierKeys.Control));
        Activated += (_, _) => RefreshLists();
    }

    private static string GreetingFor(int hour) => hour switch
    {
        >= 5 and < 12 => "Günaydın",
        >= 12 and < 18 => "İyi günler",
        >= 18 and < 23 => "İyi akşamlar",
        _ => "İyi geceler",
    };

    private void NewDocument_Click(object sender, RoutedEventArgs e)
    {
        AppActions.NewDocument();
        Close();
    }

    private void NewSheet_Click(object sender, RoutedEventArgs e)
    {
        AppActions.NewSheet();
        Close();
    }

    private void Open_Click(object sender, RoutedEventArgs e)
    {
        var before = Application.Current.Windows.Count;
        AppActions.ShowOpenDialog(this);
        if (Application.Current.Windows.Count > before) Close();
    }

    private void RefreshLists()
    {
        RecentList.Children.Clear();
        var recent = RecentFiles.Load();
        RecentEmpty.Visibility = recent.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (var r in recent) RecentList.Children.Add(MakeRecentRow(r));

        RecoveryList.Children.Clear();
        var recovered = AutoRecovery.List();
        RecoveryPanel.Visibility = recovered.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (var r in recovered) RecoveryList.Children.Add(MakeRecoveryRow(r));
    }

    private Button MakeRecentRow(RecentFile r)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var icon = new TextBlock
        {
            Text = AppActions.IsSheetFile(r.Path) ? "" : "",
            FontFamily = (FontFamily)FindResource("IconFont"),
            FontSize = 20,
            Foreground = (Brush)FindResource("AccentBrush"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        var names = new StackPanel();
        names.Children.Add(new TextBlock { Text = r.FileName, FontWeight = FontWeights.SemiBold });
        names.Children.Add(new TextBlock { Text = r.Folder, Foreground = (Brush)FindResource("SubtleTextBrush"), FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis });
        var date = new TextBlock
        {
            Text = r.LastOpened.ToString("d MMMM yyyy HH:mm"),
            Foreground = (Brush)FindResource("SubtleTextBrush"),
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(16, 0, 0, 0),
        };
        Grid.SetColumn(names, 1);
        Grid.SetColumn(date, 2);
        grid.Children.Add(icon);
        grid.Children.Add(names);
        grid.Children.Add(date);

        var btn = new Button { Content = grid, Style = (Style)FindResource("TileButton"), Padding = new Thickness(10, 7, 10, 7), ToolTip = r.Path };
        if (!r.Exists) btn.Opacity = 0.5;
        btn.Click += (_, _) =>
        {
            if (AppActions.OpenFile(r.Path, this)) Close();
            else RefreshLists();
        };

        var menu = new ContextMenu();
        var remove = new MenuItem { Header = "Listeden kaldır" };
        remove.Click += (_, _) => { RecentFiles.Remove(r.Path); RefreshLists(); };
        menu.Items.Add(remove);
        btn.ContextMenu = menu;
        return btn;
    }

    private UIElement MakeRecoveryRow(RecoveryEntry r)
    {
        var panel = new DockPanel { Margin = new Thickness(0, 2, 0, 2) };
        var discard = new Button { Content = "Sil", Style = (Style)FindResource("PlainButton"), Padding = new Thickness(10, 3, 10, 3), Margin = new Thickness(8, 0, 0, 0) };
        var open = new Button { Content = "Aç", Style = (Style)FindResource("AccentButton"), Padding = new Thickness(12, 3, 12, 3) };
        DockPanel.SetDock(discard, Dock.Right);
        DockPanel.SetDock(open, Dock.Right);
        panel.Children.Add(discard);
        panel.Children.Add(open);
        panel.Children.Add(new TextBlock
        {
            Text = $"{r.Title}   ·   {r.SavedAt:d MMMM HH:mm}",
            VerticalAlignment = VerticalAlignment.Center,
        });

        open.Click += (_, _) =>
        {
            Window win;
            bool ok;
            if (r.IsSheet)
            {
                var sw = new SheetWindow();
                ok = sw.LoadRecovery(r);
                win = sw;
            }
            else
            {
                var ww = new WriterWindow();
                ok = ww.LoadRecovery(r);
                win = ww;
            }
            if (ok)
            {
                win.Show();
                Close();
            }
            else win.Close();
        };
        discard.Click += (_, _) =>
        {
            if (MessageBox.Show($"\"{r.Title}\" kurtarma kopyası silinsin mi?", AppInfo.Name, MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                AutoRecovery.Delete(r.Id);
                RefreshLists();
            }
        };
        return panel;
    }

    private void BuildSheetThumbnail()
    {
        var line = new SolidColorBrush(Color.FromRgb(0xE3, 0xE4, 0xEA));
        for (int c = 0; c < 4; c++) SheetThumb.ColumnDefinitions.Add(new ColumnDefinition());
        for (int r = 0; r < 9; r++) SheetThumb.RowDefinitions.Add(new RowDefinition());
        for (int r = 0; r < 9; r++)
        {
            for (int c = 0; c < 4; c++)
            {
                var cell = new Border { BorderBrush = line, BorderThickness = new Thickness(0, 0, 1, 1) };
                if (r == 0) cell.Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE7, 0xFF));
                Grid.SetRow(cell, r);
                Grid.SetColumn(cell, c);
                SheetThumb.Children.Add(cell);
            }
        }
    }
}

/// Kod icinde kisayol tusu baglamak icin en basit ICommand.
public sealed class RelayCommand(Action run) : ICommand
{
    public event EventHandler? CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => run();
}
