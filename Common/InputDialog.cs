using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MOfis.Common;

/// Basit, kod ile kurulan giris kutusu (baglanti adresi gibi tek satirlik sorular icin).
public static class InputDialog
{
    public static string? Ask(Window owner, string title, string prompt, string initial = "")
    {
        var win = new Window
        {
            Title = title,
            Owner = owner,
            Width = 420,
            SizeToContent = SizeToContent.Height,
            ResizeMode = ResizeMode.NoResize,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false,
            FontFamily = (System.Windows.Media.FontFamily)Application.Current.Resources["UiFont"],
            FontSize = 13,
        };

        var box = new TextBox { Text = initial, Margin = new Thickness(0, 6, 0, 14), Padding = new Thickness(4, 3, 4, 3) };
        var ok = new Button { Content = "Tamam", IsDefault = true, Style = (Style)Application.Current.Resources["AccentButton"], MinWidth = 80 };
        var cancel = new Button { Content = "İptal", IsCancel = true, Style = (Style)Application.Current.Resources["PlainButton"], MinWidth = 80, Margin = new Thickness(8, 0, 0, 0) };
        ok.Click += (_, _) => win.DialogResult = true;

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);

        var panel = new StackPanel { Margin = new Thickness(16) };
        panel.Children.Add(new TextBlock { Text = prompt, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(box);
        panel.Children.Add(buttons);
        win.Content = panel;
        win.Loaded += (_, _) => { box.Focus(); box.SelectAll(); Keyboard.Focus(box); };

        return win.ShowDialog() == true ? box.Text : null;
    }
}
