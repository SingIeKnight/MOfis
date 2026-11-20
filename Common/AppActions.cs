using System.Windows;
using Microsoft.Win32;
using MOfis.Writer;

namespace MOfis.Common;

/// Pencerelerin ortak kullandigi "yeni / ac" islemleri.
public static class AppActions
{
    public const string OpenFilter =
        "Tüm desteklenen dosyalar|*.docx;*.rtf;*.txt;*.xlsx;*.csv|" +
        "Word belgesi (*.docx)|*.docx|Zengin metin (*.rtf)|*.rtf|Düz metin (*.txt)|*.txt|" +
        "Excel çalışma kitabı (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv";

    public static void NewDocument()
    {
        new WriterWindow().Show();
    }

    public static void ShowOpenDialog(Window? owner)
    {
        var dlg = new OpenFileDialog { Filter = OpenFilter, Title = "Aç" };
        if (dlg.ShowDialog(owner) == true) OpenFile(dlg.FileName, owner);
    }

    public static bool IsWriterFile(string path) =>
        Path.GetExtension(path).ToLowerInvariant() is ".docx" or ".rtf" or ".txt";

    public static bool IsSheetFile(string path) =>
        Path.GetExtension(path).ToLowerInvariant() is ".xlsx" or ".csv";

    /// Dosyayi uygun pencerede acar. Basarili olursa true.
    public static bool OpenFile(string path, Window? owner)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show($"Dosya bulunamadı:\n{path}", AppInfo.Name, MessageBoxButton.OK, MessageBoxImage.Warning);
            RecentFiles.Remove(path);
            return false;
        }

        // Ayni dosya zaten aciksa o pencereyi one getir.
        foreach (Window w in Application.Current.Windows)
        {
            if (w is IDocumentWindow dw && string.Equals(dw.FilePath, path, StringComparison.OrdinalIgnoreCase))
            {
                if (w.WindowState == WindowState.Minimized) w.WindowState = WindowState.Normal;
                w.Activate();
                return true;
            }
        }

        if (IsWriterFile(path))
        {
            var win = new WriterWindow();
            if (!win.LoadFile(path))
            {
                win.Close();
                return false;
            }
            win.Show();
            return true;
        }

        if (IsSheetFile(path))
        {
            MessageBox.Show("Hesap tablosu bir sonraki sürümde geliyor.", AppInfo.Name, MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        MessageBox.Show("Bu dosya türü desteklenmiyor.", AppInfo.Name, MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
    }
}

/// Dosya acan pencerelerin ortak arayuzu (ayni dosya iki kez acilmasin diye).
public interface IDocumentWindow
{
    string? FilePath { get; }
}
