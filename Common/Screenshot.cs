using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MOfis.Shell;
using MOfis.Sheet;
using MOfis.Writer;

namespace MOfis.Common;

/// Gelistirme araci: pencereyi acip icerigini PNG olarak kaydeder ve uygulamayi kapatir.
/// Kullanim: MOfis.exe --shot start|writer|sheet cikti.png [acilacak-dosya]
public static class Screenshot
{
    public static void Run(string kind, string output, string? file)
    {
        RecentFiles.Disabled = true;
        Window win;
        if (kind == "writer")
        {
            var w = new WriterWindow();
            if (file != null) w.LoadFile(file);
            win = w;
        }
        else if (kind == "sheet")
        {
            var s = new SheetWindow();
            if (file != null) s.LoadFile(file);
            win = s;
        }
        else
        {
            win = new StartWindow();
        }

        win.WindowStartupLocation = WindowStartupLocation.Manual;
        win.Left = 0;
        win.Top = 0;
        win.Width = 1280;
        win.Height = 820;
        win.ContentRendered += async (_, _) =>
        {
            await Task.Delay(1500);
            try
            {
                if (win is WriterWindow ww)
                {
                    File.WriteAllText(Path.ChangeExtension(output, ".txt"), ww.DescribeLayout());
                    if (Environment.GetEnvironmentVariable("MOFIS_RULERTEST") == "1")
                        File.WriteAllText(Path.ChangeExtension(output, ".ruler.txt"), ww.SimulateRulerDrags());
                }
                if (win is SheetWindow sw)
                {
                    var txt = new System.Text.StringBuilder();
                    if (Environment.GetEnvironmentVariable("MOFIS_SHEETTEST") == "1") txt.AppendLine(sw.SimulateTyping());
                    txt.Append(sw.DescribeLayout());
                    File.WriteAllText(Path.ChangeExtension(output, ".txt"), txt.ToString());
                    await Task.Delay(300);
                }
                Save((FrameworkElement)win.Content, output);
            }
            finally
            {
                Application.Current.Shutdown();
            }
        };
        win.Show();
    }

    private static void Save(FrameworkElement el, string path)
    {
        var bmp = new RenderTargetBitmap((int)el.ActualWidth, (int)el.ActualHeight, 96, 96, PixelFormats.Pbgra32);
        bmp.Render(el);
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(bmp));
        using var fs = File.Create(path);
        enc.Save(fs);
    }
}
