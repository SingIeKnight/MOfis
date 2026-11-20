using System.Globalization;
using System.Windows;
using System.Windows.Markup;
using MOfis.Common;
using MOfis.Shell;

namespace MOfis;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Arayuz Turkce: tarih/sayi bicimleri ve yazim denetimi dili tr-TR.
        var tr = new CultureInfo("tr-TR");
        CultureInfo.DefaultThreadCurrentCulture = tr;
        CultureInfo.DefaultThreadCurrentUICulture = tr;
        Thread.CurrentThread.CurrentCulture = tr;
        Thread.CurrentThread.CurrentUICulture = tr;
        FrameworkElement.LanguageProperty.OverrideMetadata(typeof(FrameworkElement),
            new FrameworkPropertyMetadata(XmlLanguage.GetLanguage("tr-TR")));

        DispatcherUnhandledException += (_, ex) =>
        {
            ErrorLog.Write(ex.Exception);
            MessageBox.Show("Beklenmeyen bir hata oluştu:\n\n" + ex.Exception.Message +
                            "\n\nAyrıntılar şu dosyaya yazıldı:\n" + ErrorLog.FilePath, AppInfo.Name,
                MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };

        // Gelistirme araci: --shot <start|writer> <cikti.png> [dosya]  -> pencerenin resmini cekip kapanir.
        if (e.Args.Length >= 3 && e.Args[0] == "--shot")
        {
            Screenshot.Run(e.Args[1], e.Args[2], e.Args.Length > 3 ? e.Args[3] : null);
            return;
        }

        if (e.Args.Length > 0 && File.Exists(e.Args[0]))
        {
            AppActions.OpenFile(e.Args[0], null);
            if (Windows.Count > 0) return;
        }
        new StartWindow().Show();
    }
}
