using System.Reflection;

namespace MOfis.Common;

public static class AppInfo
{
    public const string Name = "MOfis";
    public const string WriterName = "MOfis Yazı";
    public const string SheetName = "MOfis Hesap";

    public static string Version =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.0";

    /// Ayarlar, son dosyalar: %AppData%\MOfis
    public static string DataDir
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Name);
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// Otomatik kurtarma kopyalari: %LocalAppData%\MOfis\Kurtarma
    public static string RecoveryDir
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Name, "Kurtarma");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }
}

public static class Units
{
    public static double CmToPx(double cm) => cm / 2.54 * 96.0;
    public static double PxToCm(double px) => px / 96.0 * 2.54;
    public static double PtToPx(double pt) => pt * 96.0 / 72.0;
    public static double PxToPt(double px) => px * 72.0 / 96.0;
    public static double TwipsToPx(double twips) => twips / 15.0;
    public static long PxToTwips(double px) => (long)Math.Round(px * 15.0);
    public const long EmuPerPx = 9525;
}
