namespace MOfis.Common;

/// Beklenmeyen hatalarin tam ayrintisi: %LocalAppData%\MOfis\hatalar.log (en fazla ~1 MB, sonra bastan baslar).
public static class ErrorLog
{
    public static string FilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppInfo.Name, "hatalar.log");

    public static void Write(Exception ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            if (File.Exists(FilePath) && new FileInfo(FilePath).Length > 1_000_000) File.Delete(FilePath);
            File.AppendAllText(FilePath, $"===== {DateTime.Now:yyyy-MM-dd HH:mm:ss} sürüm {AppInfo.Version}\n{ex}\n\n");
        }
        catch
        {
            // gunluk yazilamazsa yapacak bir sey yok
        }
    }
}
