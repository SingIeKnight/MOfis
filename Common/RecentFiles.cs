using System.Text.Json;

namespace MOfis.Common;

public sealed class RecentFile
{
    public string Path { get; set; } = "";
    public DateTime LastOpened { get; set; }

    public string FileName => System.IO.Path.GetFileName(Path);
    public string Folder => System.IO.Path.GetDirectoryName(Path) ?? "";
    public bool Exists => File.Exists(Path);
}

/// Son acilan dosyalar listesi (%AppData%\MOfis\son-dosyalar.json), en fazla 20 kayit.
public static class RecentFiles
{
    private const int Max = 20;
    private static string StorePath => System.IO.Path.Combine(AppInfo.DataDir, "son-dosyalar.json");

    public static List<RecentFile> Load()
    {
        try
        {
            if (!File.Exists(StorePath)) return [];
            return JsonSerializer.Deserialize<List<RecentFile>>(File.ReadAllText(StorePath)) ?? [];
        }
        catch
        {
            return [];
        }
    }

    public static void Add(string path)
    {
        var list = Load();
        list.RemoveAll(r => string.Equals(r.Path, path, StringComparison.OrdinalIgnoreCase));
        list.Insert(0, new RecentFile { Path = path, LastOpened = DateTime.Now });
        Save(list.Take(Max).ToList());
    }

    public static void Remove(string path)
    {
        var list = Load();
        list.RemoveAll(r => string.Equals(r.Path, path, StringComparison.OrdinalIgnoreCase));
        Save(list);
    }

    private static void Save(List<RecentFile> list)
    {
        try
        {
            File.WriteAllText(StorePath, JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch
        {
            // Liste kaydedilemezse uygulama yine calismaya devam eder.
        }
    }
}
