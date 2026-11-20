using System.Diagnostics;
using System.Printing;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using MOfis.Common;

namespace MOfis.Writer;

public sealed record LoadedDocument(FlowDocument Document, PageSettings Page, IReadOnlyList<string> Warnings);

/// Belgeyi uzantisina gore (docx / rtf / txt) acar ve kaydeder.
public static class DocumentIO
{
    static DocumentIO()
    {
        // Eski Turkce .txt dosyalari (Windows-1254) icin.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public const string SaveFilter =
        "Word belgesi (*.docx)|*.docx|Zengin metin (*.rtf)|*.rtf|Düz metin (*.txt)|*.txt";

    public static LoadedDocument Load(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        // Word'de acik dosyalar da okunabilsin diye paylasimli acilir.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        switch (ext)
        {
            case ".docx":
            {
                using var ms = new MemoryStream();
                fs.CopyTo(ms);
                ms.Position = 0;
                var r = DocxImporter.Load(ms);
                return new LoadedDocument(r.Document, r.Page, r.Warnings);
            }
            case ".rtf":
            {
                var doc = DocHelpers.CreateBlank();
                new TextRange(doc.ContentStart, doc.ContentEnd).Load(fs, DataFormats.Rtf);
                return new LoadedDocument(doc, new PageSettings(), []);
            }
            default:
            {
                using var ms = new MemoryStream();
                fs.CopyTo(ms);
                var doc = DocHelpers.CreateBlank();
                doc.Blocks.Clear();
                var text = DecodeText(ms.ToArray()).Replace("\r\n", "\n").Replace('\r', '\n');
                foreach (var line in text.Split('\n')) doc.Blocks.Add(new Paragraph(new Run(line)));
                if (doc.Blocks.Count == 0) doc.Blocks.Add(new Paragraph());
                return new LoadedDocument(doc, new PageSettings(), []);
            }
        }
    }

    /// UTF-8 (BOM'lu/BOM'suz) veya UTF-16 dener; gecersizse Turkce Windows kod sayfasi (1254).
    private static string DecodeText(byte[] bytes)
    {
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        int start = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        try
        {
            return new UTF8Encoding(false, true).GetString(bytes, start, bytes.Length - start);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.GetEncoding(1254).GetString(bytes);
        }
    }

    /// Once gecici dosyaya yazar, sonra degistirir: kayit yarida kalirsa eski dosya bozulmaz.
    public static void Save(FlowDocument doc, PageSettings page, string path)
    {
        var tmp = path + ".mofis-tmp";
        try
        {
            using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                switch (Path.GetExtension(path).ToLowerInvariant())
                {
                    case ".rtf":
                        new TextRange(doc.ContentStart, doc.ContentEnd).Save(fs, DataFormats.Rtf);
                        break;
                    case ".txt":
                    {
                        var text = new TextRange(doc.ContentStart, doc.ContentEnd).Text;
                        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(text)).ToArray();
                        fs.Write(bytes);
                        break;
                    }
                    default:
                        DocxExporter.Save(doc, page, fs);
                        break;
                }
            }
            File.Move(tmp, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(tmp)) File.Delete(tmp);
        }
    }

    /// Canli belgenin bagimsiz bir kopyasi (yazdirma icin); DOCX uzerinden gidip gelir.
    public static FlowDocument CloneViaDocx(FlowDocument doc, PageSettings page)
    {
        using var ms = new MemoryStream();
        DocxExporter.Save(doc, page, ms);
        ms.Position = 0;
        return DocxImporter.Load(ms).Document;
    }
}

public sealed class RecoveryEntry
{
    public string Id { get; set; } = "";
    public string Title { get; set; } = "";
    public string? OriginalPath { get; set; }
    public DateTime SavedAt { get; set; }
    public int ProcessId { get; set; }

    public string DocPath => Path.Combine(AppInfo.RecoveryDir, Id + ".docx");
}

/// Kaydedilmemis degisikliklerin birkac dakikada bir alinan kopyalari. Pencere duzgun kapaninca silinir;
/// program cokerse baslangic ekraninda "Kurtarilan belgeler" altinda gorunur.
public static class AutoRecovery
{
    public static void Save(RecoveryEntry entry, FlowDocument doc, PageSettings page)
    {
        try
        {
            entry.SavedAt = DateTime.Now;
            entry.ProcessId = Environment.ProcessId;
            DocumentIO.Save(doc, page, entry.DocPath);
            File.WriteAllText(Path.Combine(AppInfo.RecoveryDir, entry.Id + ".json"), JsonSerializer.Serialize(entry));
        }
        catch
        {
            // Kurtarma kopyasi alinamazsa kullaniciyi rahatsiz etme; bir sonraki denemede tekrar denenir.
        }
    }

    public static List<RecoveryEntry> List()
    {
        var result = new List<RecoveryEntry>();
        foreach (var json in Directory.EnumerateFiles(AppInfo.RecoveryDir, "*.json"))
        {
            try
            {
                var e = JsonSerializer.Deserialize<RecoveryEntry>(File.ReadAllText(json));
                if (e == null || !File.Exists(e.DocPath)) continue;
                if (IsRunning(e.ProcessId)) continue; // hala acik bir pencereye ait
                result.Add(e);
            }
            catch
            {
                // bozuk kayit: atla
            }
        }
        return result.OrderByDescending(e => e.SavedAt).ToList();
    }

    public static void Delete(string id)
    {
        foreach (var ext in new[] { ".docx", ".json" })
        {
            try { File.Delete(Path.Combine(AppInfo.RecoveryDir, id + ext)); } catch { }
        }
    }

    private static bool IsRunning(int pid)
    {
        if (pid == Environment.ProcessId) return true;
        try
        {
            using var p = Process.GetProcessById(pid);
            return string.Equals(p.ProcessName, Process.GetCurrentProcess().ProcessName, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }
}

/// Yazdirma ve PDF'e aktarma (Windows'un "Microsoft Print to PDF" yazicisi ile).
public static class Printing
{
    public const string PdfPrinter = "Microsoft Print to PDF";

    public static void Print(FlowDocument source, PageSettings page, string title, bool toPdf)
    {
        var dlg = new PrintDialog();
        if (toPdf)
        {
            try
            {
                dlg.PrintQueue = new LocalPrintServer().GetPrintQueue(PdfPrinter);
            }
            catch
            {
                MessageBox.Show($"\"{PdfPrinter}\" yazıcısı bulunamadı.\n\nWindows Ayarlar > Uygulamalar > İsteğe bağlı özellikler bölümünden etkinleştirebilirsiniz.",
                    AppInfo.Name, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else if (dlg.ShowDialog() != true)
        {
            return;
        }

        if (dlg.PrintTicket != null)
            dlg.PrintTicket.PageOrientation = page.IsLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;

        var copy = DocumentIO.CloneViaDocx(source, page);
        copy.PageWidth = page.Width;
        copy.PageHeight = page.Height;
        copy.PagePadding = new Thickness(page.MarginLeft, page.MarginTop, page.MarginRight, page.MarginBottom);
        copy.ColumnWidth = double.PositiveInfinity;
        copy.ColumnGap = 0;
        var paginator = ((IDocumentPaginatorSource)copy).DocumentPaginator;
        paginator.PageSize = new Size(page.Width, page.Height);
        dlg.PrintDocument(paginator, title);
    }
}
