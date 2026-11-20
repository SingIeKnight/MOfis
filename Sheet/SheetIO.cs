using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace MOfis.Sheet;

public sealed record LoadedSheet(XLWorkbook Workbook, IReadOnlyList<string> Warnings);

/// Hesap tablosunu uzantisina gore (xlsx / csv) acar ve kaydeder.
public static class SheetIO
{
    private static readonly CultureInfo Tr = new("tr-TR");

    static SheetIO()
    {
        // Eski Turkce CSV dosyalari (Windows-1254) icin.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public const string SaveFilter = "Excel çalışma kitabı (*.xlsx)|*.xlsx|CSV - noktalı virgülle ayrılmış (*.csv)|*.csv";

    public static bool IsCsv(string path) => Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase);

    public static LoadedSheet Load(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var ms = new MemoryStream();
        fs.CopyTo(ms);
        ms.Position = 0;

        if (IsCsv(path))
        {
            var wb = new XLWorkbook();
            var name = Path.GetFileNameWithoutExtension(path);
            var ws = wb.AddWorksheet(SafeSheetName(name));
            ReadCsv(ws, DecodeText(ms.ToArray()));
            return new LoadedSheet(wb, []);
        }

        var warnings = FindUnsupported(ms);
        ms.Position = 0;
        return new LoadedSheet(new XLWorkbook(ms), warnings);
    }

    /// ClosedXML'in kayitta koruyamadigi ogeler: bunlar varsa kullaniciya soylenir.
    private static List<string> FindUnsupported(Stream s)
    {
        var w = new List<string>();
        try
        {
            using var doc = SpreadsheetDocument.Open(s, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return w;
            bool charts = false, shapes = false, pivots = false;
            foreach (var wsPart in wbPart.WorksheetParts)
            {
                if (wsPart.DrawingsPart is { } dp)
                {
                    if (dp.ChartParts.Any()) charts = true;
                    var xml = dp.WorksheetDrawing?.OuterXml ?? "";
                    if (xml.Contains(":sp>", StringComparison.Ordinal) || xml.Contains(":sp ", StringComparison.Ordinal)) shapes = true;
                }
                if (wsPart.PivotTableParts.Any()) pivots = true;
            }
            if (wbPart.ChartsheetParts.Any()) charts = true;
            if (charts) w.Add("grafikler");
            if (shapes) w.Add("şekiller ve metin kutuları");
            if (pivots) w.Add("özet tablolar");
            if (wbPart.VbaProjectPart != null) w.Add("makrolar");
        }
        catch
        {
            // Inceleme yapilamazsa ClosedXML yine de acmayi dener.
        }
        return w;
    }

    public static void Save(XLWorkbook wb, string path, int activeSheetIndex)
    {
        var tmp = path + ".mofis-tmp";
        try
        {
            using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                if (IsCsv(path))
                {
                    var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(WriteCsv(wb.Worksheet(activeSheetIndex + 1)))).ToArray();
                    fs.Write(bytes);
                }
                else
                {
                    try
                    {
                        // Formul sonuclari dosyaya da yazilsin: baska programlar hesaplamadan gosterebilsin.
                        wb.SaveAs(fs, new SaveOptions { EvaluateFormulasBeforeSaving = true, ValidatePackage = false });
                    }
                    catch
                    {
                        // Motorun hesaplayamadigi bir formul kaydi engellemesin.
                        fs.SetLength(0);
                        wb.SaveAs(fs, new SaveOptions { EvaluateFormulasBeforeSaving = false, ValidatePackage = false });
                    }
                }
            }
            File.Move(tmp, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(tmp)) File.Delete(tmp);
        }
    }

    public static string SafeSheetName(string name)
    {
        var sb = new StringBuilder();
        foreach (var c in name) sb.Append(c is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? '_' : c);
        var s = sb.ToString().Trim('\'').Trim();
        if (s.Length == 0) s = "Sayfa1";
        return s.Length > 31 ? s[..31] : s;
    }

    // ================================================================ CSV

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

    /// Ilk satirdaki (tirnak disi) ayiricilari sayarak ';', ',' veya sekme secer.
    public static char DetectDelimiter(string text)
    {
        int semi = 0, comma = 0, tab = 0;
        bool q = false;
        foreach (var c in text)
        {
            if (c == '"') q = !q;
            else if (!q && (c == '\n' || c == '\r')) break;
            else if (!q && c == ';') semi++;
            else if (!q && c == ',') comma++;
            else if (!q && c == '\t') tab++;
        }
        if (tab > semi && tab > comma) return '\t';
        return comma > semi ? ',' : ';';
    }

    public static List<List<string>> ParseDelimited(string text, char delim)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool q = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (q)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else q = false;
                }
                else field.Append(c);
                continue;
            }
            if (c == '"' && field.Length == 0) q = true;
            else if (c == delim) { row.Add(field.ToString()); field.Clear(); }
            else if (c is '\r' or '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                row.Add(field.ToString());
                field.Clear();
                rows.Add(row);
                row = [];
            }
            else field.Append(c);
        }
        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            rows.Add(row);
        }
        return rows;
    }

    private static void ReadCsv(IXLWorksheet ws, string text)
    {
        char delim = DetectDelimiter(text);
        // Virgulle ayrilmis dosya Ingilizce yazimdir: 1.5 = bir bucuk.
        bool invariant = delim == ',';
        var rows = ParseDelimited(text, delim);
        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < rows[r].Count; c++)
            {
                var s = rows[r][c];
                if (s.Length == 0) continue;
                var p = CellInput.Parse(s, invariant);
                // CSV'deki "=..." metnini formul olarak calistirma (guvenlik: CSV ile formul enjeksiyonu).
                if (p.Kind == InputKind.Formula) p = new ParsedInput(InputKind.Text, s);
                CellInput.Apply(ws.Cell(r + 1, c + 1), p);
            }
        }
    }

    private static string WriteCsv(IXLWorksheet ws)
    {
        var sb = new StringBuilder();
        var used = ws.RangeUsed();
        if (used == null) return "";
        int lastRow = used.LastRow().RowNumber(), lastCol = used.LastColumn().ColumnNumber();
        for (int r = 1; r <= lastRow; r++)
        {
            for (int c = 1; c <= lastCol; c++)
            {
                if (c > 1) sb.Append(';');
                var cell = ws.Cell(r, c);
                string s;
                var v = cell.Value;
                if (v.IsNumber && !v.IsDateTime) s = v.GetNumber().ToString("0.###############", Tr);
                else s = SheetDisplay.Text(cell);
                if (s.IndexOfAny([';', '"', '\n', '\r']) >= 0) s = "\"" + s.Replace("\"", "\"\"") + "\"";
                sb.Append(s);
            }
            sb.Append("\r\n");
        }
        return sb.ToString();
    }

    /// Panodaki sekmeyle ayrilmis metin (Excel'den kopyalanan) -> satir/sutun dizisi.
    public static List<List<string>> ParseClipboard(string text)
    {
        var rows = ParseDelimited(text, '\t');
        // Excel kopyalamasi sonda bos bir satir birakir.
        if (rows.Count > 0 && rows[^1].Count == 1 && rows[^1][0].Length == 0) rows.RemoveAt(rows.Count - 1);
        return rows;
    }
}
