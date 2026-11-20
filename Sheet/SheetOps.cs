using System.Text;
using ClosedXML.Excel;

namespace MOfis.Sheet;

/// Pencereden bagimsiz tablo islemleri (test edilebilsin diye ayri).
public static class SheetOps
{
    /// Aralik, sayfanin kullanilan alaniyla kesistirilir (tum sutun/satir seciminde milyonlarca hucre gezilmesin).
    public static CellRange? ClipToUsed(IXLWorksheet ws, CellRange r)
    {
        var used = ws.RangeUsed(XLCellsUsedOptions.All);
        if (used == null) return null;
        int r1 = Math.Max(r.R1, 1), c1 = Math.Max(r.C1, 1);
        int r2 = Math.Min(r.R2, used.LastRow().RowNumber()), c2 = Math.Min(r.C2, used.LastColumn().ColumnNumber());
        if (r1 > r2 || c1 > c2) return null;
        return new CellRange(r1, c1, r2, c2);
    }

    /// Secimin panoya gidecek metni: hucrelerin gorunen degerleri, sekme ve satir sonlariyla (Excel bicimi).
    public static string ToClipboardText(IXLWorksheet ws, CellRange r)
    {
        var sb = new StringBuilder();
        for (int row = r.R1; row <= r.R2; row++)
        {
            for (int col = r.C1; col <= r.C2; col++)
            {
                if (col > r.C1) sb.Append('\t');
                var t = SheetDisplay.Text(ws.Cell(row, col));
                if (t.IndexOfAny(['\t', '\n', '"']) >= 0) t = "\"" + t.Replace("\"", "\"\"") + "\"";
                sb.Append(t);
            }
            sb.Append("\r\n");
        }
        return sb.ToString();
    }

    /// Hucreleri (deger, formul, bicim) baska yere kopyalar; goreli formul basvurulari kaydirilir.
    public static void CopyCells(IXLWorksheet src, CellRange from, IXLWorksheet dst, int toRow, int toCol)
    {
        int rows = from.R2 - from.R1 + 1, cols = from.C2 - from.C1 + 1;
        // Kaynak ve hedef cakisiyorsa once kaynagi bellege al.
        var buffer = new List<(int Dr, int Dc, IXLCell Cell)>();
        using var scratch = new XLWorkbook();
        var tmp = scratch.AddWorksheet("Pano");
        for (int i = 0; i < rows; i++)
        for (int j = 0; j < cols; j++)
        {
            var s = src.Cell(from.R1 + i, from.C1 + j);
            var t = tmp.Cell(from.R1 + i, from.C1 + j); // ayni adres: formul kaymasin
            t.CopyFrom(s);
            buffer.Add((i, j, t));
        }
        foreach (var (dr, dc, cell) in buffer)
        {
            int r = toRow + dr, c = toCol + dc;
            if (r > XLHelper.MaxRowNumber || c > XLHelper.MaxColumnNumber) continue;
            dst.Cell(r, c).CopyFrom(cell);
        }
    }

    /// Disaridan (Excel, not defteri...) yapistirilan sekmeli metni hucrelere yazar.
    public static void PasteText(IXLWorksheet ws, List<List<string>> rows, int toRow, int toCol)
    {
        for (int i = 0; i < rows.Count; i++)
        for (int j = 0; j < rows[i].Count; j++)
        {
            int r = toRow + i, c = toCol + j;
            if (r > XLHelper.MaxRowNumber || c > XLHelper.MaxColumnNumber) continue;
            var p = CellInput.Parse(rows[i][j]);
            if (p.Kind == InputKind.Formula && SheetEngine.Validate(p.FormulaEn!) != null) p = new ParsedInput(InputKind.Text, rows[i][j]);
            CellInput.Apply(ws.Cell(r, c), p);
        }
    }

    /// Otomatik toplam icin: etkin hucrenin ustundeki (yoksa solundaki) bitisik sayi blogu.
    public static string? AutoSumFormula(IXLWorksheet ws, int row, int col)
    {
        static bool IsNum(IXLCell c)
        {
            try { return c.Value.IsNumber; } catch { return false; }
        }
        int r = row - 1;
        while (r >= 1 && ws.Cell(r, col).IsEmpty()) r--; // aradaki bosluklari atla (Excel gibi bir kez)
        if (r >= 1 && IsNum(ws.Cell(r, col)))
        {
            int end = r;
            while (r - 1 >= 1 && IsNum(ws.Cell(r - 1, col))) r--;
            var a = XLHelper.GetColumnLetterFromNumber(col);
            return $"=TOPLA({a}{r}:{a}{end})";
        }
        int c = col - 1;
        if (c >= 1 && IsNum(ws.Cell(row, c)))
        {
            int end = c;
            while (c - 1 >= 1 && IsNum(ws.Cell(row, c - 1))) c--;
            return $"=TOPLA({XLHelper.GetColumnLetterFromNumber(c)}{row}:{XLHelper.GetColumnLetterFromNumber(end)}{row})";
        }
        return null;
    }

    /// Durum cubugu: secimdeki sayilarin toplami/ortalamasi ve dolu hucre sayisi.
    public static (double Sum, int Numbers, int Filled) Stats(IXLWorksheet ws, CellRange r)
    {
        var clip = ClipToUsed(ws, r);
        if (clip is not { } c) return (0, 0, 0);
        double sum = 0;
        int nums = 0, filled = 0;
        foreach (var cell in ws.Range(c.R1, c.C1, c.R2, c.C2).CellsUsed(XLCellsUsedOptions.Contents))
        {
            filled++;
            try
            {
                var v = cell.Value;
                if (v.IsNumber) { sum += v.GetNumber(); nums++; }
            }
            catch
            {
                // hesaplanamayan hucre sayilmaz
            }
        }
        return (sum, nums, filled);
    }

    /// "B12", "A1:C5", "C:D", "3:7" -> aralik.
    public static CellRange? ParseAddress(string text)
    {
        var t = text.Trim().Replace("$", "").ToUpperInvariant();
        if (t.Length == 0) return null;
        var parts = t.Split(':');
        if (parts.Length > 2) return null;
        if (parts.Length == 2 && parts.All(p => p.Length > 0 && p.All(char.IsAsciiLetter)))
        {
            int a = XLHelper.GetColumnNumberFromLetter(parts[0]), b = XLHelper.GetColumnNumberFromLetter(parts[1]);
            if (a < 1 || b < 1 || a > XLHelper.MaxColumnNumber || b > XLHelper.MaxColumnNumber) return null;
            return CellRange.Of(1, a, XLHelper.MaxRowNumber, b);
        }
        if (parts.Length == 2 && parts.All(p => p.Length > 0 && p.All(char.IsAsciiDigit)))
        {
            if (!int.TryParse(parts[0], out var a) || !int.TryParse(parts[1], out var b) || a < 1 || b < 1 || a > XLHelper.MaxRowNumber || b > XLHelper.MaxRowNumber) return null;
            return CellRange.Of(a, 1, b, XLHelper.MaxColumnNumber);
        }
        var first = Cell(parts[0]);
        if (first == null) return null;
        var second = parts.Length == 2 ? Cell(parts[1]) : first;
        if (second == null) return null;
        return CellRange.Of(first.Value.R, first.Value.C, second.Value.R, second.Value.C);

        static (int R, int C)? Cell(string s)
        {
            int i = 0;
            while (i < s.Length && char.IsAsciiLetter(s[i])) i++;
            if (i is 0 or > 3 || i == s.Length) return null;
            if (!int.TryParse(s[i..], out var row) || row < 1 || row > XLHelper.MaxRowNumber) return null;
            int col = XLHelper.GetColumnNumberFromLetter(s[..i]);
            if (col < 1 || col > XLHelper.MaxColumnNumber) return null;
            return (row, col);
        }
    }
}
