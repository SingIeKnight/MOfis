using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;

namespace MOfis.Sheet;

/// ClosedXML hesap motorunun kirilgan noktalarina karsi korumalar.
/// - Ayristirma hatali formul, henuz hesaplanmamis bir kitaba yazilirsa hucre bozuk kalir ("not in the chain").
///   Bu yuzden formul once ayri, bos bir deneme kitabinda sinanir.
/// - Dongusel basvuru motorun bagimlilik zincirini bozar; tek guvenilir kurtarma kitabi bellege kaydedip yeniden acmaktir.
public static class SheetEngine
{
    /// Gecersizse kullaniciya gosterilecek aciklama, gecerliyse null.
    public static string? Validate(string formulaEn)
    {
        try
        {
            using var wb = new XLWorkbook();
            var cell = wb.AddWorksheet("Deneme").Cell(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber);
            cell.FormulaA1 = formulaEn;
            _ = cell.Value;
            return null;
        }
        catch (ExpressionParseException ex)
        {
            return ex.Message.Contains("parameter", StringComparison.OrdinalIgnoreCase)
                ? "Fonksiyona verilen bağımsız değişken sayısı yanlış."
                : "Formülün yazımında bir hata var (eksik parantez, ayırıcı veya işleç olabilir).";
        }
        catch
        {
            // Deneme kitabinda olmayan sayfa vb.: gercek kitapta sorun cikarmaz.
            return null;
        }
    }

    public static bool IsCycle(Exception ex) =>
        ex is InvalidOperationException && (ex.Message.Contains("cycle", StringComparison.OrdinalIgnoreCase) ||
                                            ex.Message.Contains("not in the chain", StringComparison.OrdinalIgnoreCase));

    /// Hucreyi hesaplar; dongusel basvuru varsa true.
    public static bool HasCycle(IXLCell cell)
    {
        try
        {
            _ = cell.Value;
            return false;
        }
        catch (Exception ex) when (IsCycle(ex))
        {
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// Kitabi bellege kaydedip yeniden acar (hesap motoru sifirlanir). Eski kitap kapatilir.
    public static XLWorkbook Rebuild(XLWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.SaveAs(ms, new SaveOptions { EvaluateFormulasBeforeSaving = false, ValidatePackage = false });
        ms.Position = 0;
        var fresh = new XLWorkbook(ms);
        wb.Dispose();
        return fresh;
    }
}

/// Bir hucrenin icerik anlik goruntusu (geri al / yinele icin).
public sealed record CellSnap(int Row, int Col, string? Formula, XLCellValue Value, string NumberFormat, int NumberFormatId)
{
    public static CellSnap Take(IXLCell c)
    {
        var nf = c.Style.NumberFormat;
        return new CellSnap(c.Address.RowNumber, c.Address.ColumnNumber, c.HasFormula ? c.FormulaA1 : null,
            c.HasFormula ? Blank.Value : c.CachedOrValue(), nf.Format ?? "", nf.NumberFormatId);
    }

    public void Restore(IXLWorksheet ws)
    {
        var c = ws.Cell(Row, Col);
        if (Formula != null) c.FormulaA1 = Formula;
        else if (Value.IsBlank) c.Clear(XLClearOptions.Contents);
        else c.Value = Value;
        if (NumberFormat.Length > 0) c.Style.NumberFormat.Format = NumberFormat;
        else c.Style.NumberFormat.NumberFormatId = NumberFormatId;
    }
}

public sealed class UndoStep(string sheet, List<CellSnap> before, List<CellSnap> after)
{
    public string Sheet { get; } = sheet;
    public List<CellSnap> Before { get; } = before;
    public List<CellSnap> After { get; } = after;
}

public sealed class UndoStack
{
    private const int Max = 200;
    private readonly List<UndoStep> _undo = [];
    private readonly Stack<UndoStep> _redo = new();

    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    public void Push(UndoStep step)
    {
        _undo.Add(step);
        if (_undo.Count > Max) _undo.RemoveAt(0);
        _redo.Clear();
    }

    public UndoStep? Undo()
    {
        if (_undo.Count == 0) return null;
        var s = _undo[^1];
        _undo.RemoveAt(_undo.Count - 1);
        _redo.Push(s);
        return s;
    }

    public UndoStep? Redo()
    {
        if (_redo.Count == 0) return null;
        var s = _redo.Pop();
        _undo.Add(s);
        return s;
    }

    public void Clear()
    {
        _undo.Clear();
        _redo.Clear();
    }
}

internal static class CellExtensions
{
    /// Formulsuz hucrenin degeri (hesaplama tetiklemez).
    public static XLCellValue CachedOrValue(this IXLCell c)
    {
        try { return c.Value; } catch { return Blank.Value; }
    }
}
