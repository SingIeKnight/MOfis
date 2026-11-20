using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace MOfis.Sheet;

public enum InputKind { Empty, Text, Number, Boolean, Date, Formula }

/// Hucreye yazilan metnin ne oldugu: formul (Ingilizce), deger ve varsa onerilen sayi bicimi.
public sealed record ParsedInput(InputKind Kind, XLCellValue Value, string? FormulaEn = null, string? NumberFormat = null);

/// Turkce yazimla girilen hucre icerigini ayristirir (1,5 / 1.234,5 / %15 / 21.11.2026 / 250 TL / DOGRU / =TOPLA(...)).
public static partial class CellInput
{
    private static readonly CultureInfo Tr = new("tr-TR");

    public const string DateFormat = "dd.mm.yyyy";
    public const string DateTimeFormat = "dd.mm.yyyy hh:mm";
    public const string TimeFormat = "hh:mm";
    public const string CurrencyFormat = "#,##0.00 \"₺\"";

    [GeneratedRegex(@"^[+-]?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?([eE][+-]?\d+)?$")]
    private static partial Regex TrNumber();

    [GeneratedRegex(@"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?([eE][+-]?\d+)?$")]
    private static partial Regex EnNumber();

    [GeneratedRegex(@"^(\d{1,2})[./](\d{1,2})(?:[./](\d{2}|\d{4}))?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$")]
    private static partial Regex TrDate();

    [GeneratedRegex(@"^(\d{1,2}):(\d{2})(?::(\d{2}))?$")]
    private static partial Regex TrTime();

    /// invariantNumbers: Ingilizce yazimli CSV (virgulle ayrilmis) iceri alinirken 1.5 = bir bucuk.
    public static ParsedInput Parse(string? raw, bool invariantNumbers = false)
    {
        if (string.IsNullOrEmpty(raw)) return new ParsedInput(InputKind.Empty, Blank.Value);
        var text = raw.Trim();
        if (text.Length == 0) return new ParsedInput(InputKind.Text, raw);

        // Bastaki kesme isareti: "bunu metin olarak sakla" (Excel'deki gibi).
        if (raw.StartsWith('\'')) return new ParsedInput(InputKind.Text, raw[1..]);

        if (text.Length > 1 && text[0] == '=')
            return new ParsedInput(InputKind.Formula, Blank.Value, FormulaLocale.ToEnglish(text[1..]));

        var upper = text.ToUpper(Tr);
        if (upper is "DOĞRU" or "TRUE") return new ParsedInput(InputKind.Boolean, true);
        if (upper is "YANLIŞ" or "FALSE") return new ParsedInput(InputKind.Boolean, false);

        // Yuzde: %15 veya 15%
        if (text.Length > 1 && (text[0] == '%' || text[^1] == '%'))
        {
            var inner = (text[0] == '%' ? text[1..] : text[..^1]).Trim();
            if (TryNumber(inner, invariantNumbers, out var p))
                return new ParsedInput(InputKind.Number, p / 100.0, NumberFormat: inner.Contains(invariantNumbers ? '.' : ',') ? "0.00%" : "0%");
        }

        // Para: 250 TL, 250₺, ₺250
        var money = text;
        bool isMoney = false;
        if (money.EndsWith("TL", StringComparison.OrdinalIgnoreCase)) { money = money[..^2].Trim(); isMoney = true; }
        else if (money.EndsWith('₺')) { money = money[..^1].Trim(); isMoney = true; }
        else if (money.StartsWith('₺')) { money = money[1..].Trim(); isMoney = true; }
        if (isMoney && TryNumber(money, invariantNumbers, out var m))
            return new ParsedInput(InputKind.Number, m, NumberFormat: CurrencyFormat);

        if (TryNumber(text, invariantNumbers, out var n))
            return new ParsedInput(InputKind.Number, n);

        if (!invariantNumbers && TryDate(text, out var dt, out var fmt))
            return new ParsedInput(InputKind.Date, dt, NumberFormat: fmt);

        return new ParsedInput(InputKind.Text, raw);
    }

    public static bool TryNumber(string s, bool invariant, out double value)
    {
        value = 0;
        if (invariant)
        {
            if (!EnNumber().IsMatch(s)) return false;
            return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }
        if (!TrNumber().IsMatch(s)) return false;
        return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, Tr, out value);
    }

    private static bool TryDate(string s, out DateTime value, out string format)
    {
        value = default;
        format = DateFormat;
        var t = TrTime().Match(s);
        if (t.Success)
        {
            int h = int.Parse(t.Groups[1].Value, CultureInfo.InvariantCulture);
            int mi = int.Parse(t.Groups[2].Value, CultureInfo.InvariantCulture);
            int se = t.Groups[3].Success ? int.Parse(t.Groups[3].Value, CultureInfo.InvariantCulture) : 0;
            if (h > 23 || mi > 59 || se > 59) return false;
            // Excel'de saat, 30.12.1899 tarihine eklenmis kesirli gun olarak saklanir.
            value = new DateTime(1899, 12, 30, h, mi, se);
            format = t.Groups[3].Success ? "hh:mm:ss" : TimeFormat;
            return true;
        }

        var d = TrDate().Match(s);
        if (!d.Success) return false;
        int day = int.Parse(d.Groups[1].Value, CultureInfo.InvariantCulture);
        int month = int.Parse(d.Groups[2].Value, CultureInfo.InvariantCulture);
        int year = DateTime.Today.Year;
        if (d.Groups[3].Success)
        {
            year = int.Parse(d.Groups[3].Value, CultureInfo.InvariantCulture);
            if (year < 100) year += year < 30 ? 2000 : 1900;
        }
        if (month is < 1 or > 12 || day < 1 || day > DateTime.DaysInMonth(year, month)) return false;
        value = new DateTime(year, month, day);
        if (d.Groups[4].Success)
        {
            int h = int.Parse(d.Groups[4].Value, CultureInfo.InvariantCulture);
            int mi = int.Parse(d.Groups[5].Value, CultureInfo.InvariantCulture);
            if (h > 23 || mi > 59) return false;
            value = value.AddHours(h).AddMinutes(mi);
            format = DateTimeFormat;
        }
        return true;
    }

    /// Hucreye yazilmis bilgiyi duzenleme kutusunda gosterilecek Turkce metne cevirir.
    public static string EditText(IXLCell cell)
    {
        if (cell.HasFormula) return "=" + FormulaLocale.ToTurkish(cell.FormulaA1);
        var v = cell.Value;
        switch (v.Type)
        {
            case XLDataType.Blank: return "";
            case XLDataType.Text: return v.GetText();
            case XLDataType.Boolean: return v.GetBoolean() ? "DOĞRU" : "YANLIŞ";
            case XLDataType.Error: return FormulaLocale.ErrorToTurkish(ErrorText(v.GetError()));
            case XLDataType.DateTime:
            {
                var dt = v.GetDateTime();
                if (dt.Date == new DateTime(1899, 12, 30) || dt.Date == new DateTime(1899, 12, 31)) return dt.ToString("HH:mm:ss", Tr);
                return dt.TimeOfDay == TimeSpan.Zero ? dt.ToString("dd.MM.yyyy", Tr) : dt.ToString("dd.MM.yyyy HH:mm", Tr);
            }
            case XLDataType.TimeSpan: return v.GetTimeSpan().ToString(@"hh\:mm\:ss", Tr);
            default:
            {
                double n = v.GetNumber();
                var fmt = cell.Style.NumberFormat.Format ?? "";
                if (fmt.Contains('%')) return (n * 100).ToString("0.###############", Tr) + "%";
                return n.ToString("0.###############", Tr);
            }
        }
    }

    public static string ErrorText(XLError e) => e switch
    {
        XLError.DivisionByZero => "#DIV/0!",
        XLError.NoValueAvailable => "#N/A",
        XLError.NameNotRecognized => "#NAME?",
        XLError.IncompatibleValue => "#VALUE!",
        XLError.CellReference => "#REF!",
        XLError.NumberInvalid => "#NUM!",
        XLError.NullValue => "#NULL!",
        _ => "#VALUE!",
    };

    /// Ayristirilmis girdiyi hucreye yazar (gecersiz formulde ClosedXML hata firlatir).
    public static void Apply(IXLCell cell, ParsedInput p)
    {
        switch (p.Kind)
        {
            case InputKind.Empty:
                cell.Clear(XLClearOptions.Contents);
                break;
            case InputKind.Formula:
                cell.FormulaA1 = p.FormulaEn!;
                break;
            default:
                cell.Value = p.Value;
                break;
        }
        if (p.NumberFormat != null) cell.Style.NumberFormat.Format = p.NumberFormat;
        else if (p.Kind == InputKind.Number && IsDateFormat(cell)) cell.Style.NumberFormat.Format = "General";
    }

    /// Hucrenin sayi bicimi tarih/saat mi (sayi girilince eski tarih bicimi kalmasin diye).
    private static bool IsDateFormat(IXLCell cell)
    {
        var f = cell.Style.NumberFormat.Format ?? "";
        int id = cell.Style.NumberFormat.NumberFormatId;
        if (id is >= 14 and <= 22) return true;
        var noLiterals = Regex.Replace(f, "\"[^\"]*\"|\\[[^\\]]*\\]", "");
        return noLiterals.Contains("yy", StringComparison.OrdinalIgnoreCase) || noLiterals.Contains("dd", StringComparison.OrdinalIgnoreCase);
    }
}
