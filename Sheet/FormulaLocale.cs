using System.Globalization;
using System.Text;

namespace MOfis.Sheet;

/// Formulleri Turkce Excel yazimi (=TOPLA(A1;B1;1,5)) ile dosyadaki Ingilizce yazim (SUM(A1,B1,1.5)) arasinda cevirir.
/// Belirtec tabanlidir: metin sabitleri ("..."), tirnakli sayfa adlari ('...') ve koseli parantezler oldugu gibi korunur.
public static class FormulaLocale
{
    private static readonly CultureInfo Tr = new("tr-TR");

    /// Turkce ad -> Ingilizce ad. Turkce Excel'in resmi fonksiyon adlari.
    private static readonly (string Tr, string En)[] Pairs =
    [
        ("TOPLA", "SUM"), ("ORTALAMA", "AVERAGE"), ("EĞER", "IF"), ("ÇOKEĞER", "IFS"),
        ("DÜŞEYARA", "VLOOKUP"), ("YATAYARA", "HLOOKUP"), ("İNDİS", "INDEX"), ("KAÇINCI", "MATCH"),
        ("BAĞ_DEĞ_SAY", "COUNT"), ("BAĞ_DEĞ_DOLU_SAY", "COUNTA"), ("BOŞLUKSAY", "COUNTBLANK"),
        ("EĞERSAY", "COUNTIF"), ("ÇOKEĞERSAY", "COUNTIFS"), ("ETOPLA", "SUMIF"), ("ÇOKETOPLA", "SUMIFS"),
        ("EĞERORTALAMA", "AVERAGEIF"), ("ÇOKEĞERORTALAMA", "AVERAGEIFS"), ("TOPLA.ÇARPIM", "SUMPRODUCT"),
        ("MAK", "MAX"), ("MİN", "MIN"), ("ORTANCA", "MEDIAN"), ("BÜYÜK", "LARGE"), ("KÜÇÜK", "SMALL"),
        ("STDSAPMA", "STDEV"), ("ENÇOK_OLAN", "MODE"),
        ("YUVARLA", "ROUND"), ("YUKARIYUVARLA", "ROUNDUP"), ("AŞAĞIYUVARLA", "ROUNDDOWN"),
        ("TAVANAYUVARLA", "CEILING"), ("TABANAYUVARLA", "FLOOR"), ("NSAT", "TRUNC"), ("TAMSAYI", "INT"),
        ("MUTLAK", "ABS"), ("KAREKÖK", "SQRT"), ("ÇARPIM", "PRODUCT"), ("KUVVET", "POWER"), ("MOD", "MOD"),
        ("ÜS", "EXP"), ("Pİ", "PI"), ("İŞARET", "SIGN"), ("ÇİFT", "EVEN"), ("TEK", "ODD"),
        ("S_SAYI_ÜRET", "RAND"), ("RASTGELEARADA", "RANDBETWEEN"), ("ÇARPINIM", "FACT"),
        ("VE", "AND"), ("YADA", "OR"), ("DEĞİL", "NOT"), ("DOĞRU", "TRUE"), ("YANLIŞ", "FALSE"),
        ("EĞERHATA", "IFERROR"), ("EĞERYOKSA", "IFNA"),
        ("EBOŞSA", "ISBLANK"), ("ESAYIYSA", "ISNUMBER"), ("EMETİNSE", "ISTEXT"), ("EHATALIYSA", "ISERROR"),
        ("EYOKSA", "ISNA"), ("ÇİFTMİ", "ISEVEN"), ("TEKMİ", "ISODD"), ("YOKSAY", "NA"),
        ("BİRLEŞTİR", "CONCATENATE"), ("ARALIKBİRLEŞTİR", "CONCAT"), ("METİNBİRLEŞTİR", "TEXTJOIN"),
        ("UZUNLUK", "LEN"), ("SOLDAN", "LEFT"), ("SAĞDAN", "RIGHT"), ("PARÇAAL", "MID"), ("KIRP", "TRIM"),
        ("BÜYÜKHARF", "UPPER"), ("KÜÇÜKHARF", "LOWER"), ("YAZIM.DÜZENİ", "PROPER"), ("YERİNEKOY", "SUBSTITUTE"),
        ("DEĞİŞTİR", "REPLACE"), ("BUL", "FIND"), ("MBUL", "SEARCH"), ("METNEÇEVİR", "TEXT"),
        ("SAYIYAÇEVİR", "VALUE"), ("YİNELE", "REPT"), ("ÖZDEŞ", "EXACT"), ("DAMGA", "CHAR"), ("KOD", "CODE"),
        ("TEMİZ", "CLEAN"), ("SAYIDÜZENLE", "FIXED"),
        ("BUGÜN", "TODAY"), ("ŞİMDİ", "NOW"), ("TARİH", "DATE"), ("ZAMAN", "TIME"), ("YIL", "YEAR"),
        ("AY", "MONTH"), ("GÜN", "DAY"), ("SAAT", "HOUR"), ("DAKİKA", "MINUTE"), ("SANİYE", "SECOND"),
        ("HAFTANINGÜNÜ", "WEEKDAY"), ("HAFTASAY", "WEEKNUM"), ("SERİAY", "EOMONTH"), ("SERİTARİH", "EDATE"),
        ("TAMİŞGÜNÜ", "NETWORKDAYS"), ("İŞGÜNÜ", "WORKDAY"), ("TARİHSAYISI", "DATEVALUE"),
        ("ZAMANSAYISI", "TIMEVALUE"), ("GÜN360", "DAYS360"), ("YILORAN", "YEARFRAC"), ("GÜNSAY", "DAYS"),
        ("ELEMAN", "CHOOSE"), ("SATIR", "ROW"), ("SÜTUN", "COLUMN"), ("SATIRSAY", "ROWS"), ("SÜTUNSAY", "COLUMNS"),
        ("DOLAYLI", "INDIRECT"), ("KAYDIR", "OFFSET"), ("ADRES", "ADDRESS"), ("ARA", "LOOKUP"),
        ("ÇAPRAZARA", "XLOOKUP"), ("KÖPRÜ", "HYPERLINK"),
        ("DEVRESEL_ÖDEME", "PMT"), ("GD", "FV"), ("BD", "PV"), ("NBD", "NPV"),
    ];

    /// Adi cevrilen ama ClosedXML 0.105 hesap motorunun henuz hesaplayamadigi fonksiyonlar (#AD? verir).
    /// Excel'de yazilmis dosyalarda adlari yine Turkce gorunsun diye eslemede tutulur.
    public static readonly HashSet<string> EngineMissing =
        ["IFS", "AVERAGEIF", "AVERAGEIFS", "SMALL", "MODE", "IFNA", "CHOOSE", "INDIRECT", "OFFSET", "ADDRESS", "LOOKUP", "XLOOKUP", "PV", "NPV"];

    private static readonly Dictionary<string, string> TrToEn = new(StringComparer.Ordinal);
    private static readonly Dictionary<string, string> EnToTr = new(StringComparer.Ordinal);

    /// Excel hata degerleri: Ingilizce -> Turkce.
    public static readonly IReadOnlyDictionary<string, string> Errors = new Dictionary<string, string>
    {
        ["#DIV/0!"] = "#SAYI/0!", ["#N/A"] = "#YOK", ["#NAME?"] = "#AD?", ["#VALUE!"] = "#DEĞER!",
        ["#REF!"] = "#BAŞV!", ["#NUM!"] = "#SAYI!", ["#NULL!"] = "#BOŞ!",
    };

    static FormulaLocale()
    {
        foreach (var (tr, en) in Pairs)
        {
            TrToEn[tr] = en;
            EnToTr[en] = tr;
        }
    }

    public static IEnumerable<(string Tr, string En)> FunctionPairs => Pairs;

    /// Kullanicinin yazdigi Turkce formulu (bastaki '=' olmadan) dosyadaki Ingilizce yazima cevirir.
    public static string ToEnglish(string formula) => Translate(formula, toEnglish: true);

    /// Dosyadaki Ingilizce formulu (bastaki '=' olmadan) Turkce gosterime cevirir.
    public static string ToTurkish(string formula) => Translate(formula, toEnglish: false);

    public static string ErrorToTurkish(string error) => Errors.TryGetValue(error, out var tr) ? tr : error;

    private static string Translate(string f, bool toEnglish)
    {
        var sb = new StringBuilder(f.Length + 8);
        int i = 0;
        int braceDepth = 0;
        while (i < f.Length)
        {
            char c = f[i];

            // Metin sabiti ve tirnakli sayfa adi: icindeki "" / '' kacislari dahil oldugu gibi kopyala.
            if (c is '"' or '\'')
            {
                int start = i++;
                while (i < f.Length)
                {
                    if (f[i] == c)
                    {
                        if (i + 1 < f.Length && f[i + 1] == c) { i += 2; continue; }
                        i++;
                        break;
                    }
                    i++;
                }
                sb.Append(f, start, i - start);
                continue;
            }

            if (c == '[')
            {
                int end = f.IndexOf(']', i);
                end = end < 0 ? f.Length : end + 1;
                sb.Append(f, i, end - i);
                i = end;
                continue;
            }

            if (c == '#')
            {
                // Hata sabiti (#YOK, #N/A ...)
                int start = i++;
                while (i < f.Length && (char.IsLetterOrDigit(f[i]) || f[i] is '/' or '?' or '!' or '0')) i++;
                var token = f[start..i];
                sb.Append(TranslateError(token, toEnglish));
                continue;
            }

            if (char.IsDigit(c) || (c == (toEnglish ? ',' : '.') && i + 1 < f.Length && char.IsDigit(f[i + 1]) && !PrevIsIdentChar(f, i)))
            {
                int start = i;
                while (i < f.Length && char.IsDigit(f[i])) i++;
                // Sayidan hemen sonra harf geliyorsa (ornegin satir araligi degil de 1A gibi) sadece rakamlari kopyala.
                char dec = toEnglish ? ',' : '.';
                if (i < f.Length && f[i] == dec && i + 1 < f.Length && char.IsDigit(f[i + 1]))
                {
                    sb.Append(f, start, i - start).Append(toEnglish ? '.' : ',');
                    i++;
                    start = i;
                    while (i < f.Length && char.IsDigit(f[i])) i++;
                }
                if (i < f.Length && f[i] is 'E' or 'e' && i + 1 < f.Length && (char.IsDigit(f[i + 1]) || (f[i + 1] is '+' or '-' && i + 2 < f.Length && char.IsDigit(f[i + 2]))))
                {
                    i += 2;
                    while (i < f.Length && char.IsDigit(f[i])) i++;
                }
                sb.Append(f, start, i - start);
                continue;
            }

            if (IsIdentStart(c))
            {
                int start = i;
                while (i < f.Length && IsIdentChar(f[i])) i++;
                var word = f[start..i];
                int j = i;
                while (j < f.Length && f[j] == ' ') j++;
                bool isFunction = j < f.Length && f[j] == '(';
                sb.Append(TranslateWord(word, isFunction, toEnglish));
                continue;
            }

            if (c == '{') braceDepth++;
            else if (c == '}') braceDepth = Math.Max(0, braceDepth - 1);

            if (toEnglish)
            {
                // Dizi sabitinde Turkce Excel ';' sutun, '\' satir ayiricidir; Ingilizce ',' ve ';'.
                if (braceDepth > 0 && c == '\\') sb.Append(';');
                else if (c == ';') sb.Append(',');
                else sb.Append(c);
            }
            else
            {
                if (braceDepth > 0 && c == ';') sb.Append('\\');
                else if (c == ',') sb.Append(';');
                else sb.Append(c);
            }
            i++;
        }
        return sb.ToString();
    }

    private static bool PrevIsIdentChar(string f, int i) => i > 0 && IsIdentChar(f[i - 1]);

    private static bool IsIdentStart(char c) => char.IsLetter(c) || c is '_' or '$';

    private static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c is '_' or '.' or '$' or '!';

    private static string TranslateError(string token, bool toEnglish)
    {
        foreach (var (en, tr) in Errors)
        {
            if (toEnglish && string.Equals(token, tr, StringComparison.OrdinalIgnoreCase)) return en;
            if (!toEnglish && string.Equals(token, en, StringComparison.OrdinalIgnoreCase)) return tr;
        }
        return token;
    }

    private static string TranslateWord(string word, bool isFunction, bool toEnglish)
    {
        // Sayfa1!A1 gibi sayfa onekli basvurulara dokunma (sadece hucre adresini buyut).
        if (word.Contains('!'))
        {
            int bang = word.LastIndexOf('!');
            return word[..(bang + 1)] + TranslateWord(word[(bang + 1)..], isFunction, toEnglish);
        }

        if (toEnglish)
        {
            var trKey = word.ToUpper(Tr);
            if (isFunction || trKey is "DOĞRU" or "YANLIŞ")
            {
                if (TrToEn.TryGetValue(trKey, out var en)) return en;
                // Ingilizce adla yazilmis olabilir (sum, if...).
                var enKey = word.ToUpperInvariant();
                if (isFunction || enKey is "TRUE" or "FALSE") return enKey;
            }
            // Hucre adresi (a1, $b$2) ise buyuk harfe cevir; tanimli adlara dokunma.
            return IsCellRef(word) ? word.ToUpperInvariant() : word;
        }
        else
        {
            var key = word.ToUpperInvariant();
            if (key.StartsWith("_XLFN.", StringComparison.Ordinal)) key = key[6..];
            if (key.StartsWith("_XLWS.", StringComparison.Ordinal)) key = key[6..];
            if ((isFunction || key is "TRUE" or "FALSE") && EnToTr.TryGetValue(key, out var tr)) return tr;
            return isFunction ? key : word;
        }
    }

    private static bool IsCellRef(string w)
    {
        int i = 0;
        if (i < w.Length && w[i] == '$') i++;
        int letters = 0;
        while (i < w.Length && char.IsAsciiLetter(w[i])) { i++; letters++; }
        if (letters is < 1 or > 3) return false;
        if (i < w.Length && w[i] == '$') i++;
        int digits = 0;
        while (i < w.Length && char.IsAsciiDigit(w[i])) { i++; digits++; }
        return digits > 0 && i == w.Length;
    }
}
