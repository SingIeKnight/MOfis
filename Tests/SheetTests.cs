using ClosedXML.Excel;
using MOfis.Sheet;

namespace MOfis.Tests;

public class SheetTests
{
    [Theory]
    [InlineData("TOPLA(A1;B1;1,5)", "SUM(A1,B1,1.5)")]
    [InlineData("EĞER(A1>1,5;\"evet;hayır, belki\";\"hayır\")", "IF(A1>1.5,\"evet;hayır, belki\",\"hayır\")")]
    [InlineData("eğer(a1>0;topla(b1:b3);0)", "IF(A1>0,SUM(B1:B3),0)")]
    [InlineData("sum(a1:a3;max(1;2))", "SUM(A1:A3,MAX(1,2))")]
    [InlineData("min(i1:i5)", "MIN(I1:I5)")]
    [InlineData("'Satış 2026'!A1*2", "'Satış 2026'!A1*2")]
    [InlineData("DÜŞEYARA(\"x\";$A$1:$C$9;2;YANLIŞ)", "VLOOKUP(\"x\",$A$1:$C$9,2,FALSE)")]
    [InlineData("EĞERHATA(1/0;#YOK)", "IFERROR(1/0,#N/A)")]
    [InlineData("TOPLA.ÇARPIM(A1:A2;B1:B2)", "SUMPRODUCT(A1:A2,B1:B2)")]
    [InlineData("YUVARLA(2,345E+2;1)", "ROUND(2.345E+2,1)")]
    [InlineData("TOPLA(1:1)", "SUM(1:1)")]
    public void FormulaLocale_ToEnglish(string tr, string en) => Assert.Equal(en, FormulaLocale.ToEnglish(tr));

    [Theory]
    [InlineData("SUM(A1,B1,1.5)", "TOPLA(A1;B1;1,5)")]
    [InlineData("IF(A1>1.5,\"a,b\",\"c\")", "EĞER(A1>1,5;\"a,b\";\"c\")")]
    [InlineData("_xlfn.IFS(A1>1,1,TRUE,0)", "ÇOKEĞER(A1>1;1;DOĞRU;0)")]
    [InlineData("Sheet2!B3+1", "Sheet2!B3+1")]
    [InlineData("IFERROR(1/0,#DIV/0!)", "EĞERHATA(1/0;#SAYI/0!)")]
    public void FormulaLocale_ToTurkish(string en, string tr) => Assert.Equal(tr, FormulaLocale.ToTurkish(en));

    /// Esleme tablosundaki her Ingilizce fonksiyon ClosedXML tarafindan gercekten taniniyor mu?
    [Fact]
    public void FormulaLocale_AllMappedFunctionsKnownToEngine()
    {
        bool Known(string en)
        {
            // Her deneme ayri kitapta: ClosedXML ayristirma hatasindan sonra hucreyi bozuk birakiyor.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S");
            try
            {
                ws.Cell("C3").FormulaA1 = en + "()";
                var v = ws.Cell("C3").Value;
                return !(v.IsError && v.GetError() == XLError.NameNotRecognized);
            }
            catch (Exception ex)
            {
                // "Too few parameters" gibi hatalar: fonksiyon taniniyor demektir.
                return ex.Message.Contains("parameter", StringComparison.OrdinalIgnoreCase);
            }
        }

        Assert.False(Known("BOYLEBIRFONKSIYONYOK")); // testin kendisi bilinmeyeni yakalayabiliyor mu
        var unknown = FormulaLocale.FunctionPairs
            .Where(p => p.En is not ("TRUE" or "FALSE") && !Known(p.En) && !FormulaLocale.EngineMissing.Contains(p.En))
            .Select(p => $"{p.Tr}={p.En}").ToList();
        Assert.True(unknown.Count == 0, "Motorun tanımadığı fonksiyonlar: " + string.Join(", ", unknown));
    }

    [Fact]
    public void Formulas_TurkishInput_Evaluates()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        CellInput.Apply(ws.Cell("A1"), CellInput.Parse("1,5"));
        CellInput.Apply(ws.Cell("A2"), CellInput.Parse("2.500"));
        CellInput.Apply(ws.Cell("A3"), CellInput.Parse("=TOPLA(A1:A2;0,5)"));
        CellInput.Apply(ws.Cell("A4"), CellInput.Parse("=EĞER(A3>2500;\"büyük\";\"küçük\")"));
        CellInput.Apply(ws.Cell("A5"), CellInput.Parse("=YUVARLA(ORTALAMA(A1:A2);1)"));
        CellInput.Apply(ws.Cell("A6"), CellInput.Parse("=1/0"));
        Assert.Equal(2502.0, ws.Cell("A3").Value.GetNumber());
        Assert.Equal("büyük", ws.Cell("A4").Value.GetText());
        Assert.Equal(1250.8, ws.Cell("A5").Value.GetNumber());
        Assert.Equal("=TOPLA(A1:A2;0,5)", CellInput.EditText(ws.Cell("A3")));
        Assert.Equal("#SAYI/0!", SheetDisplay.Text(ws.Cell("A6")));

        // Bagimli hucre degisince formul yeniden hesaplanir.
        CellInput.Apply(ws.Cell("A1"), CellInput.Parse("10"));
        Assert.Equal(2510.5, ws.Cell("A3").Value.GetNumber());
    }

    [Theory]
    [InlineData("1,5", 1.5, null)]
    [InlineData("-1.234,25", -1234.25, null)]
    [InlineData("1234", 1234.0, null)]
    [InlineData("%15", 0.15, "0%")]
    [InlineData("12,5%", 0.125, "0.00%")]
    [InlineData("250 TL", 250.0, CellInput.CurrencyFormat)]
    [InlineData("₺1.000,5", 1000.5, CellInput.CurrencyFormat)]
    public void CellInput_Numbers(string input, double expected, string? format)
    {
        var p = CellInput.Parse(input);
        Assert.Equal(InputKind.Number, p.Kind);
        Assert.Equal(expected, p.Value.GetNumber(), 9);
        Assert.Equal(format, p.NumberFormat);
    }

    [Fact]
    public void CellInput_OtherKinds()
    {
        var d = CellInput.Parse("21.11.2026");
        Assert.Equal(InputKind.Date, d.Kind);
        Assert.Equal(new DateTime(2026, 11, 21), d.Value.GetDateTime());

        Assert.Equal(InputKind.Boolean, CellInput.Parse("doğru").Kind);
        Assert.Equal(InputKind.Text, CellInput.Parse("'0532").Kind);
        Assert.Equal("0532", CellInput.Parse("'0532").Value.GetText());
        Assert.Equal(InputKind.Text, CellInput.Parse("31.02.2026").Kind); // gecersiz tarih
        Assert.Equal(InputKind.Text, CellInput.Parse("Ankara").Kind);
        Assert.Equal(1.5, CellInput.Parse("1.5", invariantNumbers: true).Value.GetNumber());
    }

    [Fact]
    public void Display_UsesTurkishFormats()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        CellInput.Apply(ws.Cell("A1"), CellInput.Parse("1234,5"));
        CellInput.Apply(ws.Cell("A2"), CellInput.Parse("250 TL"));
        CellInput.Apply(ws.Cell("A3"), CellInput.Parse("%15"));
        CellInput.Apply(ws.Cell("A4"), CellInput.Parse("21.11.2026"));
        CellInput.Apply(ws.Cell("A5"), CellInput.Parse("doğru"));
        Assert.Equal("1234,5", SheetDisplay.Text(ws.Cell("A1")));
        Assert.Equal("250,00 ₺", SheetDisplay.Text(ws.Cell("A2")));
        Assert.Equal("15%", SheetDisplay.Text(ws.Cell("A3")));
        Assert.Equal("21.11.2026", SheetDisplay.Text(ws.Cell("A4")));
        Assert.Equal("DOĞRU", SheetDisplay.Text(ws.Cell("A5")));
    }

    [Fact]
    public void Xlsx_RoundTrip_KeepsFormulasAndValues()
    {
        var path = Path.Combine(Path.GetTempPath(), "mofis-test-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Satışlar");
                wb.AddWorksheet("İkinci");
                CellInput.Apply(ws.Cell("A1"), CellInput.Parse("Ürün"));
                CellInput.Apply(ws.Cell("B1"), CellInput.Parse("2,5"));
                CellInput.Apply(ws.Cell("B2"), CellInput.Parse("4"));
                CellInput.Apply(ws.Cell("B3"), CellInput.Parse("=TOPLA(B1:B2)"));
                ws.Range("C1:D1").Merge();
                SheetIO.Save(wb, path, 0);
            }
            var loaded = SheetIO.Load(path);
            using var wb2 = loaded.Workbook;
            var s = wb2.Worksheet("Satışlar");
            Assert.Equal("SUM(B1:B2)", s.Cell("B3").FormulaA1);
            Assert.Equal(6.5, s.Cell("B3").Value.GetNumber());
            Assert.Equal("Ürün", s.Cell("A1").GetString());
            Assert.Single(s.MergedRanges);
            Assert.Equal(2, wb2.Worksheets.Count);
            Assert.Empty(loaded.Warnings);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Csv_TurkishSemicolon_And1254()
    {
        var path = Path.Combine(Path.GetTempPath(), "mofis-test-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            var enc = System.Text.CodePagesEncodingProvider.Instance.GetEncoding(1254)!;
            File.WriteAllText(path, "Şehir;Tutar;Not\r\nİzmir;1.250,5;\"a;b\"\r\nÇorum;3;\"tırnak \"\"x\"\"\"\r\n", enc);
            var loaded = SheetIO.Load(path);
            var ws = loaded.Workbook.Worksheet(1);
            Assert.Equal("Şehir", ws.Cell("A1").GetString());
            Assert.Equal(1250.5, ws.Cell("B2").Value.GetNumber());
            Assert.Equal("a;b", ws.Cell("C2").GetString());
            Assert.Equal("tırnak \"x\"", ws.Cell("C3").GetString());

            SheetIO.Save(loaded.Workbook, path, 0);
            var text = File.ReadAllText(path);
            Assert.Contains("İzmir;1250,5;\"a;b\"", text);

            File.WriteAllText(path, "name,amount\nx,1.5\n");
            var en = SheetIO.Load(path).Workbook.Worksheet(1);
            Assert.Equal(1.5, en.Cell("B2").Value.GetNumber());
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// Excel ile karsilastirma icin ornek kitap: MOFIS_SHEET_DEMO_OUT=yol dotnet test --filter DemoWorkbook
    [Fact]
    public void DemoWorkbook()
    {
        var output = Environment.GetEnvironmentVariable("MOFIS_SHEET_DEMO_OUT");
        if (string.IsNullOrEmpty(output)) return;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Bütçe");
        string[,] input =
        {
            { "Kalem", "Tutar", "Oran", "Not" },
            { "Kira", "12.500", "%40", "=EĞER(B2>10000;\"yüksek\";\"normal\")" },
            { "Market", "4.250,75", "=B3/B$6", "=YUVARLA(B3*1,2;1)" },
            { "Fatura", "1.890,4", "=B4/B$6", "=MAK(B2:B4)-MİN(B2:B4)" },
            { "Ulaşım", "950", "=B5/B$6", "=EĞERSAY(B2:B5;\">1000\")" },
            { "Toplam", "=TOPLA(B2:B5)", "=TOPLA(C3:C5)", "=METNEÇEVİR(B6;\"0,00\")" },
            { "Tarih", "21.11.2026", "=YIL(B7)", "=ORTALAMA(B2:B5)" },
        };
        for (int r = 0; r < input.GetLength(0); r++)
        for (int c = 0; c < input.GetLength(1); c++)
            CellInput.Apply(ws.Cell(r + 1, c + 1), CellInput.Parse(input[r, c]));
        ws.Range("C3:C6").Style.NumberFormat.Format = "0.0%";
        SheetIO.Save(wb, output, 0);
    }

    [Fact]
    public void CopyCells_ShiftsRelativeFormulas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("B1").FormulaA1 = "A1*10";
        ws.Cell("B1").Style.Font.Bold = true;
        SheetOps.CopyCells(ws, new CellRange(1, 2, 1, 2), ws, 2, 2);
        Assert.Equal("A2*10", ws.Cell("B2").FormulaA1);
        Assert.Equal(30.0, ws.Cell("B2").Value.GetNumber());
        Assert.True(ws.Cell("B2").Style.Font.Bold);
    }

    [Fact]
    public void Address_Parse()
    {
        Assert.Equal(new CellRange(12, 2, 12, 2), SheetOps.ParseAddress("b12"));
        Assert.Equal(new CellRange(1, 1, 5, 3), SheetOps.ParseAddress("C5:A1"));
        Assert.Equal(new CellRange(1, 3, SheetGrid.MaxRow, 4), SheetOps.ParseAddress("C:D"));
        Assert.Equal(new CellRange(3, 1, 7, SheetGrid.MaxCol), SheetOps.ParseAddress("3:7"));
        Assert.Null(SheetOps.ParseAddress("ABCD1"));
        Assert.Equal("=TOPLA(A1:A2)", Run(ws => { ws.Cell("A1").Value = 1; ws.Cell("A2").Value = 2; return SheetOps.AutoSumFormula(ws, 3, 1); }));

        static string? Run(Func<IXLWorksheet, string?> f)
        {
            using var wb = new XLWorkbook();
            return f(wb.AddWorksheet("S"));
        }
    }

    [Fact]
    public void Axis_OffsetsAndHitTest()
    {
        var a = new AxisLayout(20);
        a.Set(3, 50);
        a.Set(5, 0);
        Assert.Equal(0, a.Offset(1));
        Assert.Equal(40, a.Offset(3));
        Assert.Equal(90, a.Offset(4));
        Assert.Equal(110, a.Offset(5));
        Assert.Equal(110, a.Offset(6));
        Assert.Equal(130, a.Offset(7));
        Assert.Equal(1, a.IndexAt(5));
        Assert.Equal(3, a.IndexAt(60));
        Assert.Equal(6, a.IndexAt(115));
        Assert.Equal(7, a.IndexAt(130));
    }
}
