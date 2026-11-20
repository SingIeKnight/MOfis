using System.Windows.Documents;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MOfis.Common;
using MOfis.Writer;
using Xunit.Abstractions;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MOfis.Tests;

public class DocxRoundTripTests(ITestOutputHelper output)
{
    /// FlowDocument STA is parcacigi ister.
    private static void Sta(Action action)
    {
        Exception? error = null;
        var t = new Thread(() =>
        {
            try { action(); } catch (Exception e) { error = e; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) ExceptionDispatchInfo.Throw(error);
    }

    private static byte[] TinyPng()
    {
        var bmp = BitmapSource.Create(4, 4, 96, 96, PixelFormats.Bgr32, null, new byte[4 * 4 * 4], 16);
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(bmp));
        using var ms = new MemoryStream();
        enc.Save(ms);
        return ms.ToArray();
    }

    private static FlowDocument BuildSample()
    {
        var doc = DocHelpers.CreateBlank();
        doc.Blocks.Clear();

        var h1 = new Paragraph(new Run("Başlık Bir"));
        doc.Blocks.Add(h1);
        DocHelpers.ApplyStyle(h1, StyleCatalog.Heading1);

        var p = new Paragraph();
        p.Inlines.Add(new Run("Normal "));
        p.Inlines.Add(new Run("kalın") { FontWeight = FontWeights.Bold });
        p.Inlines.Add(new Run("eğik") { FontStyle = FontStyles.Italic });
        var deco = new TextDecorationCollection();
        deco.Add(TextDecorations.Underline);
        deco.Add(TextDecorations.Strikethrough);
        p.Inlines.Add(new Run("çizgili") { TextDecorations = deco });
        p.Inlines.Add(new Run("kırmızı") { Foreground = new SolidColorBrush(Color.FromRgb(0xC0, 0, 0)) });
        p.Inlines.Add(new Run("vurgulu") { Background = new SolidColorBrush(Colors.Yellow) });
        p.Inlines.Add(new Run("2") { BaselineAlignment = BaselineAlignment.Superscript });
        p.Inlines.Add(new Run("Arial") { FontFamily = new FontFamily("Arial"), FontSize = Units.PtToPx(14) });
        p.Inlines.Add(new Run("sekme\tsonrası"));
        doc.Blocks.Add(p);

        doc.Blocks.Add(new Paragraph(new Run("Ortalı")) { TextAlignment = TextAlignment.Center });
        var spaced = new Paragraph(new Run("İki yana, 1,5 satır")) { TextAlignment = TextAlignment.Justify };
        doc.Blocks.Add(spaced); // yazi tipi belgeden gelir; satir araligi ondan sonra hesaplanir
        spaced.LineHeight = 1.5 * spaced.FontFamily.LineSpacing * spaced.FontSize;

        var bullets = new List { MarkerStyle = TextMarkerStyle.Disc };
        bullets.ListItems.Add(new ListItem(new Paragraph(new Run("Madde 1"))));
        var item2 = new ListItem(new Paragraph(new Run("Madde 2")));
        var nested = new List { MarkerStyle = TextMarkerStyle.Decimal };
        nested.ListItems.Add(new ListItem(new Paragraph(new Run("Alt 1"))));
        nested.ListItems.Add(new ListItem(new Paragraph(new Run("Alt 2"))));
        item2.Blocks.Add(nested);
        bullets.ListItems.Add(item2);
        doc.Blocks.Add(bullets);

        // 3 sutun: ilk satirda 2 sutunluk birlesik hucre; ucuncu sutunda 2 satirlik birlesik hucre.
        var table = new Table();
        for (int i = 0; i < 3; i++) table.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var g = new TableRowGroup();
        var r1 = new TableRow();
        r1.Cells.Add(new TableCell(new Paragraph(new Run("A1-B1"))) { ColumnSpan = 2, Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE7, 0xFF)), BorderThickness = new Thickness(0, 0, 1, 1) });
        r1.Cells.Add(new TableCell(new Paragraph(new Run("C1-C2"))) { RowSpan = 2 });
        var r2 = new TableRow();
        r2.Cells.Add(new TableCell(new Paragraph(new Run("A2"))));
        r2.Cells.Add(new TableCell(new Paragraph(new Run("B2"))));
        g.Rows.Add(r1);
        g.Rows.Add(r2);
        table.RowGroups.Add(g);
        doc.Blocks.Add(table);

        var media = new Paragraph();
        var img = DocHelpers.CreateImageElement(new ImageData(TinyPng(), ".png"), 40, 30, 600);
        media.Inlines.Add(new InlineUIContainer(img));
        media.Inlines.Add(new Run(" "));
        media.Inlines.Add(new Hyperlink(new Run("örnek bağlantı")) { NavigateUri = new Uri("https://example.com/") });
        doc.Blocks.Add(media);

        var broken = new Paragraph(new Run("Yeni sayfa"));
        DocHelpers.MarkPageBreak(broken, true);
        doc.Blocks.Add(broken);
        return doc;
    }

    private static byte[] Export(FlowDocument doc, PageSettings? page = null)
    {
        using var ms = new MemoryStream();
        DocxExporter.Save(doc, page ?? new PageSettings(), ms);
        return ms.ToArray();
    }

    private List<string> Validate(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var pkg = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(pkg)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        foreach (var e in errors) output.WriteLine(e);
        return errors;
    }

    [Fact]
    public void Export_IsSchemaValid()
    {
        Sta(() =>
        {
            var bytes = Export(BuildSample());
            Assert.Empty(Validate(bytes));
        });
    }

    [Fact]
    public void RoundTrip_PreservesContent()
    {
        Sta(() =>
        {
            var bytes = Export(BuildSample());
            DocxImporter.Result r;
            using (var ms = new MemoryStream(bytes)) r = DocxImporter.Load(ms);
            Assert.Empty(r.Warnings);
            var blocks = r.Document.Blocks.ToList();

            // Baslik
            var h1 = Assert.IsType<Paragraph>(blocks[0]);
            Assert.Equal("Heading1", h1.Tag);
            Assert.Equal("Başlık Bir", new TextRange(h1.ContentStart, h1.ContentEnd).Text);

            // Yazi bicimleri
            var p = Assert.IsType<Paragraph>(blocks[1]);
            var runs = p.Inlines.OfType<Run>().ToList();
            Run Find(string text) => runs.Single(x => x.Text == text);
            Assert.Equal(FontWeights.Bold, Find("kalın").FontWeight);
            Assert.Equal(FontStyles.Italic, Find("eğik").FontStyle);
            var decos = Find("çizgili").TextDecorations;
            Assert.Contains(decos, d => d.Location == TextDecorationLocation.Underline);
            Assert.Contains(decos, d => d.Location == TextDecorationLocation.Strikethrough);
            Assert.Equal(Color.FromRgb(0xC0, 0, 0), ((SolidColorBrush)Find("kırmızı").Foreground).Color);
            Assert.Equal(Colors.Yellow, ((SolidColorBrush)Find("vurgulu").Background).Color);
            Assert.Equal(BaselineAlignment.Superscript, Find("2").BaselineAlignment);
            Assert.StartsWith("Arial", Find("Arial").FontFamily.Source);
            Assert.Equal(14, Math.Round(Units.PxToPt(Find("Arial").FontSize), 1));
            Assert.Equal("sekme\tsonrası", runs.Last().Text);

            // Paragraf bicimleri
            Assert.Equal(TextAlignment.Center, ((Paragraph)blocks[2]).TextAlignment);
            var spaced = (Paragraph)blocks[3];
            Assert.Equal(TextAlignment.Justify, spaced.TextAlignment);
            Assert.Equal(1.5, Math.Round(spaced.LineHeight / (spaced.FontFamily.LineSpacing * spaced.FontSize), 2));

            // Ic ice liste
            var list = Assert.IsType<List>(blocks[4]);
            Assert.Equal(TextMarkerStyle.Disc, list.MarkerStyle);
            Assert.Equal(2, list.ListItems.Count);
            var nested = list.ListItems.LastListItem.Blocks.OfType<List>().Single();
            Assert.Equal(TextMarkerStyle.Decimal, nested.MarkerStyle);
            Assert.Equal(2, nested.ListItems.Count);
            Assert.Equal(1, nested.StartIndex);

            // Birlesik hucreli tablo
            var table = Assert.IsType<Table>(blocks[5]);
            var rows = table.RowGroups[0].Rows;
            Assert.Equal(2, rows.Count);
            Assert.Equal(2, rows[0].Cells[0].ColumnSpan);
            Assert.Equal(2, rows[0].Cells[1].RowSpan);
            Assert.Equal(2, rows[1].Cells.Count);
            Assert.Equal(Color.FromRgb(0xE0, 0xE7, 0xFF), ((SolidColorBrush)rows[0].Cells[0].Background).Color);

            // Resim ve baglanti
            var media = (Paragraph)blocks[6];
            var container = media.Inlines.OfType<InlineUIContainer>().Single();
            var img = Assert.IsType<System.Windows.Controls.Image>(container.Child);
            Assert.Equal(40, Math.Round(img.Width));
            Assert.Equal(".png", ((ImageData)img.Tag).Extension);
            var link = media.Inlines.OfType<Hyperlink>().Single();
            Assert.Equal("https://example.com/", link.NavigateUri.AbsoluteUri);

            // Sayfa sonu
            Assert.True(((Paragraph)blocks[7]).BreakPageBefore);
        });
    }

    [Fact]
    public void RoundTrip_IsStable()
    {
        // Ac-kaydet dongusu belgeyi buyutmemeli (bos paragraf/tablo cogalmasi olmamali).
        Sta(() =>
        {
            var first = Export(BuildSample());
            FlowDocument again;
            using (var ms = new MemoryStream(first)) again = DocxImporter.Load(ms).Document;
            var second = Export(again);
            FlowDocument third;
            using (var ms = new MemoryStream(second)) third = DocxImporter.Load(ms).Document;

            static int Count(byte[] b)
            {
                using var ms = new MemoryStream(b);
                using var pkg = WordprocessingDocument.Open(ms, false);
                return pkg.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>().Count();
            }
            Assert.Equal(Count(first), Count(second));
            Assert.Equal(again.Blocks.Count, third.Blocks.Count);
            Assert.Empty(Validate(second));
        });
    }

    [Fact]
    public void PageSettings_RoundTrip()
    {
        Sta(() =>
        {
            var page = new PageSettings { Width = Units.CmToPx(29.7), Height = Units.CmToPx(21) };
            page.SetMargins(1.27);
            var bytes = Export(BuildSample(), page);
            using var ms = new MemoryStream(bytes);
            var r = DocxImporter.Load(ms);
            Assert.True(r.Page.IsLandscape);
            Assert.Equal(1.27, Math.Round(Units.PxToCm(r.Page.MarginLeft), 2));
        });
    }

    [Fact]
    public void Text_Turkish1254_IsDecoded()
    {
        Sta(() =>
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var path = Path.Combine(Path.GetTempPath(), $"mofis-test-{Guid.NewGuid():N}.txt");
            try
            {
                File.WriteAllBytes(path, Encoding.GetEncoding(1254).GetBytes("Işık ğüşiöç İSTANBUL"));
                var loaded = DocumentIO.Load(path);
                var text = new TextRange(loaded.Document.ContentStart, loaded.Document.ContentEnd).Text.Trim();
                Assert.Equal("Işık ğüşiöç İSTANBUL", text);
            }
            finally
            {
                File.Delete(path);
            }
        });
    }
}

public class SampleWriter
{
    /// Elle kontrol icin ornek dosya yazar: MOFIS_SAMPLE_OUT=yol dotnet test --filter SampleWriter
    [Fact]
    public void WriteSampleIfRequested()
    {
        var target = Environment.GetEnvironmentVariable("MOFIS_SAMPLE_OUT");
        if (string.IsNullOrEmpty(target)) return;
        Exception? error = null;
        var t = new Thread(() =>
        {
            try
            {
                var m = typeof(DocxRoundTripTests).GetMethod("BuildSample", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
                var doc = (System.Windows.Documents.FlowDocument)m.Invoke(null, null)!;
                MOfis.Writer.DocumentIO.Save(doc, new MOfis.Writer.PageSettings(), target);
            }
            catch (Exception e) { error = e; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) throw error;
    }
}

public class WordSampleImport(Xunit.Abstractions.ITestOutputHelper output)
{
    /// Word'un urettigi belgeyi (Flat OPC XML) DOCX'e cevirip ice aktarir:
    /// MOFIS_FLATOPC=yol.xml dotnet test --filter WordSampleImport
    [Fact]
    public void ImportWordSampleIfRequested()
    {
        var xmlPath = Environment.GetEnvironmentVariable("MOFIS_FLATOPC");
        if (string.IsNullOrEmpty(xmlPath)) return;
        var docxPath = Path.ChangeExtension(xmlPath, ".docx");
        DocumentFormat.OpenXml.Packaging.WordprocessingDocument.FromFlatOpcString(File.ReadAllText(xmlPath), docxPath, false).Dispose();

        Exception? error = null;
        var t = new Thread(() =>
        {
            try
            {
                var loaded = MOfis.Writer.DocumentIO.Load(docxPath);
                var back = Path.Combine(Path.GetDirectoryName(docxPath)!, "wordden-geri.docx");
                MOfis.Writer.DocumentIO.Save(loaded.Document, loaded.Page, back);
                using (var pkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(back, false))
                    output.WriteLine("Dogrulayici hata sayisi: " + new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pkg).Count());
                output.WriteLine("Uyarilar: " + string.Join(" | ", loaded.Warnings));
                foreach (var b in loaded.Document.Blocks)
                {
                    var text = new System.Windows.Documents.TextRange(b.ContentStart, b.ContentEnd).Text.Replace("\r\n", " / ");
                    output.WriteLine($"{b.GetType().Name,-10} {(b as System.Windows.Documents.Paragraph)?.Tag,-9} {text}");
                }
            }
            catch (Exception e) { error = e; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) throw error;
    }
}

public class DemoDocument
{
    /// README ekran goruntusu icin tanitim belgesi: MOFIS_DEMO_OUT=yol.docx dotnet test --filter DemoDocument
    [Fact]
    public void WriteDemoIfRequested()
    {
        var target = Environment.GetEnvironmentVariable("MOFIS_DEMO_OUT");
        if (string.IsNullOrEmpty(target)) return;
        Exception? error = null;
        var t = new Thread(() =>
        {
            try
            {

                var doc = MOfis.Writer.DocHelpers.CreateBlank();
                doc.Blocks.Clear();
                Paragraph P(string text, MOfis.Writer.ParaStyleDef? style = null)
                {
                    var p = new Paragraph(new Run(text));
                    doc.Blocks.Add(p);
                    if (style != null) MOfis.Writer.DocHelpers.ApplyStyle(p, style);
                    return p;
                }
                P("MOfis'e hoş geldiniz", MOfis.Writer.StyleCatalog.Title);
                var intro = new Paragraph();
                intro.Inlines.Add(new Run("MOfis, Türkçe öncelikli, "));
                intro.Inlines.Add(new Run("ücretsiz ve açık kaynak") { FontWeight = System.Windows.FontWeights.Bold });
                intro.Inlines.Add(new Run(" bir ofis programıdır. Gerçek "));
                intro.Inlines.Add(new Run(".docx") { FontStyle = System.Windows.FontStyles.Italic });
                intro.Inlines.Add(new Run(" dosyalarını açar ve kaydeder; Word'de hazırlanan belgeler burada, burada hazırlananlar Word'de sorunsuz açılır."));
                doc.Blocks.Add(intro);
                P("Neler yapabilirsiniz?", MOfis.Writer.StyleCatalog.Heading1);
                var list = new List { MarkerStyle = System.Windows.TextMarkerStyle.Disc };
                foreach (var s in new[] { "Başlık stilleri, listeler, tablolar, resimler ve bağlantılar", "Cetvelle girinti ve kenar boşluğu ayarı", "Türkçe yazım denetimi ve Türkçe kurallarına uygun Bul/Değiştir", "Yazdırma ve PDF olarak dışa aktarma" })
                    list.ListItems.Add(new ListItem(new Paragraph(new Run(s)) { Margin = new System.Windows.Thickness(0) }));
                doc.Blocks.Add(list);
                P("Dosya uyumluluğu", MOfis.Writer.StyleCatalog.Heading2);
                var table = new Table();
                var line = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x80, 0x80, 0x80));
                table.BorderBrush = line;
                table.BorderThickness = new System.Windows.Thickness(1, 1, 0, 0);
                for (int i = 0; i < 3; i++) table.Columns.Add(new TableColumn { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
                var g = new TableRowGroup();
                string[][] rows = [["Biçim", "Aç", "Kaydet"], ["Word belgesi (.docx)", "✓", "✓"], ["Zengin metin (.rtf)", "✓", "✓"], ["Düz metin (.txt)", "✓", "✓"]];
                for (int r = 0; r < rows.Length; r++)
                {
                    var row = new TableRow();
                    foreach (var c in rows[r])
                    {
                        var cell = new TableCell(new Paragraph(new Run(c)) { Margin = new System.Windows.Thickness(0) })
                        { BorderBrush = line, BorderThickness = new System.Windows.Thickness(0, 0, 1, 1), Padding = new System.Windows.Thickness(7, 3, 7, 3) };
                        if (r == 0) { cell.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0xE0, 0xE7, 0xFF)); cell.FontWeight = System.Windows.FontWeights.Bold; }
                        row.Cells.Add(cell);
                    }
                    g.Rows.Add(row);
                }
                table.RowGroups.Add(g);
                doc.Blocks.Add(table);
                P("Lisans: MIT. Kaynak kodu GitHub'da.");
                MOfis.Writer.DocumentIO.Save(doc, new MOfis.Writer.PageSettings(), target);
            }
            catch (Exception e) { error = e; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) throw error;
    }
}
