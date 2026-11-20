# MOfis: geliştirici notları

Türkçe öncelikli, açık kaynak (MIT) ofis programı. Windows, .NET 10, WPF.
Public repo: https://github.com/SingIeKnight/MOfis. Kullanıcı commit/push'u GitHub Desktop ile kendisi yapar;
değişiklik sonrası Türkçe commit Summary + Description verilir.

## Yapı

| Klasör | İçerik |
|---|---|
| `Common/` | `AppInfo` (ad, sürüm, AppData yolları, `Units` px/pt/twip/emu), `AppActions` (yeni/aç), `RecentFiles`, `InputDialog`, `ErrorLog` (`%LocalAppData%\MOfis\hatalar.log`), `Screenshot` (geliştirme aracı) |
| `Shell/` | `StartWindow` (başlangıç ekranı), `RelayCommand` |
| `Writer/` | Kelime işlemci: `WriterWindow` (şerit, olaylar), `DocumentModel` (`PageSettings`, `StyleCatalog`, `DocHelpers`), `DocxImporter` / `DocxExporter` (FlowDocument ↔ DOCX), `DocumentIO` (docx/rtf/txt, otomatik kurtarma, yazdırma/PDF), `Ruler`, `Pickers` (renk, tablo seçici, `LineIcons`, `TableOps`), `FindReplaceWindow` |
| `Sheet/` | Hesap tablosu: `SheetWindow` (şerit, formül çubuğu, düzenleme, pano, geri al, sayfa sekmeleri), `SheetGrid` (kendi çizdiğimiz sanallaştırılmış ızgara, `OnRender`), `AxisLayout` (satır/sütun boyutları, yalnızca farklı olanlar saklanır), `FormulaLocale` (TR↔EN formül çevirisi), `CellInput` (Türkçe giriş ayrıştırma, düzenleme metni), `CellFormat`/`SheetDisplay` (görünen metin, renkler), `SheetIO` (xlsx/csv), `SheetOps` (kopyala, yapıştır, otomatik toplam, adres), `SheetEngine` (formül doğrulama, döngü kurtarma, geri al anlık görüntüleri) |
| `Tests/` | xUnit (ana projeden `Compile Remove="Tests\**"` ile hariç). WPF nesneleri STA iş parçacığında çalıştırılır |
| `Theme.xaml` | Renkler (vurgu rengi indigo `#4338CA`) ve şerit stilleri (`RibbonLarge/Small/Toggle/Group/Tabs/Tab`) |

Hesap veri modeli: doğrudan ClosedXML `XLWorkbook` (ayrı model yok). Formül hücrede İngilizce saklanır
(`cell.FormulaA1`), kullanıcıya `FormulaLocale.ToTurkish` ile gösterilir; giriş `CellInput.Parse` → `CellInput.Apply`.

Yazı veri modeli: editör WPF `RichTextBox` + `FlowDocument`. Paragraf stili kimliği `Paragraph.Tag` alanında durur
(`Heading1` vb.). Resim baytları `ImageData` olarak `Image.Tag` alanında saklanır (EMF/WMF gösterilemese de kaybolmaz).

## Komutlar

```powershell
dotnet build MOfis.csproj
dotnet test Tests/MOfis.Tests.csproj
bin\Debug\net10.0-windows\MOfis.exe --shot start|writer cikti.png [dosya]   # pencere resmi + cikti.txt yerleşim ölçüleri
$env:MOFIS_RULERTEST=1; MOfis.exe --shot writer c.png dosya.docx              # cetvel sürükleme simülasyonu -> c.ruler.txt
$env:MOFIS_DEMO_OUT="x.docx"; dotnet test Tests/MOfis.Tests.csproj --filter DemoDocument   # README tanıtım belgesi
bin\Debug\net10.0-windows\MOfis.exe --shot sheet c.png [dosya.xlsx]          # Hesap: resim + c.txt (ilk hücrelerin görünen metni)
$env:MOFIS_SHEETTEST=1; MOfis.exe --shot sheet c.png                         # yazma/formül/döngü/geri al simülasyonu -> c.txt
$env:MOFIS_SHEET_DEMO_OUT="x.xlsx"; dotnet test Tests/MOfis.Tests.csproj --filter DemoWorkbook   # Excel karşılaştırması için örnek kitap
```

`--shot` çalışırken son kullanılanlar listesine dokunulmaz (`RecentFiles.Disabled`).

Excel ile karşılaştırma: Excel COM ile dosyayı açıp her hücrenin `.Text` ve `.FormulaLocal` değerini dökmek,
Hesap'ın `--shot sheet` çıktısıyla satır satır kıyaslamanın en hızlı yolu. COM'da biçim kodu verirken
`NumberFormat` değil `NumberFormatLocal` kullan (Türkçe Excel `NumberFormat`'ı da yerel kod sanıp bozuyor).

Ekran görüntüsü alıp `Read` ile bakmak, arayüz değişikliklerini doğrulamanın en ucuz yolu.

## Kurallar

- Arayüz metinleri Türkçe. Sayı ayrıştırma/biçimlendirmede her zaman `CultureInfo.InvariantCulture`;
  küçük/büyük harf dönüşümünde `ToLowerInvariant` / `OrdinalIgnoreCase` (Türkçe İ/ı tuzağı).
- Microsoft adı, logosu, Office ikonları veya lisanslı fontlar pakete konmaz (public repo, marka riski).
  Sistem fontları (Segoe UI, Segoe Fluent Icons, Calibri) yalnızca çalışma anında kullanılır.
- DOCX'te desteklenmeyen öğeler sessizce atılmaz: `DocxImporter.Warn` ile bildirilir, aynı dosyanın üzerine kaydetmeden önce sorulur.
- Kayıt önce `dosya.mofis-tmp` dosyasına yapılır, sonra yerine taşınır.

## Öğrenilmiş tuzaklar

- **RichTextBox, ilk çizimde `FlowDocument.PagePadding` değerini 5,0,5,0 yapar.** `ApplyPage` bu değişikliği
  `DependencyPropertyDescriptor` ile izleyip kenar boşluğundan düşer; yoksa yazı alanı Word'den 10 px dar kalır.
- **Editörde `BeginChange` bloğunu birden fazla ileti döngüsü boyunca açık tutma.** Cetvel sürüklemesinde bu,
  "Dispatcher işlemi askıya alındı ancak iletiler hala işleniyor" hatasını verdi. Sürüklerken yalnızca önizleme göster, bırakınca tek blokta uygula.
- `Math.Clamp(v, lo, hi)`, `lo > hi` olduğunda (kayan nokta, 1e-14) hata fırlatır; `WriterWindow.Clamp` kullan.
- İçe aktarılan tablo genişliği `ContentWidth - 2` piksele sığdırılır, yoksa son kenarlık çizgisi kırpılır.
- Word liste paragrafları `contextualSpacing` kullanır (aynı stildeki ardışık paragraflar arasında boşluk yok).
- OpenXml SDK 3.x: enum öznitelikleri için `.InnerText` (null olabilir) kullan; `W.SmartTagRun` yok.
- `is not var (a, b)` nullable tuple'da çalışmaz; `is not (var a, var b)` kullan.
- **ClosedXML hesap motoru kırılgan.** Ayrıştırma hatalı formül (`1+`, eksik parantez) henüz hesaplanmamış kitaba
  yazılırsa hücre bozuk kalır ("not in the chain"); bu yüzden formül önce `SheetEngine.Validate` ile boş bir deneme kitabında sınanır.
  Döngüsel başvuru bağımlılık zincirini kalıcı bozar (`RecalculateAllFormulas` da düzeltmez); tek çare kitabı belleğe kaydedip
  yeniden açmak (`SheetEngine.Rebuild`). `SheetWindow.ApplyChange` döngüde değişikliği geri alıp bunu yapar.
- ClosedXML 0.105'in hesaplayamadığı fonksiyonlar `FormulaLocale.EngineMissing` listesinde (ÇOKEĞER, ÇAPRAZARA, KÜÇÜK,
  EĞERORTALAMA, ELEMAN, DOLAYLI, KAYDIR...): #AD? verir ama adları yine çevrilir. Paket güncellenince test bu listeyi denetler.
- `METNEÇEVİR(x;"0,00")`: Türkçe Excel tırnak içindeki biçimi yerel yorumlar, ClosedXML İngilizce yorumlar (bilinen fark).
- `cell.GetFormattedString(tr-TR)` sayı/tarih biçimlerini doğru uygular; yerel ayara bağlı yerleşik biçimler (14, 22) elle Türkçe gösterilir.
- Word COM ile test: gizli Word'de `SaveAs2` / `ExportAsFixedFormat` takılır. Belgeyi `$doc.WordOpenXML` ile al,
  `WordprocessingDocument.FromFlatOpcString` ile DOCX'e çevir. PowerShell 5.1 betikleri Türkçe karakter için UTF-8 **BOM'lu** olmalı.
  Takılan `WINWORD` süreçlerini yalnızca son birkaç dakikada başlatılanlarla sınırlayarak kapat (kullanıcının açık Word'üne dokunma).

## Yol haritası

1. ~~Faz 1-2: başlangıç ekranı + Yazı~~ (yayınlandı, 0.1.0)
2. Faz 3: Hesap. 3a ~~ızgara, Türkçe formüller, xlsx/csv, pano, geri al~~ → 3b biçimlendirme (kalın, renk, hizalama,
   kenarlık, birleştir, sayı biçimi menüsü, sütun/satır boyutu sürükleme, satır/sütun ekle-sil, bölmeleri dondur; geri al
   anlık görüntüsüne stil de eklenmeli) → 3c sıralama, doldurma tutamacı, Bul/Değiştir, yazdır/PDF, formül yazarken
   hücreye tıklayıp başvuru ekleme, 0.2.0
3. Faz 4: paketleme (`build.ps1`, self-contained, ZIP, release), dosya ilişkilendirme
4. Sonra: yerel yapay zekâ yardımcısı, sayfa görünümü (sayfalar alt alta), resim boyutlandırma, üst/alt bilgi, dipnot
