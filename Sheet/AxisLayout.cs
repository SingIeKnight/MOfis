namespace MOfis.Sheet;

/// Bir eksendeki (satirlar veya sutunlar) boyutlar: hepsi varsayilan boyutta, yalnizca farkli olanlar saklanir.
/// Boylece 1.048.576 satirda bile konum hesabi, ozel boyutlu satir sayisi kadar surer. Indisler 1'den baslar.
public sealed class AxisLayout(double defaultSize)
{
    private readonly SortedDictionary<int, double> _custom = [];

    public double Default { get; set; } = defaultSize;

    public double Size(int i) => _custom.TryGetValue(i, out var s) ? s : Default;

    public void Set(int i, double size)
    {
        if (Math.Abs(size - Default) < 0.01) _custom.Remove(i);
        else _custom[i] = size;
    }

    public void Clear() => _custom.Clear();

    public IEnumerable<KeyValuePair<int, double>> Custom => _custom;

    /// i. ogenin baslangic konumu.
    public double Offset(int i)
    {
        double pos = (i - 1) * Default;
        foreach (var (k, s) in _custom)
        {
            if (k >= i) break;
            pos += s - Default;
        }
        return pos;
    }

    /// pos konumundaki oge (sifir boyutlu/gizli ogeler atlanir).
    public int IndexAt(double pos)
    {
        if (pos < 0) return 1;
        double start = 0;
        int next = 1;
        foreach (var (k, s) in _custom)
        {
            // k'dan onceki varsayilan boyutlu blok
            double blockEnd = start + (k - next) * Default;
            if (pos < blockEnd) return next + (int)((pos - start) / Default);
            start = blockEnd;
            if (pos < start + s) return k;
            start += s;
            next = k + 1;
        }
        return next + (int)((pos - start) / Default);
    }
}
