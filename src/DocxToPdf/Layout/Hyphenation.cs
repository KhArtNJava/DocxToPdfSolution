using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace DocxToPdf.Layout;

/// <summary>
/// Автоматические переносы по шаблонам Лианга (формат словарей Hyphen/LibreOffice: hyph_xx_YY.dic).
/// Словари встроены в программу (сжатые), выбираются по языку текста (ru-RU, en-US, de …).
/// </summary>
public sealed class Hyphenator
{
    readonly Dictionary<string, byte[]> _patterns = new(StringComparer.Ordinal);
    readonly Dictionary<string, int[]> _cache = new(StringComparer.Ordinal);
    int _maxLen;
    public int LeftMin { get; private set; } = 2;
    public int RightMin { get; private set; } = 2;

    static readonly Dictionary<string, Hyphenator?> Loaded = new(StringComparer.OrdinalIgnoreCase);
    static readonly object Lock = new();
    static string[]? _resources;

    static Hyphenator() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>Словарь для языка (BCP 47 или xx_YY); null — словаря нет.</summary>
    public static Hyphenator? For(string? lang)
    {
        if (string.IsNullOrWhiteSpace(lang)) return null;
        string key = lang.Trim().Replace('-', '_');
        lock (Lock)
        {
            if (Loaded.TryGetValue(key, out var h)) return h;
            h = null;
            var name = FindResource(key);
            if (name != null)
            {
                try { h = Load(name); } catch { h = null; }
            }
            Loaded[key] = h;
            return h;
        }
    }

    static string? FindResource(string key)
    {
        _resources ??= typeof(Hyphenator).Assembly.GetManifestResourceNames().Where(n => n.EndsWith(".dic.gz", StringComparison.Ordinal)).ToArray();
        string Id(string res) { int i = res.LastIndexOf("hyph_", StringComparison.Ordinal); return i < 0 ? "" : res[(i + 5)..^7]; }
        var ids = _resources.Select(r => (Res: r, Id: Id(r))).ToList();
        // Точное совпадение (ru_RU), затем язык (ru → ru_RU), предпочтительные варианты для неоднозначных языков
        var exact = ids.FirstOrDefault(x => x.Id.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (exact.Res != null) return exact.Res;
        string language = key.Split('_')[0].ToLowerInvariant();
        string? preferred = language switch { "en" => "en_US", "de" => "de_DE", "pt" => "pt_PT", "no" or "nn" or "nb" => "nb_NO", _ => null };
        if (preferred != null && ids.FirstOrDefault(x => x.Id == preferred).Res is { } pr) return pr;
        return ids.FirstOrDefault(x => x.Id.Equals(language, StringComparison.OrdinalIgnoreCase) || x.Id.StartsWith(language + "_", StringComparison.OrdinalIgnoreCase)).Res;
    }

    static Hyphenator Load(string resource)
    {
        using var s = typeof(Hyphenator).Assembly.GetManifestResourceStream(resource)!;
        using var gz = new GZipStream(s, CompressionMode.Decompress);
        var ms = new MemoryStream();
        gz.CopyTo(ms);
        return Parse(ms.ToArray());
    }

    /// <summary>Разбор словаря: первая строка — кодировка, далее параметры и шаблоны (уровни NEXTLEVEL объединяются).</summary>
    public static Hyphenator Parse(byte[] data)
    {
        int nl = Array.IndexOf(data, (byte)'\n');
        string encName = Encoding.ASCII.GetString(data, 0, nl < 0 ? data.Length : nl).Trim();
        Encoding enc = encName.ToUpperInvariant() switch
        {
            "UTF-8" or "UTF8" => new UTF8Encoding(false),
            "ISO8859-1" => Encoding.Latin1,
            "KOI8-R" => Encoding.GetEncoding(20866),
            "KOI8-U" => Encoding.GetEncoding(21866),
            "MICROSOFT-CP1251" or "CP1251" => Encoding.GetEncoding(1251),
            _ when encName.StartsWith("ISO8859-", StringComparison.OrdinalIgnoreCase) => Encoding.GetEncoding("iso-8859-" + encName[8..]),
            _ => new UTF8Encoding(false),
        };
        string text = enc.GetString(data, nl + 1, data.Length - nl - 1);
        var h = new Hyphenator();
        foreach (var raw in text.Split('\n'))
        {
            string line = raw.Trim();
            if (line.Length == 0 || line[0] is '%' or '#') continue;
            if (line.StartsWith("LEFTHYPHENMIN", StringComparison.Ordinal)) { if (int.TryParse(line[13..].Trim(), out var l)) h.LeftMin = Math.Max(1, l); continue; }
            if (line.StartsWith("RIGHTHYPHENMIN", StringComparison.Ordinal)) { if (int.TryParse(line[14..].Trim(), out var r)) h.RightMin = Math.Max(1, r); continue; }
            if (line.StartsWith("COMPOUND", StringComparison.Ordinal) || line.StartsWith("NOHYPHEN", StringComparison.Ordinal) || line == "NEXTLEVEL") continue;
            int slash = line.IndexOf('/');
            if (slash >= 0) line = line[..slash];           // нестандартные переносы (замены) — только позиция
            int sp = line.IndexOf(' ');
            if (sp >= 0) line = line[..sp];
            var letters = new StringBuilder();
            var vals = new List<byte> { 0 };
            foreach (char c in line)
            {
                if (c is >= '0' and <= '9') vals[^1] = (byte)(c - '0');
                else { letters.Append(char.ToLowerInvariant(c)); vals.Add(0); }
            }
            if (letters.Length == 0) continue;
            string key = letters.ToString();
            var arr = vals.ToArray();
            if (h._patterns.TryGetValue(key, out var old))
                for (int i = 0; i < arr.Length && i < old.Length; i++) arr[i] = Math.Max(arr[i], old[i]);
            h._patterns[key] = arr;
            h._maxLen = Math.Max(h._maxLen, key.Length);
        }
        return h;
    }

    /// <summary>
    /// Допустимые переносы слова: длины префиксов, после которых можно поставить перенос
    /// (с учётом наименьшего числа букв слева и справа).
    /// </summary>
    public int[] Points(string word, int leftMin = 0, int rightMin = 0)
    {
        int lm = leftMin > 0 ? leftMin : LeftMin, rm = rightMin > 0 ? rightMin : RightMin;
        string w = word.ToLowerInvariant().Replace('’', '\'');
        string cacheKey = lm + ":" + rm + ":" + w;
        lock (_cache)
        {
            if (_cache.TryGetValue(cacheKey, out var cached)) return cached;
        }
        int n = w.Length;
        var res = new List<int>();
        if (n >= lm + rm)
        {
            string s = "." + w + ".";
            var v = new byte[s.Length + 1];
            for (int i = 0; i < s.Length; i++)
            {
                for (int len = 1; len <= _maxLen && i + len <= s.Length; len++)
                {
                    if (!_patterns.TryGetValue(s.Substring(i, len), out var pat)) continue;
                    for (int k = 0; k < pat.Length; k++) if (pat[k] > v[i + k]) v[i + k] = pat[k];
                }
            }
            // Перенос после префикса длины L: позиция L+1 в строке с точками
            for (int L = lm; L <= n - rm; L++)
                if ((v[L + 1] & 1) == 1) res.Add(L);
        }
        var arr = res.ToArray();
        lock (_cache) { if (_cache.Count < 50000) _cache[cacheKey] = arr; }
        return arr;
    }
}
