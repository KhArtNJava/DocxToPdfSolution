namespace DocxToPdf.Fonts;

/// <summary>Конкретное начертание с флагами синтеза (если нужного файла нет).</summary>
public sealed record FontFace(TrueTypeFont Font, bool SynthBold, bool SynthItalic);

/// <summary>
/// Находит системные шрифты (Windows/Linux/macOS), подбирает замены
/// для отсутствующих гарнитур и запасные шрифты для недостающих символов.
/// </summary>
public sealed class FontManager
{
    readonly Dictionary<string, List<FontFaceInfo>> _index = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<(string, int), TrueTypeFont?> _loaded = new();
    readonly Dictionary<(string, bool, bool), FontFace> _faceCache = new();

    static readonly Dictionary<string, string[]> Substitutes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = ["Carlito", "Liberation Sans", "Arial"],
        ["Calibri Light"] = ["Carlito", "Calibri", "Liberation Sans", "Arial"],
        ["Cambria"] = ["Caladea", "Liberation Serif", "Times New Roman"],
        ["Cambria Math"] = ["Cambria", "Caladea", "Liberation Serif", "Times New Roman"],
        ["Arial"] = ["Liberation Sans", "Arimo", "Helvetica", "DejaVu Sans"],
        ["Helvetica"] = ["Arial", "Liberation Sans", "Arimo"],
        ["Arial Narrow"] = ["Liberation Sans Narrow", "Arial", "Liberation Sans"],
        ["Times New Roman"] = ["Liberation Serif", "Tinos", "Times", "DejaVu Serif"],
        ["Times"] = ["Times New Roman", "Liberation Serif", "Tinos"],
        ["Courier New"] = ["Liberation Mono", "Cousine", "Courier", "DejaVu Sans Mono"],
        ["Courier"] = ["Courier New", "Liberation Mono", "Cousine"],
        ["Consolas"] = ["Liberation Mono", "Courier New", "DejaVu Sans Mono"],
        ["Segoe UI"] = ["Selawik", "Open Sans", "Arial", "Liberation Sans", "DejaVu Sans"],
        ["Tahoma"] = ["DejaVu Sans", "Verdana", "Arial", "Liberation Sans"],
        ["Verdana"] = ["DejaVu Sans", "Tahoma", "Arial", "Liberation Sans"],
        ["Georgia"] = ["Gelasio", "DejaVu Serif", "Liberation Serif"],
        ["Aptos"] = ["Calibri", "Carlito", "Arial", "Liberation Sans"],
        ["Aptos Display"] = ["Aptos", "Calibri", "Carlito", "Arial", "Liberation Sans"],
        ["PT Sans"] = ["Arial", "Liberation Sans"],
        ["PT Serif"] = ["Times New Roman", "Liberation Serif"],
        ["MS Mincho"] = ["Noto Serif CJK JP", "IPAMincho", "Noto Sans CJK JP"],
        ["MS Gothic"] = ["Noto Sans CJK JP", "IPAGothic"],
        ["SimSun"] = ["Noto Serif CJK SC", "Noto Sans CJK SC", "WenQuanYi Zen Hei"],
    };

    static readonly string[] SansFallbacks = ["Arial", "Liberation Sans", "Arimo", "Carlito", "DejaVu Sans", "Noto Sans", "FreeSans", "Segoe UI", "Helvetica", "Verdana"];
    static readonly string[] SerifFallbacks = ["Times New Roman", "Liberation Serif", "Tinos", "DejaVu Serif", "Noto Serif", "FreeSerif", "Georgia"];
    static readonly string[] MonoFallbacks = ["Courier New", "Liberation Mono", "Cousine", "DejaVu Sans Mono", "Noto Sans Mono", "FreeMono", "Consolas"];
    static readonly string[] GlyphFallbacks =
        ["Segoe UI Symbol", "Arial Unicode MS", "DejaVu Sans", "Noto Sans", "Noto Sans Symbols", "Noto Sans Symbols2", "Symbola",
         "FreeSerif", "Segoe UI Emoji", "Microsoft YaHei", "SimSun", "MS Gothic", "Noto Sans CJK SC", "Noto Sans CJK JP",
         "WenQuanYi Zen Hei", "Droid Sans Fallback", "Arial", "Liberation Sans", "Times New Roman", "Liberation Serif"];

    public FontManager(IEnumerable<string>? extraDirs = null)
    {
        var dirs = new List<string>();
        if (extraDirs != null) dirs.AddRange(extraDirs);
        string win = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (!string.IsNullOrEmpty(win)) dirs.Add(win);
        string? windir = Environment.GetEnvironmentVariable("WINDIR");
        if (windir != null) dirs.Add(Path.Combine(windir, "Fonts"));
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (!string.IsNullOrEmpty(local)) dirs.Add(Path.Combine(local, "Microsoft", "Windows", "Fonts"));
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        dirs.AddRange(["/usr/share/fonts", "/usr/local/share/fonts", "/Library/Fonts", "/System/Library/Fonts", "/System/Library/Fonts/Supplemental"]);
        if (!string.IsNullOrEmpty(home))
            dirs.AddRange([Path.Combine(home, ".fonts"), Path.Combine(home, ".local/share/fonts"), Path.Combine(home, "Library/Fonts")]);

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dir in dirs.Distinct())
        {
            if (!Directory.Exists(dir)) continue;
            IEnumerable<string> files;
            try { files = Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories); }
            catch { continue; }
            foreach (var file in files)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext is not (".ttf" or ".ttc" or ".otf")) continue;
                string full = Path.GetFullPath(file);
                if (!seen.Add(full)) continue;
                foreach (var info in TrueTypeFont.ReadInfo(full))
                    foreach (var fam in info.Families)
                    {
                        if (!_index.TryGetValue(fam, out var list)) _index[fam] = list = [];
                        list.Add(info);
                    }
            }
        }
        if (_index.Count == 0)
            throw new InvalidOperationException("Не найдено ни одного TrueType-шрифта. Укажите каталог со шрифтами опцией --fonts.");
    }

    public int FamilyCount => _index.Count;

    TrueTypeFont? LoadFace(FontFaceInfo info)
    {
        var key = (info.Path, info.Index);
        if (_loaded.TryGetValue(key, out var f)) return f;
        try { f = TrueTypeFont.Load(info.Path, info.Index); }
        catch { f = null; }
        _loaded[key] = f;
        return f;
    }

    FontFace? TryFamily(string family, bool bold, bool italic)
    {
        if (!_index.TryGetValue(family, out var faces) || faces.Count == 0) return null;
        // Точное совпадение, затем совпадение по жирности, затем любое
        var ordered = faces
            .OrderByDescending(f => (f.Bold == bold ? 2 : 0) + (f.Italic == italic ? 1 : 0))
            .ThenBy(f => f.PostScriptName.Length);
        foreach (var info in ordered)
        {
            var font = LoadFace(info);
            if (font == null) continue;
            return new FontFace(font, bold && !info.Bold, italic && !info.Italic);
        }
        return null;
    }

    static bool LooksMono(string f) => f.Contains("mono", StringComparison.OrdinalIgnoreCase) || f.Contains("courier", StringComparison.OrdinalIgnoreCase) || f.Contains("consol", StringComparison.OrdinalIgnoreCase);

    static bool LooksSerif(string f)
    {
        f = f.ToLowerInvariant();
        if (f.Contains("sans")) return false;
        return f.Contains("serif") || f.Contains("times") || f.Contains("roman") || f.Contains("georgia") || f.Contains("garamond")
            || f.Contains("cambria") || f.Contains("book") || f.Contains("bodoni") || f.Contains("palatino") || f.Contains("century");
    }

    /// <summary>Возвращает начертание для гарнитуры (с подбором замены).</summary>
    public FontFace Get(string family, bool bold, bool italic)
    {
        var key = (family, bold, italic);
        if (_faceCache.TryGetValue(key, out var cached)) return cached;

        FontFace? face = TryFamily(family, bold, italic);
        if (face == null && Substitutes.TryGetValue(family, out var subs))
            foreach (var s in subs) if ((face = TryFamily(s, bold, italic)) != null) break;
        if (face == null)
        {
            var chain = LooksMono(family) ? MonoFallbacks : LooksSerif(family) ? SerifFallbacks : SansFallbacks;
            foreach (var s in chain.Concat(SansFallbacks)) if ((face = TryFamily(s, bold, italic)) != null) break;
        }
        if (face == null)
            foreach (var fam in _index.Keys.OrderBy(k => k)) if ((face = TryFamily(fam, bold, italic)) != null) break;

        _faceCache[key] = face ?? throw new InvalidOperationException("Не удалось загрузить ни один шрифт");
        return face;
    }

    readonly Dictionary<(int, string, bool, bool), FontFace?> _glyphFallbackCache = new();

    /// <summary>
    /// Ищет шрифт, содержащий символ: сначала замены исходной гарнитуры,
    /// затем шрифты того же типа (антиква/гротеск/моноширинный), затем любые запасные.
    /// </summary>
    public FontFace? FallbackFor(int codepoint, string family, bool bold, bool italic)
    {
        var key = (codepoint, family, bold, italic);
        if (_glyphFallbackCache.TryGetValue(key, out var r)) return r;
        r = null;
        var primary = Get(family, bold, italic);
        bool serif = LooksSerif(family) || primary.Font.IsSerif || LooksSerif(primary.Font.Family);
        var chain = new List<string>();
        if (Substitutes.TryGetValue(family, out var subs)) chain.AddRange(subs);
        chain.AddRange(LooksMono(family) ? MonoFallbacks : serif ? SerifFallbacks : SansFallbacks);
        chain.AddRange(GlyphFallbacks);
        foreach (var fam in chain.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var f = TryFamily(fam, bold, italic);
            if (f != null && f.Font.TryGetGlyph(codepoint, out _)) { r = f; break; }
        }
        _glyphFallbackCache[key] = r;
        return r;
    }

    /// <summary>
    /// Преобразует символы из символьных шрифтов (Symbol, Wingdings) в Unicode —
    /// на случай, если сам символьный шрифт недоступен.
    /// </summary>
    public static int MapSymbolChar(string family, int cp)
    {
        int c = cp >= 0xF000 && cp <= 0xF0FF ? cp - 0xF000 : cp;
        if (c > 0xFF) return cp;
        string f = family.ToLowerInvariant();
        if (f.StartsWith("wingdings"))
        {
            return c switch
            {
                0xA7 => 0x25AA, 0xA8 => 0x25FB, 0x6C => 0x25CF, 0x6E => 0x25A0, 0x6F => 0x25A1, 0x71 => 0x2751,
                0x75 => 0x25C6, 0x76 => 0x2756, 0x77 => 0x2B25, 0xD8 => 0x27A2, 0xE0 => 0x2192, 0xE8 => 0x2794,
                0xFC => 0x2714, 0xFB => 0x2718, 0xFE => 0x2611, 0xA1 => 0x25CB, 0x9F => 0x2022, 0x4A => 0x263A,
                0x46 => 0x261E, 0x2A => 0x2709, 0x3F => 0x270D, 0xF0 => 0x21E8, 0xDF => 0x2190, 0x9E => 0x00B7,
                _ => 0x2022
            };
        }
        if (f.StartsWith("symbol"))
        {
            const string greekLower = "αβχδεφγηιϕκλμνοπθρστυϖωξψζ";
            const string greekUpper = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ";
            if (c >= 'a' && c <= 'z') return greekLower[c - 'a'];
            if (c >= 'A' && c <= 'Z') return greekUpper[c - 'A'];
            return c switch
            {
                0xB7 => 0x2022, 0x2D => 0x2212, 0xA3 => 0x2264, 0xB3 => 0x2265, 0xB9 => 0x2260, 0xB1 => 0x00B1,
                0xB4 => 0x00D7, 0xB8 => 0x00F7, 0xA5 => 0x221E, 0xD6 => 0x221A, 0xE5 => 0x2211, 0xF2 => 0x222B,
                0xAE => 0x2192, 0xAC => 0x2190, 0xDE => 0x21D2, 0xB0 => 0x00B0, 0xC6 => 0x2205, 0xCE => 0x2208,
                0xD1 => 0x2207, 0xB6 => 0x2202, 0xBB => 0x2248, 0xBA => 0x2261, 0xF0 => 0x2022, 0xE0 => 0x25CA,
                0xA8 => 0x2663, 0xA9 => 0x2666, 0xAA => 0x2665, 0xAB => 0x2660, 0xA7 => 0x2022,
                _ => c
            };
        }
        return cp;
    }
}
