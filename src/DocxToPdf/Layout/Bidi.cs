using System.Text;
namespace DocxToPdf.Layout;

/// <summary>
/// Упрощённый двунаправленный текст (иврит, арабский): направление символов, контекстные формы
/// арабских букв (Presentation Forms-B) и визуальное переупорядочивание строки.
/// </summary>
public static class Bidi
{
    /// <summary>0 — нейтральный, 1 — слева направо, 2 — справа налево.</summary>
    public static int Direction(int cp)
    {
        if (cp is >= 0x0590 and <= 0x08FF or >= 0xFB1D and <= 0xFDFF or >= 0xFE70 and <= 0xFEFF) return 2;
        if (cp < 0x80) return char.IsLetter((char)cp) ? 1 : 0;
        if (cp is >= 0x2000 and <= 0x2BFF) return 0;   // пунктуация, символы
        return Rune.IsValid(cp) && Rune.IsLetter(new Rune(cp)) ? 1 : 0;
    }

    // Арабские буквы: (буква, первая форма в блоке FExx, число форм: 2 — только справа, 4 — с обеих сторон)
    static readonly Dictionary<int, (int Iso, int Forms)> Arabic = new()
    {
        [0x0621] = (0xFE80, 1), [0x0622] = (0xFE81, 2), [0x0623] = (0xFE83, 2), [0x0624] = (0xFE85, 2), [0x0625] = (0xFE87, 2),
        [0x0626] = (0xFE89, 4), [0x0627] = (0xFE8D, 2), [0x0628] = (0xFE8F, 4), [0x0629] = (0xFE93, 2), [0x062A] = (0xFE95, 4),
        [0x062B] = (0xFE99, 4), [0x062C] = (0xFE9D, 4), [0x062D] = (0xFEA1, 4), [0x062E] = (0xFEA5, 4), [0x062F] = (0xFEA9, 2),
        [0x0630] = (0xFEAB, 2), [0x0631] = (0xFEAD, 2), [0x0632] = (0xFEAF, 2), [0x0633] = (0xFEB1, 4), [0x0634] = (0xFEB5, 4),
        [0x0635] = (0xFEB9, 4), [0x0636] = (0xFEBD, 4), [0x0637] = (0xFEC1, 4), [0x0638] = (0xFEC5, 4), [0x0639] = (0xFEC9, 4),
        [0x063A] = (0xFECD, 4), [0x0641] = (0xFED1, 4), [0x0642] = (0xFED5, 4), [0x0643] = (0xFED9, 4), [0x0644] = (0xFEDD, 4),
        [0x0645] = (0xFEE1, 4), [0x0646] = (0xFEE5, 4), [0x0647] = (0xFEE9, 4), [0x0648] = (0xFEED, 2), [0x0649] = (0xFEEF, 2),
        [0x064A] = (0xFEF1, 4),
    };

    static bool Transparent(int cp) => cp is >= 0x064B and <= 0x065F or 0x0670;
    static bool JoinsBoth(int cp) => cp == 0x0640 || (Arabic.TryGetValue(cp, out var a) && a.Forms == 4);
    static bool Joins(int cp) => cp == 0x0640 || (Arabic.TryGetValue(cp, out var a) && a.Forms >= 2);

    /// <summary>Контекстные формы: для каждой позиции UTF-16 — код формы (0, если не меняется).</summary>
    public static int[]? ShapeArabic(string text)
    {
        if (!text.Any(c => c is >= '؀' and <= 'ۿ')) return null;
        var res = new int[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            int cp = text[i];
            if (res[i] == -1) continue;   // алеф, вошедший в лигатуру
            if (!Arabic.TryGetValue(cp, out var info) || info.Forms == 1) continue;
            // Лигатура лам-алеф: лам + алеф (с хамзой/маддой) → одна форма (FEF5…FEFC); алеф не выводится
            if (cp == 0x0644)
            {
                int a = i + 1; while (a < text.Length && Transparent(text[a])) a++;
                if (a < text.Length && text[a] is '\u0622' or '\u0623' or '\u0625' or '\u0627')
                {
                    int lp = i - 1; while (lp >= 0 && Transparent(text[lp])) lp--;
                    bool jp = lp >= 0 && JoinsBoth(text[lp]);
                    int baseLig = text[a] switch { '\u0622' => 0xFEF5, '\u0623' => 0xFEF7, '\u0625' => 0xFEF9, _ => 0xFEFB };
                    res[i] = baseLig + (jp ? 1 : 0);
                    res[a] = -1;
                    continue;
                }
            }
            int p = i - 1; while (p >= 0 && Transparent(text[p])) p--;
            int n = i + 1; while (n < text.Length && Transparent(text[n])) n++;
            bool joinPrev = p >= 0 && JoinsBoth(text[p]);
            bool joinNext = info.Forms == 4 && n < text.Length && Joins(text[n]);
            int form = (joinPrev, joinNext) switch
            {
                (true, true) => 3,    // срединная
                (false, true) => 2,   // начальная
                (true, false) => 1,   // конечная
                _ => 0                // изолированная
            };
            res[i] = info.Iso + form;
        }
        return res;
    }

    // ================================================================
    //  UAX #9: уровни встраивания символов абзаца
    // ================================================================

    public enum BC : byte { L, R, AL, EN, ES, ET, AN, CS, NSM, BN, B, S, WS, ON, LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI }

    /// <summary>Двунаправленный класс символа (основные диапазоны Unicode).</summary>
    public static BC Class(int cp)
    {
        switch (cp)
        {
            case 0x202A: return BC.LRE; case 0x202B: return BC.RLE; case 0x202C: return BC.PDF;
            case 0x202D: return BC.LRO; case 0x202E: return BC.RLO;
            case 0x2066: return BC.LRI; case 0x2067: return BC.RLI; case 0x2068: return BC.FSI; case 0x2069: return BC.PDI;
            case 0x200E: return BC.L; case 0x200F: return BC.R; case 0x061C: return BC.AL;
            case 0x200B or 0x200C or 0x200D or 0x2060 or 0xFEFF or 0x00AD: return BC.BN;
            case '\t' or 0x1F: return BC.S;
            case '\n' or '\r' or 0x1C or 0x1D or 0x1E or 0x85 or 0x2029: return BC.B;
            case ' ' or 0x0C or 0x2028: return BC.WS;
            case '+' or '-' or 0x2212 or 0x207A or 0x207B or 0x208A or 0x208B: return BC.ES;
            case '#' or '$' or '%' or 0xA2 or 0xA3 or 0xA4 or 0xA5 or 0xB0 or 0xB1 or 0x2030 or 0x2031 or 0x2032 or 0x2033 or 0x20AC or 0x066A: return BC.ET;
            case ',' or '.' or ':' or '/' or 0xA0 or 0x060C or 0x202F or 0x2044 or 0xFE50 or 0xFE52 or 0xFE55 or 0xFF0C or 0xFF0E or 0xFF1A: return BC.CS;
        }
        if (cp is >= '0' and <= '9' or >= 0x06F0 and <= 0x06F9 or >= 0x2070 and <= 0x2079 or >= 0x2080 and <= 0x2089 or >= 0xFF10 and <= 0xFF19) return BC.EN;
        if (cp is >= 0x0660 and <= 0x0669 or 0x066B or 0x066C or >= 0x0600 and <= 0x0605 or 0x06DD) return BC.AN;
        if (cp is >= 0x20A0 and <= 0x20CF) return BC.ET;
        if (cp < 0x10000 && !Rune.IsValid(cp)) return BC.ON;
        var cat = Rune.IsValid(cp) ? Rune.GetUnicodeCategory(new Rune(cp)) : System.Globalization.UnicodeCategory.OtherNotAssigned;
        if (cat is System.Globalization.UnicodeCategory.NonSpacingMark or System.Globalization.UnicodeCategory.EnclosingMark) return BC.NSM;
        if (cp is >= 0x0590 and <= 0x05FF or >= 0x07C0 and <= 0x085F or >= 0xFB1D and <= 0xFB4F or >= 0x10800 and <= 0x10FFF or >= 0x1E800 and <= 0x1EFFF) return BC.R;
        if (cp is >= 0x0600 and <= 0x07BF or >= 0x0860 and <= 0x08FF or >= 0xFB50 and <= 0xFDFF or >= 0xFE70 and <= 0xFEFF) return BC.AL;
        if (cat == System.Globalization.UnicodeCategory.SpaceSeparator) return BC.WS;
        if (cat is System.Globalization.UnicodeCategory.Control or System.Globalization.UnicodeCategory.Format) return BC.BN;
        if (cat is System.Globalization.UnicodeCategory.UppercaseLetter or System.Globalization.UnicodeCategory.LowercaseLetter
            or System.Globalization.UnicodeCategory.TitlecaseLetter or System.Globalization.UnicodeCategory.ModifierLetter
            or System.Globalization.UnicodeCategory.OtherLetter or System.Globalization.UnicodeCategory.SpacingCombiningMark
            or System.Globalization.UnicodeCategory.LetterNumber or System.Globalization.UnicodeCategory.DecimalDigitNumber
            or System.Globalization.UnicodeCategory.OtherNumber) return BC.L;
        return BC.ON;
    }

    /// <summary>Управляющие символы направления (не выводятся).</summary>
    public static bool IsControl(int cp) => cp is >= 0x202A and <= 0x202E or >= 0x2066 and <= 0x2069 or 0x200E or 0x200F or 0x061C;

    /// <summary>
    /// Уровни встраивания для каждой позиции UTF-16 текста абзаца (UAX #9: X1–X10 с изолятами, W1–W7, N1–N2, I1–I2,
    /// L1 для разделителей сегментов и конца абзаца).
    /// </summary>
    public static byte[] ResolveLevels(string text, bool rtl)
    {
        int n = text.Length;
        var levels = new byte[n];
        if (n == 0) return levels;
        var types = new BC[n];
        var orig = new BC[n];
        for (int i = 0; i < n; i++)
        {
            int cp = char.IsHighSurrogate(text[i]) && i + 1 < n ? char.ConvertToUtf32(text[i], text[i + 1]) : text[i];
            types[i] = orig[i] = char.IsLowSurrogate(text[i]) && i > 0 && char.IsHighSurrogate(text[i - 1]) ? BC.BN : Class(cp);
        }
        byte paraLevel = (byte)(rtl ? 1 : 0);
        if (!types.Any(t => t is BC.R or BC.AL or BC.AN or BC.RLE or BC.RLO or BC.RLI or BC.FSI) && !rtl) return levels;

        // X1–X8: явные встраивания, переопределения, изоляты
        var stack = new List<(byte Level, sbyte Override, bool Isolate)> { (paraLevel, -1, false) };
        int overflowIsolates = 0, overflowEmbeds = 0, validIsolates = 0;
        for (int i = 0; i < n; i++)
        {
            var t = types[i];
            var top = stack[^1];
            byte NextOdd(byte l) => (byte)(l % 2 == 0 ? l + 1 : l + 2);
            byte NextEven(byte l) => (byte)(l % 2 == 0 ? l + 2 : l + 1);
            switch (t)
            {
                case BC.RLE or BC.LRE or BC.RLO or BC.LRO:
                {
                    byte nl = t is BC.RLE or BC.RLO ? NextOdd(top.Level) : NextEven(top.Level);
                    levels[i] = top.Level;
                    if (nl <= 125 && overflowIsolates == 0 && overflowEmbeds == 0)
                        stack.Add((nl, (sbyte)(t == BC.RLO ? 1 : t == BC.LRO ? 0 : -1), false));
                    else if (overflowIsolates == 0) overflowEmbeds++;
                    types[i] = BC.BN;
                    break;
                }
                case BC.RLI or BC.LRI or BC.FSI:
                {
                    levels[i] = top.Level;
                    if (top.Override >= 0) types[i] = top.Override == 1 ? BC.R : BC.L;
                    bool r = t == BC.RLI;
                    if (t == BC.FSI) r = FirstStrongIsRtl(orig, i + 1);
                    byte nl = r ? NextOdd(top.Level) : NextEven(top.Level);
                    if (nl <= 125 && overflowIsolates == 0 && overflowEmbeds == 0) { validIsolates++; stack.Add((nl, -1, true)); }
                    else overflowIsolates++;
                    break;
                }
                case BC.PDI:
                {
                    if (overflowIsolates > 0) overflowIsolates--;
                    else if (validIsolates > 0)
                    {
                        overflowEmbeds = 0;
                        while (!stack[^1].Isolate) stack.RemoveAt(stack.Count - 1);
                        stack.RemoveAt(stack.Count - 1);
                        validIsolates--;
                    }
                    top = stack[^1];
                    levels[i] = top.Level;
                    if (top.Override >= 0) types[i] = top.Override == 1 ? BC.R : BC.L;
                    break;
                }
                case BC.PDF:
                    if (overflowIsolates == 0)
                    {
                        if (overflowEmbeds > 0) overflowEmbeds--;
                        else if (!stack[^1].Isolate && stack.Count > 1) stack.RemoveAt(stack.Count - 1);
                    }
                    levels[i] = top.Level;
                    types[i] = BC.BN;
                    break;
                case BC.B:
                    levels[i] = paraLevel;
                    break;
                case BC.BN:
                    levels[i] = top.Level;
                    break;
                default:
                    levels[i] = top.Level;
                    if (top.Override >= 0) types[i] = top.Override == 1 ? BC.R : BC.L;
                    break;
            }
        }

        // W1–W7, N1–N2, I1–I2 по последовательностям одного уровня (BN пропускаются)
        int s = 0;
        while (s < n)
        {
            int e = s;
            while (e < n && levels[e] == levels[s]) e++;
            byte lvl = levels[s];
            byte prevLvl = s > 0 ? levels[s - 1] : paraLevel, nextLvl = e < n ? levels[e] : paraLevel;
            BC sos = Math.Max(prevLvl, lvl) % 2 == 1 ? BC.R : BC.L, eos = Math.Max(nextLvl, lvl) % 2 == 1 ? BC.R : BC.L;
            var idx = Enumerable.Range(s, e - s).Where(i => types[i] != BC.BN).ToList();
            ResolveRun(types, idx, sos, eos, lvl);
            foreach (var i in idx)
            {
                var t = types[i];
                if (lvl % 2 == 0) { if (t == BC.R) levels[i] = (byte)(lvl + 1); else if (t is BC.AN or BC.EN) levels[i] = (byte)(lvl + 2); }
                else if (t is BC.L or BC.EN or BC.AN) levels[i] = (byte)(lvl + 1);
            }
            s = e;
        }
        // BN — уровень предыдущего символа; L1: разделители сегментов и пробелы перед ними и в конце — уровень абзаца
        for (int i = 0; i < n; i++) if (types[i] == BC.BN && orig[i] is not (BC.RLI or BC.LRI or BC.FSI or BC.PDI)) levels[i] = i > 0 ? levels[i - 1] : paraLevel;
        bool trailing = true;
        for (int i = n - 1; i >= 0; i--)
        {
            var o = orig[i];
            if (o is BC.S or BC.B) { levels[i] = paraLevel; trailing = true; continue; }
            if (trailing && o is BC.WS or BC.BN or BC.LRI or BC.RLI or BC.FSI or BC.PDI or BC.LRE or BC.RLE or BC.LRO or BC.RLO or BC.PDF) { levels[i] = paraLevel; continue; }
            trailing = false;
        }
        return levels;
    }

    static bool FirstStrongIsRtl(BC[] types, int from)
    {
        int depth = 0;
        for (int i = from; i < types.Length; i++)
        {
            var t = types[i];
            if (t is BC.LRI or BC.RLI or BC.FSI) depth++;
            else if (t == BC.PDI) { if (depth == 0) return false; depth--; }
            else if (depth == 0 && t == BC.L) return false;
            else if (depth == 0 && t is BC.R or BC.AL) return true;
        }
        return false;
    }

    static void ResolveRun(BC[] t, List<int> idx, BC sos, BC eos, byte lvl)
    {
        int m = idx.Count;
        if (m == 0) return;
        // W1: NSM — тип предыдущего (после изолята — ON)
        BC prev = sos;
        foreach (var i in idx)
        {
            if (t[i] == BC.NSM) t[i] = prev is BC.LRI or BC.RLI or BC.FSI or BC.PDI ? BC.ON : prev;
            prev = t[i];
        }
        // W2: EN после AL — AN; W3: AL → R
        BC lastStrong = sos;
        foreach (var i in idx)
        {
            if (t[i] is BC.L or BC.R or BC.AL) lastStrong = t[i];
            else if (t[i] == BC.EN && lastStrong == BC.AL) t[i] = BC.AN;
        }
        foreach (var i in idx) if (t[i] == BC.AL) t[i] = BC.R;
        // W4: одиночный разделитель между числами
        for (int k = 1; k + 1 < m; k++)
        {
            BC a = t[idx[k - 1]], c = t[idx[k]], b = t[idx[k + 1]];
            if (c == BC.ES && a == BC.EN && b == BC.EN) t[idx[k]] = BC.EN;
            else if (c == BC.CS && a == b && a is BC.EN or BC.AN) t[idx[k]] = a;
        }
        // W5: ET рядом с EN → EN
        for (int k = 0; k < m; k++)
        {
            if (t[idx[k]] != BC.ET) continue;
            int j = k;
            while (j < m && t[idx[j]] == BC.ET) j++;
            bool adj = (k > 0 && t[idx[k - 1]] == BC.EN) || (j < m && t[idx[j]] == BC.EN);
            if (adj) for (int q = k; q < j; q++) t[idx[q]] = BC.EN;
            k = j - 1;
        }
        // W6: оставшиеся разделители → ON; W7: EN после L → L
        foreach (var i in idx) if (t[i] is BC.ES or BC.ET or BC.CS) t[i] = BC.ON;
        lastStrong = sos;
        foreach (var i in idx)
        {
            if (t[i] is BC.L or BC.R) lastStrong = t[i];
            else if (t[i] == BC.EN && lastStrong == BC.L) t[i] = BC.L;
        }
        // N1/N2: нейтральные между сильными одного направления (EN/AN как R) — это направление, иначе — направление встраивания
        BC Dir(BC x) => x is BC.EN or BC.AN ? BC.R : x;
        bool Neutral(BC x) => x is BC.B or BC.S or BC.WS or BC.ON or BC.LRI or BC.RLI or BC.FSI or BC.PDI;
        BC embed = lvl % 2 == 1 ? BC.R : BC.L;
        for (int k = 0; k < m; k++)
        {
            if (!Neutral(t[idx[k]])) continue;
            int j = k;
            while (j < m && Neutral(t[idx[j]])) j++;
            BC before = k > 0 ? Dir(t[idx[k - 1]]) : sos, after = j < m ? Dir(t[idx[j]]) : eos;
            BC res = before == after ? before : embed;
            for (int q = k; q < j; q++) t[idx[q]] = res;
            k = j - 1;
        }
    }

    // Зеркальные пары символов (Bidi_Mirrored) для текста справа налево
    static readonly Dictionary<int, int> Mirrors = BuildMirrors();
    static Dictionary<int, int> BuildMirrors()
    {
        var d = new Dictionary<int, int>();
        int[] pairs = ['(', ')', '<', '>', '[', ']', '{', '}', 0xAB, 0xBB, 0x2039, 0x203A, 0x2045, 0x2046, 0x207D, 0x207E, 0x208D, 0x208E,
            0x2208, 0x220B, 0x2209, 0x220C, 0x220A, 0x220D, 0x2264, 0x2265, 0x2266, 0x2267, 0x226A, 0x226B, 0x2282, 0x2283, 0x2286, 0x2287,
            0x2329, 0x232A, 0x27E8, 0x27E9, 0x27EA, 0x27EB, 0x2983, 0x2984, 0x3008, 0x3009, 0x300A, 0x300B, 0x300C, 0x300D, 0x300E, 0x300F,
            0x3010, 0x3011, 0xFF08, 0xFF09, 0xFF3B, 0xFF3D, 0xFF5B, 0xFF5D, 0xFE59, 0xFE5A, 0xFE5B, 0xFE5C];
        for (int i = 0; i + 1 < pairs.Length; i += 2) { d[pairs[i]] = pairs[i + 1]; d[pairs[i + 1]] = pairs[i]; }
        return d;
    }

    /// <summary>Зеркальный символ для нечётного уровня (скобки и т. п.), иначе тот же.</summary>
    public static int Mirror(int cp) => Mirrors.TryGetValue(cp, out var m) ? m : cp;

    /// <summary>
    /// Визуальный порядок строки (UAX #9, L1–L2): пробелы и табуляции в конце строки получают уровень абзаца,
    /// затем последовательности с уровнем ≥ k переворачиваются для k от максимального до наименьшего нечётного;
    /// глифы кусков с нечётным уровнем идут справа налево.
    /// </summary>
    public static void Reorder(List<Piece> pieces, bool rtlBase)
    {
        byte baseLevel = (byte)(rtlBase ? 1 : 0);
        for (int i = pieces.Count - 1; i >= 0 && pieces[i].Kind is PieceKind.Space or PieceKind.Tab; i--) pieces[i].Level = baseLevel;
        foreach (var p in pieces) if (p.Kind == PieceKind.Tab) p.Level = baseLevel;
        if (pieces.Count == 0 || pieces.All(p => p.Level == 0)) return;
        foreach (var p in pieces.Where(p => p.Level % 2 == 1)) ReverseGlyphs(p);
        int max = pieces.Max(p => p.Level), minOdd = pieces.Where(p => p.Level % 2 == 1).Select(p => (int)p.Level).DefaultIfEmpty(max + 1).Min();
        for (int k = max; k >= minOdd && k >= 1; k--)
        {
            int i = 0;
            while (i < pieces.Count)
            {
                if (pieces[i].Level < k) { i++; continue; }
                int j = i;
                while (j < pieces.Count && pieces[j].Level >= k) j++;
                pieces.Reverse(i, j - i);
                i = j;
            }
        }
    }

    static void ReverseGlyphs(Piece p)
    {
        p.Gids.Reverse(); p.Texts.Reverse(); p.Adv.Reverse(); p.Pos.Reverse();
    }
}
