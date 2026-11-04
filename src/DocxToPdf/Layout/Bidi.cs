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
            if (!Arabic.TryGetValue(cp, out var info) || info.Forms == 1) continue;
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

    /// <summary>
    /// Визуальный порядок строки: последовательности фрагментов справа налево переворачиваются
    /// (вместе с глифами внутри фрагментов); при базовом направлении справа налево — вся строка.
    /// </summary>
    public static void Reorder(List<Piece> pieces, bool rtlBase)
    {
        if (!rtlBase && !pieces.Any(p => p.Dir == 2)) return;
        foreach (var p in pieces.Where(p => p.Dir == 2)) ReverseGlyphs(p);
        if (rtlBase)
        {
            pieces.Reverse();
            ReverseRuns(pieces, 1);   // фрагменты слева направо внутри — обратно в прямой порядок
        }
        else ReverseRuns(pieces, 2);
    }

    /// <summary>Переворачивает максимальные последовательности с направлением dir (нейтральные между ними включаются).</summary>
    static void ReverseRuns(List<Piece> pieces, int dir)
    {
        int i = 0;
        while (i < pieces.Count)
        {
            // Число после арабского текста (Dir = 3) — член фрагмента справа налево, но глифы не переворачиваются
            bool Member(Piece p) => p.Dir == dir || (dir == 2 && p.Dir == 3);
            if (!Member(pieces[i])) { i++; continue; }
            int j = i, last = i;
            while (j + 1 < pieces.Count && (Member(pieces[j + 1]) || pieces[j + 1].Dir == 0))
            {
                j++;
                if (Member(pieces[j])) last = j;
            }
            if (last > i) pieces.Reverse(i, last - i + 1);
            i = last + 1;
        }
    }

    static void ReverseGlyphs(Piece p)
    {
        p.Gids.Reverse(); p.Texts.Reverse(); p.Adv.Reverse(); p.Pos.Reverse();
    }
}
