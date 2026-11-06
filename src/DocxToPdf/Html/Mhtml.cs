using System.Text;

namespace DocxToPdf.Html;

/// <summary>
/// Веб-архив MHTML (.mht, .mhtml; RFC 2557): MIME-сообщение multipart/related с HTML-страницей
/// и её ресурсами (картинки, стили), адресуемыми по Content-Location (URL) или Content-ID (cid:).
/// Понимает архивы Chromium/Edge («Сохранить как: веб-страница, один файл»), Word и Internet Explorer.
/// </summary>
public sealed class MhtmlArchive
{
    public sealed record Part(string ContentType, string? Charset, string? Location, string? ContentId, byte[] Data);

    public List<Part> Parts { get; } = [];
    /// <summary>Основная HTML-часть (параметр start, иначе первая text/html).</summary>
    public Part? Main { get; private set; }

    readonly Dictionary<string, Part> _byLocation = new(StringComparer.Ordinal);
    readonly Dictionary<string, Part> _byLocationCi = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, Part> _byCid = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Похож ли файл на MHTML: MIME-заголовки до начала разметки и тип multipart/related (или text/html с MIME-Version).</summary>
    public static bool Looks(byte[] b)
    {
        int len = Math.Min(b.Length, 16384);
        string head = Encoding.Latin1.GetString(b, 0, len);
        if (head.StartsWith("﻿")) head = head[1..];
        else if (head.StartsWith("ï»¿")) head = head[3..];
        string t = head.TrimStart();
        if (t.StartsWith("<")) return false;
        // Заголовки — до первой пустой строки
        int end = t.IndexOf("\n\n", StringComparison.Ordinal);
        int end2 = t.IndexOf("\r\n\r\n", StringComparison.Ordinal);
        if (end2 >= 0 && (end < 0 || end2 < end)) end = end2;
        string hdr = (end >= 0 ? t[..end] : t).ToLowerInvariant();
        if (!System.Text.RegularExpressions.Regex.IsMatch(hdr, @"^[\w-]+:", System.Text.RegularExpressions.RegexOptions.Multiline)) return false;
        return hdr.Contains("multipart/related") || hdr.Contains("mime-version:") && hdr.Contains("content-type:")
            || hdr.Contains("multipart/") && hdr.Contains("boundary");
    }

    public static MhtmlArchive Parse(byte[] data)
    {
        var a = new MhtmlArchive();
        int start = 0;
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) start = 3;
        var (headers, bodyStart) = ReadHeaders(data, start);
        string? startCid = null;
        a.ParseEntity(data, headers, bodyStart, data.Length, ref startCid, top: true);

        foreach (var p in a.Parts)
        {
            if (p.Location != null)
            {
                foreach (var key in LocationKeys(p.Location))
                {
                    a._byLocation.TryAdd(key, p);
                    a._byLocationCi.TryAdd(key, p);
                }
            }
            if (p.ContentId != null) a._byCid.TryAdd(p.ContentId, p);
        }
        if (startCid != null && a._byCid.TryGetValue(startCid, out var sp)) a.Main = sp;
        a.Main ??= a.Parts.FirstOrDefault(p => p.ContentType is "text/html" or "application/xhtml+xml")
                   ?? a.Parts.FirstOrDefault(p => p.ContentType.StartsWith("text/", StringComparison.Ordinal));
        return a;
    }

    // ---------------- MIME ----------------

    /// <summary>Заголовки сущности: имя (в нижнем регистре) → значение (со склеенными продолжениями строк).</summary>
    static (Dictionary<string, string> Headers, int BodyStart) ReadHeaders(byte[] d, int p)
    {
        var h = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string? last = null;
        while (p < d.Length)
        {
            int e = Array.IndexOf(d, (byte)'\n', p);
            if (e < 0) e = d.Length;
            int le = e > p && d[e - 1] == '\r' ? e - 1 : e;
            string line = Encoding.Latin1.GetString(d, p, le - p);
            p = Math.Min(d.Length, e + 1);
            if (line.Length == 0) break;
            if ((line[0] == ' ' || line[0] == '\t') && last != null) { h[last] += " " + line.Trim(); continue; }
            int c = line.IndexOf(':');
            if (c <= 0) { if (h.Count == 0) continue; break; }
            last = line[..c].Trim();
            string v = line[(c + 1)..].Trim();
            h[last] = h.TryGetValue(last, out var old) ? old : v;   // первое значение
        }
        return (h, p);
    }

    /// <summary>Значение заголовка без параметров и параметры (name=value; кавычки снимаются).</summary>
    static (string Value, Dictionary<string, string> Params) SplitHeader(string? v)
    {
        var ps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(v)) return ("", ps);
        var parts = new List<string>();
        var sb = new StringBuilder();
        bool q = false;
        foreach (char ch in v)
        {
            if (ch == '"') { q = !q; sb.Append(ch); }
            else if (ch == ';' && !q) { parts.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(ch);
        }
        parts.Add(sb.ToString());
        for (int i = 1; i < parts.Count; i++)
        {
            int eq = parts[i].IndexOf('=');
            if (eq <= 0) continue;
            string val = parts[i][(eq + 1)..].Trim();
            if (val.Length >= 2 && val[0] == '"' && val[^1] == '"') val = val[1..^1];
            ps[parts[i][..eq].Trim()] = val;
        }
        return (parts[0].Trim().ToLowerInvariant(), ps);
    }

    void ParseEntity(byte[] d, Dictionary<string, string> headers, int s, int e, ref string? startCid, bool top)
    {
        var (type, ps) = SplitHeader(headers.GetValueOrDefault("Content-Type"));
        if (type.Length == 0) type = "text/plain";
        if (type.StartsWith("multipart/", StringComparison.Ordinal) && ps.TryGetValue("boundary", out var boundary) && boundary.Length > 0)
        {
            if (top || startCid == null) if (ps.TryGetValue("start", out var st)) startCid = st.Trim().Trim('<', '>');
            foreach (var (ps2, pe2) in SplitMultipart(d, s, e, boundary))
            {
                var (h2, b2) = ReadHeaders(d, ps2);
                ParseEntity(d, h2, Math.Min(b2, pe2), pe2, ref startCid, top: false);
            }
            return;
        }
        string cte = headers.GetValueOrDefault("Content-Transfer-Encoding")?.Trim().ToLowerInvariant() ?? "";
        var raw = d.AsSpan(s, Math.Max(0, e - s));
        byte[] body = cte switch
        {
            "base64" => Base64(raw),
            "quoted-printable" => QuotedPrintable(raw),
            _ => raw.ToArray(),
        };
        string? loc = headers.GetValueOrDefault("Content-Location");
        if (loc != null) loc = DecodeWords(loc).Trim();
        string? cid = headers.GetValueOrDefault("Content-ID")?.Trim().Trim('<', '>');
        Parts.Add(new Part(type, ps.GetValueOrDefault("charset"), string.IsNullOrEmpty(loc) ? null : loc, string.IsNullOrEmpty(cid) ? null : cid, body));
    }

    /// <summary>Части multipart: границы «--boundary» в начале строки, завершение «--boundary--».</summary>
    static List<(int S, int E)> SplitMultipart(byte[] d, int s, int e, string boundary)
    {
        var res = new List<(int, int)>();
        var marker = Encoding.Latin1.GetBytes("--" + boundary);
        int partStart = -1;
        int p = s;
        while (p < e)
        {
            int i = d.AsSpan(p, e - p).IndexOf(marker);
            if (i < 0) break;
            i += p;
            bool lineStart = i == s || d[i - 1] == '\n';
            if (!lineStart) { p = i + marker.Length; continue; }
            if (partStart >= 0)
            {
                // Перевод строки перед границей принадлежит ей
                int pe = i;
                if (pe > partStart && d[pe - 1] == '\n') pe--;
                if (pe > partStart && d[pe - 1] == '\r') pe--;
                res.Add((partStart, pe));
            }
            int after = i + marker.Length;
            if (after + 1 < e && d[after] == '-' && d[after + 1] == '-') return res;
            int nl = Array.IndexOf(d, (byte)'\n', after, e - after);
            if (nl < 0) return res;
            partStart = nl + 1;
            p = partStart;
        }
        if (partStart >= 0 && partStart < e) res.Add((partStart, e));   // нет завершающей границы
        return res;
    }

    static byte[] Base64(ReadOnlySpan<byte> raw)
    {
        var sb = new StringBuilder(raw.Length);
        foreach (byte b in raw)
            if (b is >= (byte)'A' and <= (byte)'Z' or >= (byte)'a' and <= (byte)'z' or >= (byte)'0' and <= (byte)'9' or (byte)'+' or (byte)'/') sb.Append((char)b);
            else if (b is (byte)'-') sb.Append('+');
            else if (b is (byte)'_') sb.Append('/');
        while (sb.Length % 4 != 0) sb.Append('=');
        if (sb.Length % 4 == 0 && sb.Length >= 4 && sb[^1] == '=' && sb[^2] == '=' && sb[^3] == '=') sb.Length -= 4;   // остаток в 1 символ
        try { return Convert.FromBase64String(sb.ToString()); } catch { return []; }
    }

    static int Hex(byte c) => c switch { >= (byte)'0' and <= (byte)'9' => c - '0', >= (byte)'A' and <= (byte)'F' => c - 'A' + 10, >= (byte)'a' and <= (byte)'f' => c - 'a' + 10, _ => -1 };

    static byte[] QuotedPrintable(ReadOnlySpan<byte> raw)
    {
        var m = new MemoryStream(raw.Length);
        for (int i = 0; i < raw.Length; i++)
        {
            byte b = raw[i];
            if (b != '=') { m.WriteByte(b); continue; }
            // Мягкий перенос: «=» в конце строки (возможно, с пробелами)
            int j = i + 1;
            while (j < raw.Length && raw[j] is (byte)' ' or (byte)'\t') j++;
            if (j < raw.Length && raw[j] == '\r') j++;
            if (j < raw.Length && raw[j] == '\n') { i = j; continue; }
            if (j >= raw.Length) break;
            if (i + 2 < raw.Length && Hex(raw[i + 1]) >= 0 && Hex(raw[i + 2]) >= 0)
            {
                m.WriteByte((byte)(Hex(raw[i + 1]) << 4 | Hex(raw[i + 2])));
                i += 2;
            }
            else m.WriteByte(b);
        }
        return m.ToArray();
    }

    /// <summary>Закодированные слова RFC 2047 (=?charset?B|Q?...?=) в заголовках.</summary>
    static string DecodeWords(string v)
    {
        if (!v.Contains("=?")) return v;
        return System.Text.RegularExpressions.Regex.Replace(v, @"=\?([^?]+)\?([BbQq])\?([^?]*)\?=\s*", m =>
        {
            try
            {
                var enc = Encoding.GetEncoding(m.Groups[1].Value);
                byte[] bytes = m.Groups[2].Value is "B" or "b"
                    ? Convert.FromBase64String(m.Groups[3].Value)
                    : QuotedPrintable(Encoding.Latin1.GetBytes(m.Groups[3].Value.Replace('_', ' ')));
                return enc.GetString(bytes);
            }
            catch { return m.Value; }
        });
    }

    // ---------------- Поиск ресурсов ----------------

    /// <summary>Варианты написания адреса для поиска: как есть, без фрагмента, с раскодированными %XX, нормализованный URI.</summary>
    static IEnumerable<string> LocationKeys(string loc)
    {
        yield return loc;
        int h = loc.IndexOf('#');
        string noFrag = h >= 0 ? loc[..h] : loc;
        yield return noFrag;
        string unesc;
        try { unesc = Uri.UnescapeDataString(noFrag); } catch { unesc = noFrag; }
        yield return unesc;
        yield return unesc.Replace('\\', '/');
        if (Uri.TryCreate(noFrag, UriKind.Absolute, out var u))
        {
            yield return u.AbsoluteUri;
            yield return Uri.UnescapeDataString(u.AbsoluteUri);
        }
    }

    /// <summary>
    /// Ресурс архива по ссылке из страницы: cid:, абсолютный URL или относительный путь
    /// (разрешается относительно <paramref name="baseUrl"/>), в крайнем случае — по имени файла.
    /// </summary>
    public Part? Find(string href, string? baseUrl)
    {
        href = href.Trim();
        if (href.Length == 0) return null;
        if (href.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
        {
            string id = Uri.UnescapeDataString(href[4..]).Trim('<', '>');
            // Content-ID, либо Content-Location вида «cid:…» (так Chromium сохраняет встроенные стили)
            return _byCid.GetValueOrDefault(id) ?? _byLocationCi.GetValueOrDefault(href) ?? _byLocationCi.GetValueOrDefault("cid:" + id);
        }
        var candidates = new List<string> { href };
        if (baseUrl != null && Uri.TryCreate(baseUrl, UriKind.Absolute, out var bu) && Uri.TryCreate(bu, href, out var abs))
            candidates.Insert(0, abs.AbsoluteUri);
        if (href.StartsWith("//") && baseUrl != null && Uri.TryCreate(baseUrl, UriKind.Absolute, out var bu2))
            candidates.Add(bu2.Scheme + ":" + href);
        foreach (var c in candidates)
            foreach (var k in LocationKeys(c))
                if (_byLocation.TryGetValue(k, out var p) || _byLocationCi.TryGetValue(k, out p)) return p;
        // По имени файла (архивы, где адреса записаны иначе, чем в разметке)
        string name = FileName(href);
        if (name.Length > 0)
        {
            var same = Parts.Where(p => p.Location != null && FileName(p.Location).Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();
            if (same.Count == 1) return same[0];
        }
        return null;
    }

    static string FileName(string url)
    {
        string s = url.Split('?', '#')[0].Replace('\\', '/');
        try { s = Uri.UnescapeDataString(s); } catch { }
        int i = s.LastIndexOf('/');
        return i >= 0 ? s[(i + 1)..] : s;
    }
}
