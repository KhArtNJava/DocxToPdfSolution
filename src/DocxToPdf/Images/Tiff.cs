using System.Buffers.Binary;
using System.IO.Compression;

namespace DocxToPdf.Images;

/// <summary>
/// Декодер TIFF (первая страница): несжатый, PackBits, LZW (с предиктором), Deflate, JPEG (новый, одна полоса),
/// CCITT RLE / Group 3 (1D и 2D) / Group 4; полосы и плитки; RGB(A), оттенки серого, палитра, CMYK, 1–16 бит.
/// </summary>
static class Tiff
{
    public static bool IsTiff(byte[] d) =>
        d.Length > 8 && ((d[0] == 'I' && d[1] == 'I' && d[2] == 42 && d[3] == 0) || (d[0] == 'M' && d[1] == 'M' && d[2] == 0 && d[3] == 42));

    sealed class Reader(byte[] d, bool le)
    {
        public readonly byte[] D = d;
        public int U16(int o) => le ? BinaryPrimitives.ReadUInt16LittleEndian(D.AsSpan(o)) : BinaryPrimitives.ReadUInt16BigEndian(D.AsSpan(o));
        public uint U32(int o) => le ? BinaryPrimitives.ReadUInt32LittleEndian(D.AsSpan(o)) : BinaryPrimitives.ReadUInt32BigEndian(D.AsSpan(o));
        public bool LittleEndian => le;
    }

    public static PdfImageData? Decode(byte[] d)
    {
        var r = new Reader(d, d[0] == 'I');
        int ifd = (int)r.U32(4);
        if (ifd <= 0 || ifd + 2 > d.Length) return null;
        int n = r.U16(ifd);
        var tags = new Dictionary<int, uint[]>();
        for (int i = 0; i < n; i++)
        {
            int e = ifd + 2 + i * 12;
            if (e + 12 > d.Length) break;
            int tag = r.U16(e), type = r.U16(e + 2);
            uint count = r.U32(e + 4);
            int size = type switch { 1 or 2 or 6 or 7 => 1, 3 or 8 => 2, 4 or 9 or 11 => 4, 5 or 10 or 12 => 8, _ => 1 };
            if (count > 1 << 24) continue;
            int off = size * count <= 4 ? e + 8 : (int)r.U32(e + 8);
            var vals = new uint[count];
            for (int k = 0; k < count; k++)
            {
                int p = off + k * size;
                if (p + size > d.Length) break;
                vals[k] = size switch { 1 => d[p], 2 => (uint)r.U16(p), 4 => r.U32(p), _ => r.U32(p) };
            }
            tags[tag] = vals;
        }
        uint T(int tag, uint def = 0) => tags.TryGetValue(tag, out var v) && v.Length > 0 ? v[0] : def;

        int w = (int)T(256), h = (int)T(257);
        if (w <= 0 || h <= 0 || (long)w * h > 200_000_000) return null;
        int spp = (int)T(277, 1);
        int bps = (int)T(258, 1);
        int comp = (int)T(259, 1), photo = (int)T(262, spp >= 3 ? 2u : 1u), planar = (int)T(284, 1), predictor = (int)T(317, 1);
        bool fillLsb = T(266, 1) == 2;
        int extra = tags.TryGetValue(338, out var ex) ? ex.Length : 0;
        int alphaKind = extra > 0 ? (int)ex![0] : 0;   // 1 — премультиплицированная, 2 — обычная
        bool tiled = tags.ContainsKey(322);
        int tw = tiled ? (int)T(322) : w, th = tiled ? (int)T(323) : (int)Math.Min((uint)h, T(278, (uint)h));
        if (tw <= 0 || th <= 0) return null;
        var offsets = tags.GetValueOrDefault(tiled ? 324 : 273) ?? [];
        var counts = tags.GetValueOrDefault(tiled ? 325 : 279) ?? [];
        if (offsets.Length == 0) return null;

        // JPEG (новый стиль): полосы склеиваются в один поток JPEG без перекодирования и передаются в PDF как есть
        if (comp == 7 && !tiled)
        {
            var jt = tags.TryGetValue(347, out var jtv) && jtv.Length > 4 ? jtv.Select(b => (byte)b).ToArray() : null;
            var strips = new List<byte[]>();
            for (int i = 0; i < offsets.Length; i++)
            {
                int o = (int)offsets[i], len = i < counts.Length ? (int)counts[i] : d.Length - o;
                if (o < 0 || o >= d.Length) return null;
                var strip = d.AsSpan(o, Math.Min(len, d.Length - o)).ToArray();
                if (jt != null && strip.Length > 2) strip = [.. jt.AsSpan(0, jt.Length - 2), .. strip.AsSpan(2)];
                strips.Add(strip);
            }
            var joined = strips.Count == 1 ? strips[0] : JoinJpegStrips(strips, w, h, th);
            return joined != null ? ImageDecoder.Decode(joined) : null;
        }
        if (comp is not (1 or 2 or 3 or 4 or 5 or 8 or 32946 or 32773)) return null;

        int planes = planar == 2 ? spp : 1, sppChunk = planar == 2 ? 1 : spp;
        int rowBytes = (w * sppChunk * bps + 7) / 8;
        var planeData = new byte[planes][];
        for (int pl = 0; pl < planes; pl++) planeData[pl] = new byte[rowBytes * h];
        int across = (w + tw - 1) / tw, down = (h + th - 1) / th;
        int perPlane = across * down;
        for (int idx = 0; idx < offsets.Length; idx++)
        {
            int pl = idx / perPlane, cell = idx % perPlane;
            if (pl >= planes) break;
            int o = (int)offsets[idx];
            int len = idx < counts.Length ? (int)counts[idx] : d.Length - o;
            if (o < 0 || o >= d.Length) continue;
            len = Math.Min(len, d.Length - o);
            var src = d.AsSpan(o, len).ToArray();
            if (fillLsb) for (int i = 0; i < src.Length; i++) src[i] = Reverse(src[i]);
            int cx = cell % across, cy = cell / across;
            int cw = tiled ? tw : w, ch = tiled ? th : Math.Min(th, h - cy * th);
            int cRow = (cw * sppChunk * bps + 7) / 8;
            byte[] raw = comp switch
            {
                1 => src,
                32773 => PackBits(src, cRow * ch),
                5 => Lzw(src, cRow * ch),
                8 or 32946 => Inflate(src),
                2 => Ccitt.Decode(src, cw, ch, Ccitt.Mode.Rle, false, false),
                3 => Ccitt.Decode(src, cw, ch, Ccitt.Mode.G3, (T(292) & 1) != 0, true),
                4 => Ccitt.Decode(src, cw, ch, Ccitt.Mode.G4, false, false),
                _ => src
            };
            if (predictor == 2) UnPredict(raw, cRow, ch, sppChunk, bps, r.LittleEndian);
            // Копирование полосы/плитки в общий буфер плоскости
            int dstX = cx * tw, dstY = cy * th;
            int copyRows = Math.Min(ch, h - dstY);
            int bitX = dstX * sppChunk * bps;
            int copyBits = Math.Min(cw, w - dstX) * sppChunk * bps;
            for (int y = 0; y < copyRows; y++)
            {
                int srcOff = y * cRow, dstOff = (dstY + y) * rowBytes + bitX / 8;
                if (srcOff >= raw.Length) break;
                int nBytes = Math.Min((copyBits + 7) / 8, Math.Min(raw.Length - srcOff, planeData[pl].Length - dstOff));
                if (nBytes > 0) Array.Copy(raw, srcOff, planeData[pl], dstOff, nBytes);
            }
        }

        // Выборка: значение отсчёта s пикселя x строки y (масштаб 0…255)
        int maxV = (1 << Math.Min(bps, 16)) - 1;
        int Sample(int y, int x, int s)
        {
            byte[] buf; int sIdx, ss;
            if (planar == 2) { buf = planeData[s]; sIdx = x; ss = 1; }
            else { buf = planeData[0]; sIdx = x * spp + s; ss = spp; }
            _ = ss;
            int row = y * rowBytes;
            switch (bps)
            {
                case 8: return buf[row + sIdx];
                case 16:
                {
                    int p = row + sIdx * 2;
                    return r.LittleEndian ? buf[p + 1] : buf[p];
                }
                default:
                {
                    int bit = sIdx * bps;
                    int b = buf[row + (bit >> 3)];
                    int v = (b >> (8 - bps - (bit & 7))) & ((1 << bps) - 1);
                    return v * 255 / maxV;
                }
            }
        }
        int RawSample(int y, int x)
        {
            // Индекс палитры (без масштабирования)
            int row = y * rowBytes;
            var buf = planeData[0];
            if (bps == 8) return buf[row + x * spp];
            if (bps == 16) { int p = row + x * spp * 2; return r.LittleEndian ? buf[p] | buf[p + 1] << 8 : buf[p] << 8 | buf[p + 1]; }
            int bit = x * spp * bps;
            return (buf[row + (bit >> 3)] >> (8 - bps - (bit & 7))) & ((1 << bps) - 1);
        }

        var rgb = new byte[w * h * 3];
        byte[]? alpha = null;
        var cmap = tags.GetValueOrDefault(320);
        int colorSamples = photo switch { 2 => 3, 5 => 4, 6 => 3, _ => 1 };
        if (extra > 0 && spp > colorSamples) alpha = new byte[w * h];
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++)
            {
                int o = (y * w + x) * 3;
                int rr, gg, bb;
                switch (photo)
                {
                    case 0: rr = gg = bb = 255 - Sample(y, x, 0); break;
                    case 1: rr = gg = bb = Sample(y, x, 0); break;
                    case 3:
                    {
                        int i = RawSample(y, x), nc = 1 << bps;
                        if (cmap != null && cmap.Length >= nc * 3 && i < nc)
                        { rr = (int)(cmap[i] >> 8); gg = (int)(cmap[nc + i] >> 8); bb = (int)(cmap[2 * nc + i] >> 8); }
                        else rr = gg = bb = i * 255 / Math.Max(1, nc - 1);
                        break;
                    }
                    case 5:
                    {
                        int c = Sample(y, x, 0), m = Sample(y, x, 1), yy = Sample(y, x, 2), k = Sample(y, x, 3);
                        rr = (255 - c) * (255 - k) / 255; gg = (255 - m) * (255 - k) / 255; bb = (255 - yy) * (255 - k) / 255;
                        break;
                    }
                    case 6:
                    {
                        double Y = Sample(y, x, 0), cb = Sample(y, x, 1) - 128, cr = Sample(y, x, 2) - 128;
                        rr = Clamp(Y + 1.402 * cr); gg = Clamp(Y - 0.344136 * cb - 0.714136 * cr); bb = Clamp(Y + 1.772 * cb);
                        break;
                    }
                    default: rr = Sample(y, x, 0); gg = Sample(y, x, Math.Min(1, spp - 1)); bb = Sample(y, x, Math.Min(2, spp - 1)); break;
                }
                if (alpha != null)
                {
                    int a = Sample(y, x, colorSamples);
                    alpha[y * w + x] = (byte)a;
                    if (alphaKind == 1 && a > 0 && a < 255)
                    { rr = Math.Min(255, rr * 255 / a); gg = Math.Min(255, gg * 255 / a); bb = Math.Min(255, bb * 255 / a); }
                }
                rgb[o] = (byte)rr; rgb[o + 1] = (byte)gg; rgb[o + 2] = (byte)bb;
            }
        if (alpha != null && alpha.All(a => a == 255)) alpha = null;
        return new PdfImageData { Width = w, Height = h, Data = rgb, Alpha = alpha };
    }

    /// <summary>
    /// Склейка полос JPEG одного изображения: заголовок первой полосы (с полной высотой в SOF и интервалом
    /// перезапуска DRI = одна полоса), затем сжатые данные полос, разделённые маркерами RSTn.
    /// </summary>
    static byte[]? JoinJpegStrips(List<byte[]> strips, int width, int height, int rowsPerStrip)
    {
        // Разбор первой полосы: до конца заголовка SOS
        var first = strips[0];
        int pos = 2, sofPos = -1, sosEnd = -1, hMax = 1, vMax = 1;
        while (pos + 4 <= first.Length)
        {
            if (first[pos] != 0xFF) return null;
            int marker = first[pos + 1];
            int segLen = first[pos + 2] << 8 | first[pos + 3];
            if (marker is 0xC0 or 0xC1)
            {
                sofPos = pos;
                int nc = first[pos + 9];
                for (int i = 0; i < nc; i++)
                {
                    int sf = first[pos + 11 + i * 3];
                    hMax = Math.Max(hMax, sf >> 4); vMax = Math.Max(vMax, sf & 15);
                }
            }
            else if (marker == 0xDD) return null;   // свой интервал перезапуска — не склеиваем
            else if (marker == 0xDA) { sosEnd = pos + 2 + segLen; break; }
            else if (marker is 0xC2 or 0xC3 or (>= 0xC5 and <= 0xCF and not 0xC8 and not 0xCC)) return null;   // не базовый JPEG
            pos += 2 + segLen;
        }
        if (sofPos < 0 || sosEnd < 0) return null;
        int mcuW = 8 * hMax, mcuH = 8 * vMax;
        if (rowsPerStrip % mcuH != 0) return null;
        int interval = (width + mcuW - 1) / mcuW * (rowsPerStrip / mcuH);
        if (interval <= 0 || interval > 65535) return null;

        var o = new MemoryStream();
        // Заголовок до SOS, высота в SOF — полная, перед SOS — DRI
        var head = first.AsSpan(0, sosEnd).ToArray();
        head[sofPos + 5] = (byte)(height >> 8); head[sofPos + 6] = (byte)height;
        int sosPos = FindSos(head);
        o.Write(head, 0, sosPos);
        o.Write([0xFF, 0xDD, 0x00, 0x04, (byte)(interval >> 8), (byte)interval]);
        o.Write(head, sosPos, head.Length - sosPos);
        for (int i = 0; i < strips.Count; i++)
        {
            var st = strips[i];
            int start = i == 0 ? sosEnd : EntropyStart(st);
            if (start < 0) return null;
            int end = st.Length;
            // Конец данных — маркер EOI (FFD9) в конце полосы
            while (end > start + 1 && !(st[end - 2] == 0xFF && st[end - 1] == 0xD9)) end--;
            if (end > start + 1) end -= 2; else end = st.Length;
            o.Write(st, start, end - start);
            if (i < strips.Count - 1) o.Write([0xFF, (byte)(0xD0 + i % 8)]);
        }
        o.Write([0xFF, 0xD9]);
        return o.ToArray();

        static int FindSos(byte[] b) { int p = 2; while (p + 4 <= b.Length) { if (b[p + 1] == 0xDA) return p; p += 2 + (b[p + 2] << 8 | b[p + 3]); } return -1; }
        static int EntropyStart(byte[] b)
        {
            int p = 2;
            while (p + 4 <= b.Length)
            {
                if (b[p] != 0xFF) return -1;
                int seg = b[p + 2] << 8 | b[p + 3];
                if (b[p + 1] == 0xDD) return -1;
                if (b[p + 1] == 0xDA) return p + 2 + seg;
                p += 2 + seg;
            }
            return -1;
        }
    }

    static int Clamp(double v) => (int)Math.Clamp(Math.Round(v), 0, 255);

    static byte Reverse(byte b)
    {
        b = (byte)((b & 0xF0) >> 4 | (b & 0x0F) << 4);
        b = (byte)((b & 0xCC) >> 2 | (b & 0x33) << 2);
        return (byte)((b & 0xAA) >> 1 | (b & 0x55) << 1);
    }

    static void UnPredict(byte[] raw, int rowBytes, int rows, int spp, int bps, bool le)
    {
        for (int y = 0; y < rows; y++)
        {
            int row = y * rowBytes;
            if (bps == 8)
                for (int i = spp; i < rowBytes && row + i < raw.Length; i++) raw[row + i] += raw[row + i - spp];
            else if (bps == 16)
                for (int i = spp * 2; i + 1 < rowBytes && row + i + 1 < raw.Length; i += 2)
                {
                    int p = row + i, q = p - spp * 2;
                    int v = le ? raw[p] | raw[p + 1] << 8 : raw[p] << 8 | raw[p + 1];
                    int u = le ? raw[q] | raw[q + 1] << 8 : raw[q] << 8 | raw[q + 1];
                    v = (v + u) & 0xFFFF;
                    if (le) { raw[p] = (byte)v; raw[p + 1] = (byte)(v >> 8); } else { raw[p] = (byte)(v >> 8); raw[p + 1] = (byte)v; }
                }
        }
    }

    static byte[] PackBits(byte[] s, int expected)
    {
        var o = new List<byte>(expected);
        int i = 0;
        while (i < s.Length && o.Count < expected)
        {
            sbyte n = (sbyte)s[i++];
            if (n >= 0) { for (int k = 0; k <= n && i < s.Length; k++) o.Add(s[i++]); }
            else if (n != -128 && i < s.Length) { byte b = s[i++]; for (int k = 0; k < 1 - n; k++) o.Add(b); }
        }
        return [.. o];
    }

    static byte[] Inflate(byte[] s)
    {
        try
        {
            using var z = new ZLibStream(new MemoryStream(s), CompressionMode.Decompress);
            using var ms = new MemoryStream();
            z.CopyTo(ms);
            return ms.ToArray();
        }
        catch
        {
            using var z = new DeflateStream(new MemoryStream(s), CompressionMode.Decompress);
            using var ms = new MemoryStream();
            try { z.CopyTo(ms); } catch { }
            return ms.ToArray();
        }
    }

    /// <summary>LZW TIFF: коды 9–12 бит, старший бит первым, «раннее» увеличение ширины.</summary>
    static byte[] Lzw(byte[] s, int expected)
    {
        var o = new List<byte>(expected);
        var table = new List<byte[]>(4096);
        void Reset() { table.Clear(); for (int i = 0; i < 256; i++) table.Add([(byte)i]); table.Add([]); table.Add([]); }
        Reset();
        int bitPos = 0, width = 9;
        byte[]? prev = null;
        int Next()
        {
            if (bitPos + width > s.Length * 8) return 257;
            int v = 0;
            for (int k = 0; k < width; k++)
            {
                int b = bitPos + k;
                v = v << 1 | (s[b >> 3] >> (7 - (b & 7)) & 1);
            }
            bitPos += width;
            return v;
        }
        while (true)
        {
            int code = Next();
            if (code == 257) break;
            if (code == 256) { Reset(); width = 9; prev = null; continue; }
            byte[] entry;
            if (code < table.Count) entry = table[code];
            else if (prev != null && code == table.Count) entry = [.. prev, prev[0]];
            else break;
            o.AddRange(entry);
            if (prev != null && table.Count < 4096) table.Add([.. prev, entry[0]]);
            prev = entry;
            if (table.Count + 1 >= 1 << width && width < 12) width++;
            if (o.Count >= expected && expected > 0) break;
        }
        return [.. o];
    }
}

/// <summary>Факсовое сжатие CCITT: Modified Huffman (RLE), T.4 (Group 3, 1D/2D), T.6 (Group 4). Белый = 0, чёрный = 1.</summary>
static class Ccitt
{
    public enum Mode { Rle, G3, G4 }

    static readonly string[] WhiteTerm =
    [
        "00110101", "000111", "0111", "1000", "1011", "1100", "1110", "1111", "10011", "10100", "00111", "01000", "001000", "000011", "110100", "110101",
        "101010", "101011", "0100111", "0001100", "0001000", "0010111", "0000011", "0000100", "0101000", "0101011", "0010011", "0100100", "0011000", "00000010", "00000011", "00011010",
        "00011011", "00010010", "00010011", "00010100", "00010101", "00010110", "00010111", "00101000", "00101001", "00101010", "00101011", "00101100", "00101101", "00000100", "00000101", "00001010",
        "00001011", "01010010", "01010011", "01010100", "01010101", "00100100", "00100101", "01011000", "01011001", "01011010", "01011011", "01001010", "01001011", "00110010", "00110011", "00110100",
    ];
    static readonly string[] WhiteMakeup =
    [
        "11011", "10010", "010111", "0110111", "00110110", "00110111", "01100100", "01100101", "01101000", "01100111", "011001100", "011001101", "011010010", "011010011",
        "011010100", "011010101", "011010110", "011010111", "011011000", "011011001", "011011010", "011011011", "010011000", "010011001", "010011010", "011000", "010011011",
    ];
    static readonly string[] BlackTerm =
    [
        "0000110111", "010", "11", "10", "011", "0011", "0010", "00011", "000101", "000100", "0000100", "0000101", "0000111", "00000100", "00000111", "000011000",
        "0000010111", "0000011000", "0000001000", "00001100111", "00001101000", "00001101100", "00000110111", "00000101000", "00000010111", "00000011000", "000011001010", "000011001011", "000011001100", "000011001101", "000001101000", "000001101001",
        "000001101010", "000001101011", "000011010010", "000011010011", "000011010100", "000011010101", "000011010110", "000011010111", "000001101100", "000001101101", "000011011010", "000011011011", "000001010100", "000001010101", "000001010110", "000001010111",
        "000001100100", "000001100101", "000001010010", "000001010011", "000000100100", "000000110111", "000000111000", "000000100111", "000000101000", "000001011000", "000001011001", "000000101011", "000000101100", "000001011010", "000001100110", "000001100111",
    ];
    static readonly string[] BlackMakeup =
    [
        "0000001111", "000011001000", "000011001001", "000001011011", "000000110011", "000000110100", "000000110101", "0000001101100", "0000001101101", "0000001001010", "0000001001011", "0000001001100", "0000001001101", "0000001110010",
        "0000001110011", "0000001110100", "0000001110101", "0000001110110", "0000001110111", "0000001010010", "0000001010011", "0000001010100", "0000001010101", "0000001011010", "0000001011011", "0000001100100", "0000001100101",
    ];
    static readonly string[] ExtMakeup =
    [
        "00000001000", "00000001100", "00000001101", "000000010010", "000000010011", "000000010100", "000000010101", "000000010110", "000000010111", "000000011100", "000000011101", "000000011110", "000000011111",
    ];

    // Таблицы кодов: (длина, код) → длина серии
    static readonly Dictionary<(int, int), int> White = Build(WhiteTerm, WhiteMakeup), Black = Build(BlackTerm, BlackMakeup);

    static Dictionary<(int, int), int> Build(string[] term, string[] makeup)
    {
        var d = new Dictionary<(int, int), int>();
        for (int i = 0; i < term.Length; i++) d[(term[i].Length, Convert.ToInt32(term[i], 2))] = i;
        for (int i = 0; i < makeup.Length; i++) d[(makeup[i].Length, Convert.ToInt32(makeup[i], 2))] = (i + 1) * 64;
        for (int i = 0; i < ExtMakeup.Length; i++) d[(ExtMakeup[i].Length, Convert.ToInt32(ExtMakeup[i], 2))] = 1792 + i * 64;
        return d;
    }

    sealed class Bits(byte[] d)
    {
        public int Pos;
        public bool End => Pos >= d.Length * 8;
        public int Bit() { if (End) { Pos++; return 0; } int b = d[Pos >> 3] >> (7 - (Pos & 7)) & 1; Pos++; return b; }
        public int Peek(int n) { int v = 0; for (int i = 0; i < n; i++) { int p = Pos + i; v = v << 1 | (p < d.Length * 8 ? d[p >> 3] >> (7 - (p & 7)) & 1 : 0); } return v; }
        public void Align() => Pos = (Pos + 7) & ~7;
    }

    /// <summary>Длина серии одного цвета (с учётом кодов-дополнений).</summary>
    static int Run(Bits b, bool black)
    {
        var table = black ? Black : White;
        int total = 0;
        while (true)
        {
            int code = 0, len = 0, run = -1;
            while (len < 14)
            {
                code = code << 1 | b.Bit();
                len++;
                if (table.TryGetValue((len, code), out run)) break;
                run = -1;
            }
            if (run < 0) return -1;
            total += run;
            if (run < 64) return total;
        }
    }

    public static byte[] Decode(byte[] data, int width, int height, Mode mode, bool g3TwoD, bool g3)
    {
        int rowBytes = (width + 7) / 8;
        var outp = new byte[rowBytes * height];
        var b = new Bits(data);
        var refLine = new List<int> { width, width };
        for (int y = 0; y < height && !b.End; y++)
        {
            bool twoD = mode == Mode.G4;
            if (mode == Mode.G3)
            {
                // Необязательный EOL (с битами заполнения) перед строкой
                int zeros = 0;
                while (zeros < 400 && b.Peek(1) == 0 && !b.End) { b.Bit(); zeros++; }
                if (zeros >= 11) b.Bit(); else b.Pos -= zeros;
                if (g3TwoD) twoD = b.Bit() == 0;
            }
            var cur = new List<int>();
            bool ok = twoD ? Decode2D(b, width, refLine, cur) : Decode1D(b, width, cur);
            if (!ok) break;
            // Строка: чётные изменения — начало чёрной серии
            int row = y * rowBytes;
            for (int i = 0; i + 1 < cur.Count; i += 2)
            {
                int s = Math.Max(0, cur[i]), e = Math.Min(width, cur[i + 1]);
                for (int x = s; x < e; x++) outp[row + (x >> 3)] |= (byte)(0x80 >> (x & 7));
            }
            if (cur.Count % 2 == 1) for (int x = Math.Max(0, cur[^1]); x < width; x++) outp[row + (x >> 3)] |= (byte)(0x80 >> (x & 7));
            refLine = [.. cur.Where(c => c < width), width, width];
            if (mode == Mode.Rle) b.Align();
        }
        return outp;
    }

    static bool Decode1D(Bits b, int width, List<int> cur)
    {
        int a0 = 0;
        bool black = false;
        while (a0 < width)
        {
            int run = Run(b, black);
            if (run < 0) return cur.Count > 0 || a0 > 0;
            a0 += run;
            cur.Add(Math.Min(a0, width));
            black = !black;
        }
        return true;
    }

    static bool Decode2D(Bits b, int width, List<int> refLine, List<int> cur)
    {
        int a0 = -1;
        bool black = false;
        int guard = 0;
        while (a0 < width && guard++ < width * 4 + 16)
        {
            // b1 — первое изменение опорной строки правее a0 с цветом, противоположным текущему; b2 — следующее
            int i = 0;
            while (i < refLine.Count && (refLine[i] <= a0 || (i % 2 == 1) != black)) i++;
            int b1 = i < refLine.Count ? refLine[i] : width, b2 = i + 1 < refLine.Count ? refLine[i + 1] : width;
            if (b.Peek(1) == 1) { b.Bit(); V(0); continue; }
            int c3 = b.Peek(3);
            if (c3 == 0b011) { b.Pos += 3; V(1); continue; }
            if (c3 == 0b010) { b.Pos += 3; V(-1); continue; }
            if (c3 == 0b001)
            {
                b.Pos += 3;
                int start = Math.Max(a0, 0);
                int r1 = Run(b, black), r2 = Run(b, !black);
                if (r1 < 0 || r2 < 0) return false;
                int a1 = start + r1, a2 = a1 + r2;
                cur.Add(Math.Min(a1, width)); cur.Add(Math.Min(a2, width));
                a0 = a2;
                continue;
            }
            int c4 = b.Peek(4);
            if (c4 == 0b0001) { b.Pos += 4; a0 = b2; continue; }   // проход
            int c6 = b.Peek(6);
            if (c6 == 0b000011) { b.Pos += 6; V(2); continue; }
            if (c6 == 0b000010) { b.Pos += 6; V(-2); continue; }
            int c7 = b.Peek(7);
            if (c7 == 0b0000011) { b.Pos += 7; V(3); continue; }
            if (c7 == 0b0000010) { b.Pos += 7; V(-3); continue; }
            return cur.Count > 0;   // EOL/EOFB или неизвестный код

            void V(int dlt)
            {
                int a1 = Math.Clamp(b1 + dlt, 0, width);
                cur.Add(a1);
                a0 = a1;
                black = !black;
            }
        }
        return true;
    }
}
