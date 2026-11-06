using System.Buffers.Binary;
using System.IO.Compression;

namespace DocxToPdf.Images;

/// <summary>Изображение в виде, пригодном для записи в PDF.</summary>
public sealed class PdfImageData
{
    public int Width, Height;
    public string ColorSpace = "/DeviceRGB";
    public int BitsPerComponent = 8;
    public string Filter = "/FlateDecode";     // /DCTDecode для JPEG — данные как есть
    public byte[] Data = [];                   // для Flate — несжатые пиксели
    public byte[]? Alpha;                      // несжатый канал прозрачности (8 бит)
    public string? Decode;
    /// <summary>Векторный метафайл (EMF/WMF): Data пусто, размер — собственный размер в пикселях 96 dpi.</summary>
    public bool Metafile;
}

/// <summary>Декодеры JPEG (только заголовок), PNG, GIF и BMP без сторонних библиотек.</summary>
public static class ImageDecoder
{
    public static PdfImageData? Decode(byte[] d)
    {
        try
        {
            if (d.Length > 3 && d[0] == 0xFF && d[1] == 0xD8) return Jpeg(d);
            if (d.Length > 8 && d[0] == 0x89 && d[1] == 'P' && d[2] == 'N' && d[3] == 'G') return Png(d);
            if (d.Length > 6 && d[0] == 'G' && d[1] == 'I' && d[2] == 'F') return Gif(d);
            if (d.Length > 26 && d[0] == 'B' && d[1] == 'M') return Bmp(d);
            if (Tiff.IsTiff(d)) return Tiff.Decode(d);
            if (DocxToPdf.Pdf.Metafile.Size(d) is { } ms)
                return new PdfImageData { Width = Math.Max(1, (int)Math.Round(ms.W / 0.75)), Height = Math.Max(1, (int)Math.Round(ms.H / 0.75)), Metafile = true };
            if (DocxToPdf.Pdf.Svg.IsSvg(d) && DocxToPdf.Pdf.Svg.Size(d) is { } ss)
                return new PdfImageData { Width = Math.Max(1, (int)Math.Round(ss.W / 0.75)), Height = Math.Max(1, (int)Math.Round(ss.H / 0.75)), Metafile = true };
        }
        catch { /* битое изображение */ }
        return null;
    }

    // ------------------------------------------------------------ JPEG
    static PdfImageData? Jpeg(byte[] d)
    {
        int p = 2;
        bool adobe = false;
        while (p + 4 < d.Length)
        {
            if (d[p] != 0xFF) { p++; continue; }
            int marker = d[p + 1];
            if (marker == 0xFF) { p++; continue; }
            if (marker is 0xD8 or 0x01 || marker is >= 0xD0 and <= 0xD7) { p += 2; continue; }
            int len = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(p + 2));
            if (marker == 0xEE && len >= 12 && d[p + 4] == 'A' && d[p + 5] == 'd') adobe = true;
            if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
            {
                int h = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(p + 5));
                int w = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(p + 7));
                int comps = d[p + 9];
                var img = new PdfImageData
                {
                    Width = w, Height = h, Filter = "/DCTDecode", Data = d,
                    ColorSpace = comps switch { 1 => "/DeviceGray", 4 => "/DeviceCMYK", _ => "/DeviceRGB" }
                };
                if (comps == 4 && adobe) img.Decode = "[1 0 1 0 1 0 1 0]";
                return img;
            }
            p += 2 + len;
        }
        return null;
    }

    // ------------------------------------------------------------ PNG
    static PdfImageData? Png(byte[] d)
    {
        int p = 8;
        int w = 0, h = 0, depth = 8, ctype = 0, interlace = 0;
        byte[]? palette = null, trns = null;
        using var idat = new MemoryStream();
        while (p + 8 <= d.Length)
        {
            int len = (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(p));
            string type = System.Text.Encoding.ASCII.GetString(d, p + 4, 4);
            int data = p + 8;
            if (data + len > d.Length) break;
            switch (type)
            {
                case "IHDR":
                    w = (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(data));
                    h = (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(data + 4));
                    depth = d[data + 8]; ctype = d[data + 9]; interlace = d[data + 12];
                    break;
                case "PLTE": palette = d.AsSpan(data, len).ToArray(); break;
                case "tRNS": trns = d.AsSpan(data, len).ToArray(); break;
                case "IDAT": idat.Write(d, data, len); break;
            }
            if (type == "IEND") break;
            p = data + len + 4;
        }
        if (w <= 0 || h <= 0) return null;

        byte[] raw;
        idat.Position = 0;
        using (var z = new ZLibStream(idat, CompressionMode.Decompress))
        using (var outMs = new MemoryStream())
        {
            z.CopyTo(outMs);
            raw = outMs.ToArray();
        }

        int channels = ctype switch { 0 => 1, 2 => 3, 3 => 1, 4 => 2, 6 => 4, _ => 1 };
        int bpp = Math.Max(1, channels * depth / 8);
        var rgb = new byte[w * h * 3];
        byte[]? alpha = (ctype is 4 or 6 || trns != null) ? new byte[w * h] : null;
        if (alpha != null) Array.Fill(alpha, (byte)255);

        void Pass(int x0, int y0, int dx, int dy, ref int pos)
        {
            int pw = (w - x0 + dx - 1) / dx, ph = (h - y0 + dy - 1) / dy;
            if (pw <= 0 || ph <= 0) return;
            int stride = (pw * channels * depth + 7) / 8;
            var prev = new byte[stride];
            var cur = new byte[stride];
            for (int y = 0; y < ph; y++)
            {
                int filter = raw[pos++];
                Array.Copy(raw, pos, cur, 0, Math.Min(stride, raw.Length - pos));
                pos += stride;
                Unfilter(filter, cur, prev, bpp);
                for (int x = 0; x < pw; x++)
                    Pixel(cur, x, (y0 + y * dy) * w + x0 + x * dx);
                (prev, cur) = (cur, prev);
            }
        }

        int Sample(byte[] row, int index)
        {
            switch (depth)
            {
                case 8: return row[index];
                case 16: return row[index * 2];
                default:
                    int bitPos = index * depth;
                    int b = row[bitPos >> 3];
                    int shift = 8 - depth - (bitPos & 7);
                    return (b >> shift) & ((1 << depth) - 1);
            }
        }

        int Scale(int v) => depth switch { 1 => v * 255, 2 => v * 85, 4 => v * 17, _ => v };

        int Raw16(byte[] row, int index) => depth == 16 ? (row[index * 2] << 8) | row[index * 2 + 1] : Sample(row, index);

        void Pixel(byte[] row, int x, int idx)
        {
            int o = idx * 3;
            switch (ctype)
            {
                case 0:
                {
                    int v = Scale(Sample(row, x));
                    rgb[o] = rgb[o + 1] = rgb[o + 2] = (byte)v;
                    if (trns is { Length: >= 2 } && alpha != null)
                    {
                        int t = (trns[0] << 8) | trns[1];
                        if (Raw16(row, x) == t) alpha[idx] = 0;
                    }
                    break;
                }
                case 2:
                {
                    rgb[o] = (byte)Sample(row, x * 3); rgb[o + 1] = (byte)Sample(row, x * 3 + 1); rgb[o + 2] = (byte)Sample(row, x * 3 + 2);
                    if (trns is { Length: >= 6 } && alpha != null)
                    {
                        if (Raw16(row, x * 3) == ((trns[0] << 8) | trns[1]) && Raw16(row, x * 3 + 1) == ((trns[2] << 8) | trns[3]) && Raw16(row, x * 3 + 2) == ((trns[4] << 8) | trns[5]))
                            alpha[idx] = 0;
                    }
                    break;
                }
                case 3:
                {
                    int i = Sample(row, x);
                    if (palette != null && i * 3 + 2 < palette.Length)
                    { rgb[o] = palette[i * 3]; rgb[o + 1] = palette[i * 3 + 1]; rgb[o + 2] = palette[i * 3 + 2]; }
                    if (alpha != null && trns != null && i < trns.Length) alpha[idx] = trns[i];
                    break;
                }
                case 4:
                    rgb[o] = rgb[o + 1] = rgb[o + 2] = (byte)Sample(row, x * 2);
                    alpha![idx] = (byte)Sample(row, x * 2 + 1);
                    break;
                case 6:
                    rgb[o] = (byte)Sample(row, x * 4); rgb[o + 1] = (byte)Sample(row, x * 4 + 1); rgb[o + 2] = (byte)Sample(row, x * 4 + 2);
                    alpha![idx] = (byte)Sample(row, x * 4 + 3);
                    break;
            }
        }

        int pos = 0;
        if (interlace == 0) Pass(0, 0, 1, 1, ref pos);
        else
        {
            int[][] adam = [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]];
            foreach (var a in adam) Pass(a[0], a[1], a[2], a[3], ref pos);
        }

        if (alpha != null && alpha.All(a => a == 255)) alpha = null;
        return new PdfImageData { Width = w, Height = h, Data = rgb, Alpha = alpha };
    }

    static void Unfilter(int filter, byte[] cur, byte[] prev, int bpp)
    {
        int n = cur.Length;
        switch (filter)
        {
            case 1: for (int i = bpp; i < n; i++) cur[i] += cur[i - bpp]; break;
            case 2: for (int i = 0; i < n; i++) cur[i] += prev[i]; break;
            case 3:
                for (int i = 0; i < n; i++)
                    cur[i] += (byte)(((i >= bpp ? cur[i - bpp] : 0) + prev[i]) >> 1);
                break;
            case 4:
                for (int i = 0; i < n; i++)
                {
                    int a = i >= bpp ? cur[i - bpp] : 0, b = prev[i], c = i >= bpp ? prev[i - bpp] : 0;
                    int pp = a + b - c, pa = Math.Abs(pp - a), pb = Math.Abs(pp - b), pc = Math.Abs(pp - c);
                    cur[i] += (byte)(pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
                }
                break;
        }
    }

    // ------------------------------------------------------------ GIF (первый кадр)
    static PdfImageData? Gif(byte[] d)
    {
        int w = d[6] | (d[7] << 8), h = d[8] | (d[9] << 8);
        int flags = d[10];
        int p = 13;
        byte[]? global = null;
        if ((flags & 0x80) != 0) { int size = 3 << ((flags & 7) + 1); global = d.AsSpan(p, size).ToArray(); p += size; }
        int transparent = -1;
        while (p < d.Length)
        {
            int b = d[p++];
            if (b == 0x21)
            {
                int label = d[p++];
                if (label == 0xF9 && d[p] >= 4 && (d[p + 1] & 1) != 0) transparent = d[p + 4];
                while (p < d.Length && d[p] != 0) p += d[p] + 1;
                p++;
            }
            else if (b == 0x2C)
            {
                int ix = d[p] | (d[p + 1] << 8), iy = d[p + 2] | (d[p + 3] << 8);
                int iw = d[p + 4] | (d[p + 5] << 8), ih = d[p + 6] | (d[p + 7] << 8);
                int lf = d[p + 8];
                p += 9;
                byte[]? pal = global;
                if ((lf & 0x80) != 0) { int size = 3 << ((lf & 7) + 1); pal = d.AsSpan(p, size).ToArray(); p += size; }
                bool interlaced = (lf & 0x40) != 0;
                int minCode = d[p++];
                using var ms = new MemoryStream();
                while (p < d.Length && d[p] != 0) { ms.Write(d, p + 1, d[p]); p += d[p] + 1; }
                var idx = LzwDecode(ms.ToArray(), minCode, iw * ih);
                var rgb = new byte[w * h * 3];
                var alpha = new byte[w * h];
                int[] rows = new int[ih];
                if (interlaced)
                {
                    int r = 0;
                    foreach (var (start, step) in new[] { (0, 8), (4, 8), (2, 4), (1, 2) })
                        for (int y = start; y < ih; y += step) rows[r++] = y;
                }
                else for (int y = 0; y < ih; y++) rows[y] = y;
                for (int i = 0; i < Math.Min(idx.Length, iw * ih); i++)
                {
                    int x = ix + i % iw, y = iy + rows[i / iw];
                    if (x >= w || y >= h) continue;
                    int c = idx[i], o = y * w + x;
                    if (c == transparent) continue;
                    alpha[o] = 255;
                    if (pal != null && c * 3 + 2 < pal.Length)
                    { rgb[o * 3] = pal[c * 3]; rgb[o * 3 + 1] = pal[c * 3 + 1]; rgb[o * 3 + 2] = pal[c * 3 + 2]; }
                }
                return new PdfImageData { Width = w, Height = h, Data = rgb, Alpha = alpha.All(a => a == 255) ? null : alpha };
            }
            else break;
        }
        return null;
    }

    static byte[] LzwDecode(byte[] data, int minCode, int max)
    {
        var output = new List<byte>(max);
        int clear = 1 << minCode, eoi = clear + 1;
        var prefix = new int[4096];
        var suffix = new byte[4096];
        var stack = new byte[4097];
        int codeSize = minCode + 1, next = eoi + 1, old = -1;
        byte first = 0;
        int bits = 0, buf = 0, pos = 0;
        for (int i = 0; i < clear; i++) suffix[i] = (byte)i;
        while (output.Count < max)
        {
            while (bits < codeSize && pos < data.Length) { buf |= data[pos++] << bits; bits += 8; }
            if (bits < codeSize) break;
            int code = buf & ((1 << codeSize) - 1);
            buf >>= codeSize; bits -= codeSize;
            if (code == clear) { codeSize = minCode + 1; next = eoi + 1; old = -1; continue; }
            if (code == eoi) break;
            if (old == -1) { output.Add(suffix[code]); old = code; first = suffix[code]; continue; }
            int inCode = code, sp = 0;
            if (code >= next) { stack[sp++] = first; code = old; }
            while (code >= clear) { stack[sp++] = suffix[code]; code = prefix[code]; }
            first = suffix[code];
            stack[sp++] = first;
            while (sp > 0) output.Add(stack[--sp]);
            if (next < 4096) { prefix[next] = old; suffix[next] = first; next++; if (next == (1 << codeSize) && codeSize < 12) codeSize++; }
            old = inCode;
        }
        return output.ToArray();
    }

    // ------------------------------------------------------------ BMP
    static PdfImageData? Bmp(byte[] d)
    {
        int dataOff = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(10));
        int hdr = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(14));
        int w = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(18));
        int hRaw = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(22));
        int bpp = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(28));
        int compression = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(30));
        if (compression is not (0 or 3) || w <= 0) return null;
        bool topDown = hRaw < 0;
        int h = Math.Abs(hRaw);
        int stride = (w * bpp + 31) / 32 * 4;
        int palOff = 14 + hdr;
        var rgb = new byte[w * h * 3];
        byte[]? alpha = bpp == 32 ? new byte[w * h] : null;
        for (int y = 0; y < h; y++)
        {
            int srcRow = dataOff + (topDown ? y : h - 1 - y) * stride;
            for (int x = 0; x < w; x++)
            {
                int o = (y * w + x) * 3;
                byte r, g, b;
                switch (bpp)
                {
                    case 24: b = d[srcRow + x * 3]; g = d[srcRow + x * 3 + 1]; r = d[srcRow + x * 3 + 2]; break;
                    case 16:
                    {
                        int v = d[srcRow + x * 2] | d[srcRow + x * 2 + 1] << 8;
                        bool rgb565 = compression == 3 && hdr >= 40 && BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(54)) == 0xF800;   // маски цветов — сразу после 40 байт заголовка
                        if (rgb565) { r = (byte)((v >> 11 & 31) * 255 / 31); g = (byte)((v >> 5 & 63) * 255 / 63); b = (byte)((v & 31) * 255 / 31); }
                        else { r = (byte)((v >> 10 & 31) * 255 / 31); g = (byte)((v >> 5 & 31) * 255 / 31); b = (byte)((v & 31) * 255 / 31); }
                        break;
                    }
                    case 32: b = d[srcRow + x * 4]; g = d[srcRow + x * 4 + 1]; r = d[srcRow + x * 4 + 2]; alpha![y * w + x] = d[srcRow + x * 4 + 3]; break;
                    case 8: case 4: case 1:
                    {
                        int bit = x * bpp;
                        int v = (d[srcRow + (bit >> 3)] >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
                        int po = palOff + v * 4;
                        b = d[po]; g = d[po + 1]; r = d[po + 2];
                        break;
                    }
                    default: return null;
                }
                rgb[o] = r; rgb[o + 1] = g; rgb[o + 2] = b;
            }
        }
        if (alpha != null && (alpha.All(a => a == 255) || alpha.All(a => a == 0))) alpha = null;
        return new PdfImageData { Width = w, Height = h, Data = rgb, Alpha = alpha };
    }
}
