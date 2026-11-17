using System.Buffers;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using Lumina.Data.Structs.Excel;

namespace XivExdUnpacker.src.Core;

public sealed class Utf8CsvBuffer
{
    private const int InitialCapacity = 65536;
    private const int MaxRetainedCapacity = 2 * 1024 * 1024;

    private static readonly byte[] NewLineBytes = Encoding.UTF8.GetBytes(Environment.NewLine);
    private static readonly SearchValues<char> CsvEscapeChars = SearchValues.Create(",\"\n\r");
    private static ReadOnlySpan<byte> HexDigits => "0123456789ABCDEF"u8;

    private byte[] _buf = new byte[InitialCapacity];
    private int _pos;

    public int Length => _pos;
    public ReadOnlySpan<byte> WrittenSpan => _buf.AsSpan(0, _pos);

    public void Reset()
    {
        _pos = 0;
        if (_buf.Length > MaxRetainedCapacity)
            _buf = new byte[InitialCapacity];
    }

    public void Rollback(int pos) => _pos = pos;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void Ensure(int extra)
    {
        if (_pos + extra > _buf.Length)
            Grow(extra);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void Grow(int extra)
    {
        int next = Math.Max(_buf.Length * 2, _pos + extra);
        Array.Resize(ref _buf, next);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteBom()
    {
        Ensure(3);
        _buf[_pos++] = 0xEF;
        _buf[_pos++] = 0xBB;
        _buf[_pos++] = 0xBF;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteByte(byte b)
    {
        Ensure(1);
        _buf[_pos++] = b;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ReadOnlySpan<byte> bytes)
    {
        Ensure(bytes.Length);
        bytes.CopyTo(_buf.AsSpan(_pos));
        _pos += bytes.Length;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteNewLine() => Write(NewLineBytes);

    public void WriteString(string s)
    {
        if (string.IsNullOrEmpty(s))
            return;

        Ensure(Encoding.UTF8.GetMaxByteCount(s.Length));
        _pos += Encoding.UTF8.GetBytes(s, _buf.AsSpan(_pos));
    }

    public void WriteHexBytes(List<byte> bytes)
    {
        Ensure(bytes.Count * 2);
        var hex = HexDigits;
        int p = _pos;
        for (int i = 0; i < bytes.Count; i++)
        {
            byte b = bytes[i];
            _buf[p++] = hex[b >> 4];
            _buf[p++] = hex[b & 0xF];
        }
        _pos = p;
    }

    public void WriteHexSpan(ReadOnlySpan<byte> bytes)
    {
        Ensure(bytes.Length * 2);
        var hex = HexDigits;
        int p = _pos;
        for (int i = 0; i < bytes.Length; i++)
        {
            byte b = bytes[i];
            _buf[p++] = hex[b >> 4];
            _buf[p++] = hex[b & 0xF];
        }
        _pos = p;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteCsvEscapedInnerUtf8(ReadOnlySpan<byte> utf8)
    {
        int idx = utf8.IndexOf((byte)'"');
        if (idx < 0)
        {
            Write(utf8);
            return;
        }

        int start = 0;
        while (start < utf8.Length)
        {
            idx = utf8[start..].IndexOf((byte)'"');
            if (idx < 0)
            {
                Write(utf8[start..]);
                break;
            }
            if (idx > 0)
                Write(utf8.Slice(start, idx));
            Write("\"\""u8);
            start += idx + 1;
        }
    }

    public void WriteQuotedCsvUtf8(ReadOnlySpan<byte> utf8)
    {
        WriteByte((byte)'"');
        WriteCsvEscapedInnerUtf8(utf8);
        WriteByte((byte)'"');
    }

    public void WriteQuotedCsvString(string value)
    {
        WriteByte((byte)'"');
        if (!string.IsNullOrEmpty(value))
        {
            int start = 0;
            while (start < value.Length)
            {
                int quoteIdx = value.IndexOf('"', start);
                if (quoteIdx < 0)
                {
                    WriteSpanChars(value.AsSpan(start));
                    break;
                }
                if (quoteIdx > start)
                    WriteSpanChars(value.AsSpan(start, quoteIdx - start));
                Write("\"\""u8);
                start = quoteIdx + 1;
            }
        }
        WriteByte((byte)'"');
    }

    public void WriteCsvEscapedString(string value)
    {
        if (string.IsNullOrEmpty(value))
            return;

        if (value.AsSpan().IndexOfAny(CsvEscapeChars) < 0)
        {
            WriteString(value);
            return;
        }
        WriteQuotedCsvString(value);
    }

    private void WriteSpanChars(ReadOnlySpan<char> chars)
    {
        Ensure(Encoding.UTF8.GetMaxByteCount(chars.Length));
        _pos += Encoding.UTF8.GetBytes(chars, _buf.AsSpan(_pos));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteFormattable<T>(T val)
        where T : IUtf8SpanFormattable
    {
        Ensure(64);
        if (val.TryFormat(_buf.AsSpan(_pos), out int written, default, CultureInfo.InvariantCulture))
        {
            _pos += written;
            return;
        }
        throw new InvalidOperationException("UTF-8 formatting failed");
    }

    public void WriteColumnType(ExcelColumnDataType type)
    {
        switch (type)
        {
            case ExcelColumnDataType.String:
                Write("String"u8);
                break;
            case ExcelColumnDataType.Bool:
                Write("Bool"u8);
                break;
            case ExcelColumnDataType.Int8:
                Write("Int8"u8);
                break;
            case ExcelColumnDataType.UInt8:
                Write("UInt8"u8);
                break;
            case ExcelColumnDataType.Int16:
                Write("Int16"u8);
                break;
            case ExcelColumnDataType.UInt16:
                Write("UInt16"u8);
                break;
            case ExcelColumnDataType.Int32:
                Write("Int32"u8);
                break;
            case ExcelColumnDataType.UInt32:
                Write("UInt32"u8);
                break;
            case ExcelColumnDataType.Float32:
                Write("Float32"u8);
                break;
            case ExcelColumnDataType.Int64:
                Write("Int64"u8);
                break;
            case ExcelColumnDataType.UInt64:
                Write("UInt64"u8);
                break;
            case ExcelColumnDataType.PackedBool0:
                Write("PackedBool0"u8);
                break;
            case ExcelColumnDataType.PackedBool1:
                Write("PackedBool1"u8);
                break;
            case ExcelColumnDataType.PackedBool2:
                Write("PackedBool2"u8);
                break;
            case ExcelColumnDataType.PackedBool3:
                Write("PackedBool3"u8);
                break;
            case ExcelColumnDataType.PackedBool4:
                Write("PackedBool4"u8);
                break;
            case ExcelColumnDataType.PackedBool5:
                Write("PackedBool5"u8);
                break;
            case ExcelColumnDataType.PackedBool6:
                Write("PackedBool6"u8);
                break;
            case ExcelColumnDataType.PackedBool7:
                Write("PackedBool7"u8);
                break;
            default:
                WriteString(type.ToString());
                break;
        }
    }
}
