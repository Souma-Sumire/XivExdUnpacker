using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Unicode;
using Lumina.Text.Payloads;
using Lumina.Text.ReadOnly;
using XivExdUnpacker.src.Core;

namespace XivExdUnpacker.Decoders;

public static class SeStringDecoder
{
    private static readonly SearchValues<byte> HexSpecialOrQuote = SearchValues.Create(
        "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F\"\x7F"u8
    );

    private static readonly SearchValues<byte> ReadableSpecialOrQuote = SearchValues.Create(
        "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F\"<\\\x7F"u8
    );

    [ThreadStatic]
    private static Utf8CsvBuffer? t_utf8Buf;

    [ThreadStatic]
    private static List<byte>? t_hexBuf;

    public static string DecodeHexcode(ReadOnlySpan<byte> data) =>
        data.IsEmpty ? string.Empty : DecodeHexcode(data.ToArray().AsMemory());

    public static string DecodeHexcode(ReadOnlyMemory<byte> data)
    {
        if (data.IsEmpty)
            return string.Empty;

        var buf = t_utf8Buf ??= new Utf8CsvBuffer();
        buf.Reset();
        WriteHexcodeCore(buf, data, escapeCsvQuotes: false);
        return Encoding.UTF8.GetString(buf.WrittenSpan);
    }

    public static string DecodeReadable(ReadOnlySpan<byte> data) =>
        data.IsEmpty ? string.Empty : DecodeReadable(data.ToArray().AsMemory());

    public static string DecodeReadable(ReadOnlyMemory<byte> data)
    {
        if (data.IsEmpty)
            return string.Empty;

        try
        {
            return new ReadOnlySeString(data).ToMacroString();
        }
        catch
        {
            return Encoding.UTF8.GetString(data.Span);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteHexcodeCsvField(
        Utf8CsvBuffer target,
        byte[] data,
        int offset,
        int length
    )
    {
        var span = data.AsSpan(offset, length);
        if (TryWritePlainCsvField(target, span, HexSpecialOrQuote, disallowAngleAndSlash: false))
            return;

        target.WriteByte((byte)'"');
        WriteHexcodeCore(target, data.AsMemory(offset, length), escapeCsvQuotes: true);
        target.WriteByte((byte)'"');
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteReadableCsvField(
        Utf8CsvBuffer target,
        byte[] data,
        int offset,
        int length
    )
    {
        var span = data.AsSpan(offset, length);
        if (TryWritePlainCsvField(target, span, ReadableSpecialOrQuote, disallowAngleAndSlash: true))
            return;

        target.WriteQuotedCsvString(DecodeReadable(data.AsMemory(offset, length)));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryWritePlainCsvField(
        Utf8CsvBuffer target,
        ReadOnlySpan<byte> span,
        SearchValues<byte> specialOrQuote,
        bool disallowAngleAndSlash
    )
    {
        if (!span.ContainsAnyInRange((byte)0x80, (byte)0xFF))
        {
            if (span.IndexOfAny(specialOrQuote) < 0)
            {
                target.WriteByte((byte)'"');
                target.Write(span);
                target.WriteByte((byte)'"');
                return true;
            }

            if (
                HasNoControlBytes(span)
                && (
                    !disallowAngleAndSlash
                    || (!span.Contains((byte)'<') && !span.Contains((byte)'\\'))
                )
            )
            {
                target.WriteQuotedCsvUtf8(span);
                return true;
            }

            return false;
        }

        // 0xC2 可能是 UTF-8 中的 C1 控制字符 (U+0080..U+009F)，交给完整路径精确判断
        if (
            !span.Contains((byte)0x02)
            && HasNoControlBytes(span)
            && !span.Contains((byte)0xC2)
            && (!disallowAngleAndSlash || (!span.Contains((byte)'<') && !span.Contains((byte)'\\')))
            && Utf8.IsValid(span)
        )
        {
            target.WriteQuotedCsvUtf8(span);
            return true;
        }

        return false;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool HasNoControlBytes(ReadOnlySpan<byte> span) =>
        !span.ContainsAnyInRange((byte)0x00, (byte)0x08)
        && !span.ContainsAnyInRange((byte)0x0B, (byte)0x0C)
        && !span.ContainsAnyInRange((byte)0x0E, (byte)0x1F)
        && !span.Contains((byte)0x7F);

    private static void WriteHexcodeCore(
        Utf8CsvBuffer target,
        ReadOnlyMemory<byte> data,
        bool escapeCsvQuotes
    )
    {
        var hexBuf = t_hexBuf ??= new List<byte>(64);
        hexBuf.Clear();
        int rollbackPos = target.Length;

        try
        {
            AppendHexcode(
                target,
                hexBuf,
                new ReadOnlySeString(data),
                flushPerPayload: true,
                escapeCsvQuotes
            );
            FlushHex(target, hexBuf);
        }
        catch
        {
            hexBuf.Clear();
            target.Rollback(rollbackPos);
            target.Write("<hex:"u8);
            target.WriteHexSpan(data.Span);
            target.WriteByte((byte)'>');
        }
    }

    private static void AppendHexcode(
        Utf8CsvBuffer target,
        List<byte> hexBuf,
        ReadOnlySeString se,
        bool flushPerPayload,
        bool escapeCsvQuotes
    )
    {
        var raw = se.Data.Span;
        foreach (var (payloadOffset, payload) in se.GetOffsetEnumerator())
        {
            switch (payload.Type)
            {
                case ReadOnlySePayloadType.Text:
                    AppendSafeUtf8(target, hexBuf, payload.Body.Span, escapeCsvQuotes);
                    break;

                case ReadOnlySePayloadType.Invalid:
                    AddHex(hexBuf, payload.Body.Span);
                    break;

                case ReadOnlySePayloadType.Macro:
                {
                    var headLen = payload.EnvelopeByteLength - payload.Body.Length - 1;
                    AddHex(hexBuf, raw.Slice(payloadOffset, headLen));

                    bool keepExtraArgsAsHex = (byte)payload.MacroCode is 0x22 or 0x26;
                    var bodySpan = payload.Body.Span;
                    var bodyPos = 0;
                    var exprIndex = 0;

                    foreach (var (exprOffset, expr) in payload.GetOffsetEnumerator())
                    {
                        if (exprOffset > bodyPos)
                            AddHex(hexBuf, bodySpan[bodyPos..exprOffset]);

                        if (keepExtraArgsAsHex && exprIndex > 0)
                            AddHex(hexBuf, expr.Body.Span);
                        else
                            AppendExpression(target, hexBuf, expr, escapeCsvQuotes);

                        bodyPos = exprOffset + expr.Body.Length;
                        exprIndex++;
                    }

                    if (bodyPos < bodySpan.Length)
                        AddHex(hexBuf, bodySpan[bodyPos..]);

                    AddHex(hexBuf, raw.Slice(payloadOffset + payload.EnvelopeByteLength - 1, 1));
                    if (flushPerPayload)
                        FlushHex(target, hexBuf);
                    break;
                }
            }
        }
    }

    private static void AppendExpression(
        Utf8CsvBuffer target,
        List<byte> hexBuf,
        ReadOnlySeExpression expr,
        bool escapeCsvQuotes
    )
    {
        if (expr.TryGetString(out var inner))
        {
            if (inner.ByteLength > 0)
            {
                var prefixLen = expr.Body.Length - inner.ByteLength;
                if (prefixLen > 0)
                    AddHex(hexBuf, expr.Body.Span[..prefixLen]);

                FlushHex(target, hexBuf);
                AppendHexcode(target, hexBuf, inner, flushPerPayload: true, escapeCsvQuotes);
            }
            else
            {
                AddHex(hexBuf, expr.Body.Span);
                FlushHex(target, hexBuf);
            }
            return;
        }

        if (expr.TryGetParameterExpression(out _, out var operand))
        {
            var opStart = expr.Body.Length - operand.Body.Length;
            if (opStart > 0)
                AddHex(hexBuf, expr.Body.Span[..opStart]);
            AppendExpression(target, hexBuf, operand, escapeCsvQuotes);
            return;
        }

        if (expr.TryGetBinaryExpression(out _, out var op1, out var op2))
        {
            var opsLen = op1.Body.Length + op2.Body.Length;
            var op1Start = expr.Body.Length - opsLen;
            if (op1Start > 0)
                AddHex(hexBuf, expr.Body.Span[..op1Start]);
            AppendExpression(target, hexBuf, op1, escapeCsvQuotes);
            AppendExpression(target, hexBuf, op2, escapeCsvQuotes);
            return;
        }

        AddHex(hexBuf, expr.Body.Span);
    }

    private static void AppendSafeUtf8(
        Utf8CsvBuffer target,
        List<byte> hexBuf,
        ReadOnlySpan<byte> bytes,
        bool escapeCsvQuotes
    )
    {
        int i = 0;
        int textStart = 0;

        while (i < bytes.Length)
        {
            var status = Rune.DecodeFromUtf8(bytes[i..], out var rune, out var consumed);
            if (
                status == OperationStatus.Done
                && (!Rune.IsControl(rune) || rune.Value is '\t' or '\r' or '\n')
            )
            {
                i += consumed;
                continue;
            }

            if (i > textStart)
            {
                FlushHex(target, hexBuf);
                WriteTextSegment(target, bytes[textStart..i], escapeCsvQuotes);
            }

            hexBuf.Add(bytes[i]);
            i++;
            textStart = i;
        }

        if (i > textStart)
        {
            FlushHex(target, hexBuf);
            WriteTextSegment(target, bytes[textStart..i], escapeCsvQuotes);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void WriteTextSegment(
        Utf8CsvBuffer target,
        ReadOnlySpan<byte> bytes,
        bool escapeCsvQuotes
    )
    {
        if (escapeCsvQuotes)
            target.WriteCsvEscapedInnerUtf8(bytes);
        else
            target.Write(bytes);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void AddHex(List<byte> hexBuf, ReadOnlySpan<byte> bytes)
    {
        for (int i = 0; i < bytes.Length; i++)
            hexBuf.Add(bytes[i]);
    }

    private static void FlushHex(Utf8CsvBuffer target, List<byte> hexBuf)
    {
        if (hexBuf.Count == 0)
            return;

        target.Write("<hex:"u8);
        target.WriteHexBytes(hexBuf);
        target.WriteByte((byte)'>');
        hexBuf.Clear();
    }
}
