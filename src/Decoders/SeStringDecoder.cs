using System.Buffers;
using System.Text;
using Lumina.Text.Payloads;
using Lumina.Text.ReadOnly;

namespace XivExdUnpacker.Decoders;

public static class SeStringDecoder
{
    [ThreadStatic]
    private static StringBuilder? t_sb;

    [ThreadStatic]
    private static List<byte>? t_hexBuf;

    public static string DecodeHexcode(ReadOnlySpan<byte> data) =>
        data.IsEmpty ? string.Empty : DecodeHexcode(data.ToArray().AsMemory());

    public static string DecodeHexcode(ReadOnlyMemory<byte> data)
    {
        if (data.IsEmpty)
            return string.Empty;

        var sb = t_sb ??= new StringBuilder(256);
        var hexBuf = t_hexBuf ??= new List<byte>(64);
        sb.Clear();
        hexBuf.Clear();

        try
        {
            AppendHexcode(sb, hexBuf, new ReadOnlySeString(data), flushPerPayload: true);
            FlushHex(sb, hexBuf);
            return sb.ToString();
        }
        catch
        {
            hexBuf.Clear();
            return $"<hex:{Convert.ToHexString(data.Span)}>";
        }
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

    private static void AppendHexcode(
        StringBuilder sb,
        List<byte> hexBuf,
        ReadOnlySeString se,
        bool flushPerPayload
    )
    {
        var raw = se.Data.Span;
        foreach (var (payloadOffset, payload) in se.GetOffsetEnumerator())
        {
            switch (payload.Type)
            {
                case ReadOnlySePayloadType.Text:
                    AppendSafeUtf8(sb, hexBuf, payload.Body.Span);
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
                            AppendExpression(sb, hexBuf, expr);

                        bodyPos = exprOffset + expr.Body.Length;
                        exprIndex++;
                    }

                    if (bodyPos < bodySpan.Length)
                        AddHex(hexBuf, bodySpan[bodyPos..]);

                    AddHex(hexBuf, raw.Slice(payloadOffset + payload.EnvelopeByteLength - 1, 1));
                    if (flushPerPayload)
                        FlushHex(sb, hexBuf);
                    break;
                }
            }
        }
    }

    private static void AppendExpression(
        StringBuilder sb,
        List<byte> hexBuf,
        ReadOnlySeExpression expr
    )
    {
        if (expr.TryGetString(out var inner))
        {
            if (inner.ByteLength > 0)
            {
                var prefixLen = expr.Body.Length - inner.ByteLength;
                if (prefixLen > 0)
                    AddHex(hexBuf, expr.Body.Span[..prefixLen]);

                FlushHex(sb, hexBuf);
                AppendHexcode(sb, hexBuf, inner, flushPerPayload: true);
            }
            else
            {
                AddHex(hexBuf, expr.Body.Span);
                FlushHex(sb, hexBuf);
            }
            return;
        }

        if (expr.TryGetParameterExpression(out _, out var operand))
        {
            var opStart = expr.Body.Length - operand.Body.Length;
            if (opStart > 0)
                AddHex(hexBuf, expr.Body.Span[..opStart]);
            AppendExpression(sb, hexBuf, operand);
            return;
        }

        if (expr.TryGetBinaryExpression(out _, out var op1, out var op2))
        {
            var opsLen = op1.Body.Length + op2.Body.Length;
            var op1Start = expr.Body.Length - opsLen;
            if (op1Start > 0)
                AddHex(hexBuf, expr.Body.Span[..op1Start]);
            AppendExpression(sb, hexBuf, op1);
            AppendExpression(sb, hexBuf, op2);
            return;
        }

        AddHex(hexBuf, expr.Body.Span);
    }

    private static void AppendSafeUtf8(StringBuilder sb, List<byte> hexBuf, ReadOnlySpan<byte> bytes)
    {
        int i = 0;
        int textStart = 0;

        while (i < bytes.Length)
        {
            var status = Rune.DecodeFromUtf8(bytes[i..], out var rune, out var consumed);
            if (status == OperationStatus.Done &&
                (!Rune.IsControl(rune) || rune.Value is '\t' or '\r' or '\n'))
            {
                i += consumed;
                continue;
            }

            if (i > textStart)
            {
                FlushHex(sb, hexBuf);
                sb.Append(Encoding.UTF8.GetString(bytes[textStart..i]));
            }

            hexBuf.Add(bytes[i]);
            i++;
            textStart = i;
        }

        if (i > textStart)
        {
            FlushHex(sb, hexBuf);
            sb.Append(Encoding.UTF8.GetString(bytes[textStart..i]));
        }
    }

    private static void AddHex(List<byte> hexBuf, ReadOnlySpan<byte> bytes)
    {
        for (int i = 0; i < bytes.Length; i++)
            hexBuf.Add(bytes[i]);
    }

    private static void FlushHex(StringBuilder sb, List<byte> hexBuf)
    {
        if (hexBuf.Count == 0)
            return;

        sb.Append("<hex:");
        for (int i = 0; i < hexBuf.Count; i++)
            sb.Append(hexBuf[i].ToString("X2"));
        sb.Append('>');
        hexBuf.Clear();
    }
}
