using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SaintCoinach.Text;

namespace XivExdUnpacker.Decoders;

public class CompleteReadableDecoder
{
    private const byte TagStartMarker = 0x02;
    private const byte TagEndMarker = 0x03;
    private static readonly Encoding UTF8NoBom = new UTF8Encoding(false);

    public string Decode(byte[] buffer)
    {
        using var ms = new MemoryStream(buffer);
        using var reader = new BinaryReader(ms, UTF8NoBom);
        return DecodeString(reader, buffer.Length);
    }

    private string DecodeString(BinaryReader input, int length)
    {
        var end = input.BaseStream.Position + length;
        var sb = new StringBuilder();
        var pendingBytes = new List<byte>();

        while (input.BaseStream.Position < end)
        {
            var b = input.ReadByte();
            if (b == TagStartMarker)
            {
                if (pendingBytes.Count > 0)
                {
                    sb.Append(UTF8NoBom.GetString(pendingBytes.ToArray()));
                    pendingBytes.Clear();
                }
                sb.Append(DecodeTag(input));
            }
            else
            {
                pendingBytes.Add(b);
            }
        }

        if (pendingBytes.Count > 0)
            sb.Append(UTF8NoBom.GetString(pendingBytes.ToArray()));

        return sb.ToString();
    }

    private string DecodeTag(BinaryReader input)
    {
        var tagByte = input.ReadByte();
        var tag = (TagType)tagByte;
        var length = GetInteger(input);
        var end = input.BaseStream.Position + length;

        string result;

        if (length == 0)
        {
            result = tag switch
            {
                TagType.LineBreak => "\r\n",
                TagType.Dash => "–",
                _ => $"<{tag}/>",
            };
        }
        else
        {
            switch (tagByte)
            {
                case (byte)TagType.Color:
                    result = DecodeColorTag(input, tag, length);
                    break;
                case (byte)TagType.If:
                    result = DecodeIfTag(input, tag, end);
                    break;
                case (byte)TagType.IfEquals:
                    result = DecodeIfEqualsTag(input, tag, end);
                    break;
                case (byte)TagType.Switch:
                    result = DecodeSwitchTag(input, tag, end);
                    break;
                case (byte)TagType.Format:
                    result = DecodeFormatTag(input, tag, end);
                    break;
                case (byte)TagType.ZeroPaddedValue:
                    result = DecodeZeroPaddedValueTag(input, tag, end);
                    break;
                case (byte)TagType.Emphasis:
                case (byte)TagType.Emphasis2:
                    result = DecodeEmphasisTag(input, tag, end);
                    break;
                case (byte)TagType.Clickable:
                case (byte)TagType.CommandIcon:
                case (byte)TagType.Gui:
                case (byte)TagType.Sheet:
                case (byte)TagType.SheetEn:
                case (byte)TagType.SheetJa:
                case (byte)TagType.SheetDe:
                case (byte)TagType.SheetFr:
                case (byte)TagType.Split:
                case (byte)TagType.Time:
                    result = DecodeParameterizedTag(input, tag, end);
                    break;
                case (byte)TagType.Value:
                case (byte)TagType.Highlight:
                case (byte)TagType.TwoDigitValue:
                case (byte)TagType.InstanceContent:
                    result = DecodeWrappedExpressionTag(input, tag, end);
                    break;
                default:
                    result = DecodeHexContentTag(input, tag, end);
                    break;
            }
        }

        if (input.BaseStream.Position != end)
            input.BaseStream.Position = end;
        var endMarker = input.ReadByte();
        if (endMarker != TagEndMarker)
            throw new InvalidDataException($"Expected 0x03, got 0x{endMarker:X2}");

        return result;
    }

    private string DecodeColorTag(BinaryReader input, TagType tag, int length)
    {
        var t = input.ReadByte();
        if (length == 1 && t == 0xEC)
            return $"</{tag}>";

        var expr = DecodeExpression(input, (DecodeExpressionType)t);
        return $"<{tag}({expr})>";
    }

    private string DecodeFormatTag(BinaryReader input, TagType tag, long end)
    {
        var arg1 = DecodeExpression(input);
        var start = input.BaseStream.Position;
        var byteCount = (int)(end - start);
        var rawBytes = input.ReadBytes(byteCount);
        var arg2 = string.Concat(rawBytes.Select(b => b.ToString("X2")));
        return $"<{tag}({arg1},{arg2})/>";
    }

    private string DecodeZeroPaddedValueTag(BinaryReader input, TagType tag, long end)
    {
        var content = DecodeExpression(input);
        var arg = input.BaseStream.Position < end ? DecodeExpression(input) : "";
        return $"<{tag}({arg})>{content}</{tag}>";
    }

    private string DecodeIfTag(BinaryReader input, TagType tag, long end)
    {
        var condition = DecodeExpression(input);
        var exprs = new List<string>();
        while (input.BaseStream.Position < end)
            exprs.Add(DecodeExpression(input));

        var trueValue = exprs.Count > 0 ? exprs[0] : "";
        string? falseValue = exprs.Count > 1 ? exprs[1] : null;

        if (falseValue == null)
            return $"<{tag}({condition})>{trueValue}</{tag}>";
        return $"<{tag}({condition})>{trueValue}<Else/>{falseValue}</{tag}>";
    }

    private string DecodeSwitchTag(BinaryReader input, TagType tag, long end)
    {
        var switchValue = DecodeExpression(input);
        var sb = new StringBuilder();
        sb.Append($"<{tag}({switchValue})>");
        int caseIndex = 1;
        while (input.BaseStream.Position < end)
        {
            sb.Append($"<Case({caseIndex})>{DecodeExpression(input)}</Case>");
            caseIndex++;
        }
        sb.Append($"</{tag}>");
        return sb.ToString();
    }

    private string DecodeIfEqualsTag(BinaryReader input, TagType tag, long end)
    {
        var left = DecodeExpression(input);
        var right = DecodeExpression(input);
        var exprs = new List<string>();
        while (input.BaseStream.Position < end)
            exprs.Add(DecodeExpression(input));
        var trueValue = exprs.Count > 0 ? exprs[0] : "";
        string? falseValue = exprs.Count > 1 ? exprs[1] : null;
        if (falseValue == null)
            return $"<{tag}({left},{right})>{trueValue}</{tag}>";
        return $"<{tag}({left},{right})>{trueValue}<Else/>{falseValue}</{tag}>";
    }

    private string DecodeParameterizedTag(BinaryReader input, TagType tag, long end)
    {
        var args = new List<string>();
        while (input.BaseStream.Position < end)
            args.Add(DecodeExpression(input));
        return args.Count > 0 ? $"<{tag}({string.Join(",", args)})/>" : $"<{tag}/>";
    }

    private string DecodeWrappedExpressionTag(BinaryReader input, TagType tag, long end)
    {
        if (input.BaseStream.Position >= end)
            return $"<{tag}/>";
        var content = DecodeExpression(input);
        return $"<{tag}>{content}</{tag}>";
    }

    private string DecodeHexContentTag(BinaryReader input, TagType tag, long end)
    {
        var start = input.BaseStream.Position;
        var byteCount = (int)(end - start);
        var rawBytes = input.ReadBytes(byteCount);
        return $"<{tag}>{string.Concat(rawBytes.Select(b => b.ToString("X2")))}</{tag}>";
    }

    private string DecodeEmphasisTag(BinaryReader input, TagType tag, long end)
    {
        var status = GetInteger(input);
        if (status == 0)
            return $"</{tag}>";
        if (status == 1)
            return $"<{tag}>";
        throw new InvalidDataException($"Unexpected emphasis status: {status}");
    }

    private string DecodeExpression(BinaryReader input)
    {
        var b = input.ReadByte();
        return DecodeExpression(input, (DecodeExpressionType)b);
    }

    private string DecodeExpression(BinaryReader input, DecodeExpressionType exprType)
    {
        var t = (byte)exprType;
        if (t < 0xD0)
            return (t - 1).ToString();
        if (t < 0xE0)
            return $"TopLevelParameter({t - 1})";

        switch (exprType)
        {
            case DecodeExpressionType.Decode:
                return DecodeNestedString(input);
            case DecodeExpressionType.Byte:
                return GetInteger(input, IntegerType.Byte).ToString();
            case DecodeExpressionType.Int16_MinusOne:
                return (GetInteger(input, IntegerType.Int16) - 1).ToString();
            case DecodeExpressionType.Int16_1:
            case DecodeExpressionType.Int16_2:
                return GetInteger(input, IntegerType.Int16).ToString();
            case DecodeExpressionType.Int24_MinusOne:
                return (GetInteger(input, IntegerType.Int24) - 1).ToString();
            case DecodeExpressionType.Int24:
                return GetInteger(input, IntegerType.Int24).ToString();
            case DecodeExpressionType.Int24_Lsh8:
                return (GetInteger(input, IntegerType.Int24) << 8).ToString();
            case DecodeExpressionType.Int24_SafeZero:
                return DecodeInt24SafeZero(input).ToString();
            case DecodeExpressionType.Int32:
                return GetInteger(input, IntegerType.Int32).ToString();
            case DecodeExpressionType.IntegerParameter:
                return $"IntegerParameter({DecodeExpression(input)})";
            case DecodeExpressionType.PlayerParameter:
                return $"PlayerParameter({DecodeExpression(input)})";
            case DecodeExpressionType.StringParameter:
                return $"StringParameter({DecodeExpression(input)})";
            case DecodeExpressionType.ObjectParameter:
                return $"ObjectParameter({DecodeExpression(input)})";
            case DecodeExpressionType.GreaterThanOrEqualTo:
                return $"GreaterThanOrEqualTo({DecodeExpression(input)},{DecodeExpression(input)})";
            case DecodeExpressionType.GreaterThan:
                return $"GreaterThan({DecodeExpression(input)},{DecodeExpression(input)})";
            case DecodeExpressionType.LessThanOrEqualTo:
                return $"LessThanOrEqualTo({DecodeExpression(input)},{DecodeExpression(input)})";
            case DecodeExpressionType.LessThan:
                return $"LessThan({DecodeExpression(input)},{DecodeExpression(input)})";
            case DecodeExpressionType.Equal:
                return $"Equal({DecodeExpression(input)},{DecodeExpression(input)})";
            case DecodeExpressionType.NotEqual:
                return $"NotEqual({DecodeExpression(input)},{DecodeExpression(input)})";
            default:
                return $"Unknown(0x{t:X2})";
        }
    }

    private string DecodeNestedString(BinaryReader input)
    {
        var length = GetInteger(input);
        return DecodeString(input, length);
    }

    private int DecodeInt24SafeZero(BinaryReader input)
    {
        var v16 = input.ReadByte();
        var v8 = input.ReadByte();
        var v0 = input.ReadByte();
        int v = 0;
        if (v16 != byte.MaxValue)
            v |= v16 << 16;
        if (v8 != byte.MaxValue)
            v |= v8 << 8;
        if (v0 != byte.MaxValue)
            v |= v0;
        return v;
    }

    private int GetInteger(BinaryReader input)
    {
        var typeByte = input.ReadByte();
        return GetInteger(input, (IntegerType)typeByte);
    }

    private int GetInteger(BinaryReader input, IntegerType type)
    {
        const byte ByteLengthCutoff = 0xF0;
        var t = (byte)type;
        if (t < ByteLengthCutoff)
            return t - 1;

        switch (type)
        {
            case IntegerType.Byte:
                return input.ReadByte();
            case IntegerType.ByteTimes256:
                return input.ReadByte() * 256;
            case IntegerType.Int16:
                return (input.ReadByte() << 8) | input.ReadByte();
            case IntegerType.Int24:
                return (input.ReadByte() << 16) | (input.ReadByte() << 8) | input.ReadByte();
            case IntegerType.Int32:
                return (input.ReadByte() << 24)
                    | (input.ReadByte() << 16)
                    | (input.ReadByte() << 8)
                    | input.ReadByte();
            default:
                throw new NotSupportedException($"Type: {type}");
        }
    }
}
