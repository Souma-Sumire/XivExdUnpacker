using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using Lumina;
using Lumina.Data;
using Lumina.Data.Files;
using Lumina.Data.Files.Excel;
using Lumina.Data.Structs.Excel;
using XivExdUnpacker.Decoders;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src.Core;

public record CellReadError(
    string SheetName,
    string RowId,
    ushort ColumnOffset,
    string ColumnType,
    string Message
);

internal sealed class SheetLayout
{
    public required ExcelHeaderFile HeaderFile { get; init; }
    public required byte[] HeaderCsvBytes { get; init; }
    public required ExcelColumnDefinition[] Columns { get; init; }
    public required ushort DataOffset { get; init; }
    public required ExcelVariant Variant { get; init; }
}

public class ExdExporter(bool useHexcode = true, bool includeOffset = true)
{
    private readonly bool _useHexcode = useHexcode;
    private readonly bool _includeOffset = includeOffset;

    [ThreadStatic]
    private static Utf8CsvBuffer? t_singleBuffer;

    public void ExportSheet(
        GameData lumina,
        string sheetName,
        string outputDir,
        Language language,
        ExdSchema? schema,
        Action<CellReadError>? onCellError = null
    )
    {
        var layout = LoadLayout(lumina, sheetName, schema);
        var actualLanguage = ResolveLanguage(layout.HeaderFile, language);

        var buffer = t_singleBuffer ??= new Utf8CsvBuffer();
        buffer.Reset();
        ExportToBuffer(lumina, sheetName, actualLanguage, layout, buffer, onCellError);

        var outputPath = Path.Combine(
            outputDir,
            sheetName.Replace('/', Path.DirectorySeparatorChar) + ".csv"
        );
        StorageHelper.WriteFileDirect(outputPath, buffer.WrittenSpan);
    }

    internal static Language ResolveLanguage(ExcelHeaderFile headerFile, Language preferred)
    {
        var languages = headerFile.Languages;
        if (languages.Contains(Language.None))
            return Language.None;
        if (languages.Contains(preferred))
            return preferred;
        if (languages.Contains(Language.English))
            return Language.English;
        return languages.FirstOrDefault();
    }

    internal SheetLayout LoadLayout(GameData lumina, string sheetName, ExdSchema? schema)
    {
        var headerFile =
            lumina.GetFile<ExcelHeaderFile>($"exd/{sheetName}.exh")
            ?? throw new Exception(Localizer.T(MsgKey.ExhReadFailed));

        var columns = headerFile.ColumnDefinitions;
        int colCount = columns.Length;

        string[]? namesByOrig = null;
        if (schema?.Fields != null && schema.Fields.Count > 0)
        {
            // 按偏移量稳定排序后与 Schema 字段顺序对齐，再还原到原始列索引
            Span<ulong> sortKeys = colCount <= 256 ? stackalloc ulong[colCount] : new ulong[colCount];
            for (int i = 0; i < colCount; i++)
                sortKeys[i] = ((ulong)columns[i].Offset << 32) | (uint)i;
            sortKeys.Sort();

            var expanded = ExpandSchemaFields(schema.Fields);
            namesByOrig = new string[colCount];
            for (int sortedIdx = 0; sortedIdx < colCount; sortedIdx++)
            {
                int origIdx = (int)(uint)sortKeys[sortedIdx];
                namesByOrig[origIdx] =
                    sortedIdx < expanded.Count ? expanded[sortedIdx] : $"Unknown{sortedIdx}";
            }
        }

        var headerBuf = t_singleBuffer ??= new Utf8CsvBuffer();
        headerBuf.Reset();

        headerBuf.WriteBom();
        headerBuf.Write("key"u8);
        for (int i = 0; i < colCount; i++)
        {
            headerBuf.WriteByte((byte)',');
            headerBuf.WriteFormattable(i);
        }
        headerBuf.WriteNewLine();

        headerBuf.WriteByte((byte)'#');
        for (int i = 0; i < colCount; i++)
        {
            headerBuf.WriteByte((byte)',');
            if (namesByOrig != null)
            {
                var name = namesByOrig[i];
                if (!name.StartsWith("Unknown", StringComparison.OrdinalIgnoreCase))
                    headerBuf.WriteCsvEscapedString(name);
            }
        }
        headerBuf.WriteNewLine();

        if (_includeOffset)
        {
            headerBuf.Write("offset"u8);
            for (int i = 0; i < colCount; i++)
            {
                headerBuf.WriteByte((byte)',');
                headerBuf.WriteFormattable(columns[i].Offset);
            }
            headerBuf.WriteNewLine();
        }

        headerBuf.Write("Int32"u8);
        for (int i = 0; i < colCount; i++)
        {
            headerBuf.WriteByte((byte)',');
            headerBuf.WriteColumnType(columns[i].Type);
        }
        headerBuf.WriteNewLine();

        return new SheetLayout
        {
            HeaderFile = headerFile,
            HeaderCsvBytes = headerBuf.WrittenSpan.ToArray(),
            Columns = columns,
            DataOffset = headerFile.Header.DataOffset,
            Variant = headerFile.Header.Variant,
        };
    }

    internal void ExportToBuffer(
        GameData lumina,
        string sheetName,
        Language actualLanguage,
        SheetLayout layout,
        Utf8CsvBuffer buffer,
        Action<CellReadError>? onCellError = null
    )
    {
        buffer.Write(layout.HeaderCsvBytes);

        var columns = layout.Columns;
        int colCount = columns.Length;
        ushort dataOffset = layout.DataOffset;
        bool isSubrow = layout.Variant == ExcelVariant.Subrows;
        string? langStr =
            actualLanguage == Language.None ? null : LanguageUtil.GetLanguageStr(actualLanguage);
        bool skipErrorReport = sheetName is "CustomTalkDefineClient" or "QuestDefineClient";

        foreach (var pageDef in layout.HeaderFile.DataPages)
        {
            string exdPath =
                langStr == null
                    ? $"exd/{sheetName}_{pageDef.StartId}.exd"
                    : $"exd/{sheetName}_{pageDef.StartId}_{langStr}.exd";

            var exdFile = lumina.GetFile<ExcelDataFile>(exdPath);
            if (exdFile == null)
                continue;

            var data = exdFile.Data;
            var span = data.AsSpan();

            foreach (var rowKvp in exdFile.RowData)
            {
                uint rowId = rowKvp.Key;
                int rowDataStart = (int)(rowKvp.Value.Offset + 6);

                if (isSubrow)
                {
                    ushort subrowCount = BinaryPrimitives.ReadUInt16BigEndian(
                        span.Slice((int)(rowKvp.Value.Offset + 4), 2)
                    );
                    for (ushort subrowId = 0; subrowId < subrowCount; subrowId++)
                    {
                        buffer.WriteFormattable(rowId);
                        buffer.WriteByte((byte)'.');
                        buffer.WriteFormattable(subrowId);

                        int subrowDataStart = rowDataStart + subrowId * (dataOffset + 2) + 2;
                        for (int c = 0; c < colCount; c++)
                        {
                            buffer.WriteByte((byte)',');
                            WriteColumnValue(
                                buffer,
                                sheetName,
                                rowId,
                                subrowId,
                                span,
                                data,
                                subrowDataStart,
                                columns[c],
                                dataOffset,
                                skipErrorReport,
                                onCellError
                            );
                        }
                        buffer.WriteNewLine();
                    }
                }
                else
                {
                    buffer.WriteFormattable(rowId);

                    for (int c = 0; c < colCount; c++)
                    {
                        buffer.WriteByte((byte)',');
                        WriteColumnValue(
                            buffer,
                            sheetName,
                            rowId,
                            -1,
                            span,
                            data,
                            rowDataStart,
                            columns[c],
                            dataOffset,
                            skipErrorReport,
                            onCellError
                        );
                    }
                    buffer.WriteNewLine();
                }
            }
        }
    }

    private static List<string> ExpandSchemaFields(List<SchemaField> fields)
    {
        var expanded = new List<string>(fields.Count);
        foreach (var field in fields)
        {
            if (field.Type == "array" && field.Count > 0)
            {
                if (field.Fields != null && field.Fields.Count > 0)
                {
                    for (int i = 0; i < field.Count; i++)
                    {
                        for (int j = 0; j < field.Fields.Count; j++)
                        {
                            var subField = field.Fields[j];
                            expanded.Add(
                                field.Fields.Count == 1
                                    ? $"{field.Name}[{i}]"
                                    : $"{field.Name}[{i}].{subField.Name ?? $"Field{j}"}"
                            );
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < field.Count; i++)
                        expanded.Add($"{field.Name}[{i}]");
                }
            }
            else
            {
                expanded.Add(field.Name ?? $"Unknown{expanded.Count}");
            }
        }
        return expanded;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void WriteColumnValue(
        Utf8CsvBuffer buffer,
        string sheetName,
        uint rowId,
        int subrowId,
        ReadOnlySpan<byte> span,
        byte[] data,
        int rowDataStart,
        ExcelColumnDefinition column,
        ushort headerDataOffset,
        bool skipErrorReport,
        Action<CellReadError>? onCellError
    )
    {
        int offset = rowDataStart + column.Offset;
        try
        {
            switch (column.Type)
            {
                case ExcelColumnDataType.String:
                {
                    int stringOffset = BinaryPrimitives.ReadInt32BigEndian(span.Slice(offset, 4));
                    if (stringOffset < 0)
                    {
                        buffer.Write("\"\""u8);
                        return;
                    }

                    int absoluteOffset = rowDataStart + headerDataOffset + stringOffset;
                    if ((uint)absoluteOffset >= (uint)span.Length)
                    {
                        buffer.Write("\"\""u8);
                        return;
                    }

                    int length = span[absoluteOffset..].IndexOf((byte)0);
                    if (length < 0)
                        length = span.Length - absoluteOffset;
                    if (length == 0)
                    {
                        buffer.Write("\"\""u8);
                        return;
                    }

                    if (_useHexcode)
                        SeStringDecoder.WriteHexcodeCsvField(buffer, data, absoluteOffset, length);
                    else
                        SeStringDecoder.WriteReadableCsvField(buffer, data, absoluteOffset, length);
                    return;
                }
                case ExcelColumnDataType.Bool:
                    buffer.Write(span[offset] != 0 ? "True"u8 : "False"u8);
                    return;
                case ExcelColumnDataType.Int8:
                    buffer.WriteFormattable((sbyte)span[offset]);
                    return;
                case ExcelColumnDataType.UInt8:
                    buffer.WriteFormattable(span[offset]);
                    return;
                case ExcelColumnDataType.Int16:
                    buffer.WriteFormattable(BinaryPrimitives.ReadInt16BigEndian(span.Slice(offset, 2)));
                    return;
                case ExcelColumnDataType.UInt16:
                    buffer.WriteFormattable(BinaryPrimitives.ReadUInt16BigEndian(span.Slice(offset, 2)));
                    return;
                case ExcelColumnDataType.Int32:
                    buffer.WriteFormattable(BinaryPrimitives.ReadInt32BigEndian(span.Slice(offset, 4)));
                    return;
                case ExcelColumnDataType.UInt32:
                    buffer.WriteFormattable(BinaryPrimitives.ReadUInt32BigEndian(span.Slice(offset, 4)));
                    return;
                case ExcelColumnDataType.Float32:
                    buffer.WriteFormattable(
                        BitConverter.Int32BitsToSingle(
                            BinaryPrimitives.ReadInt32BigEndian(span.Slice(offset, 4))
                        )
                    );
                    return;
                case ExcelColumnDataType.Int64:
                    buffer.WriteFormattable(BinaryPrimitives.ReadInt64BigEndian(span.Slice(offset, 8)));
                    return;
                case ExcelColumnDataType.UInt64:
                    buffer.WriteFormattable(BinaryPrimitives.ReadUInt64BigEndian(span.Slice(offset, 8)));
                    return;
                case >= ExcelColumnDataType.PackedBool0 and <= ExcelColumnDataType.PackedBool7:
                    buffer.Write(
                        (span[offset] & (1 << (column.Type - ExcelColumnDataType.PackedBool0))) != 0
                            ? "True"u8
                            : "False"u8
                    );
                    return;
                default:
                    buffer.WriteString($"<unknown:{column.Type}>");
                    return;
            }
        }
        catch (Exception ex)
        {
            if (skipErrorReport)
                return;

            string fullRowId = subrowId >= 0 ? $"{rowId}.{subrowId}" : rowId.ToString();
            onCellError?.Invoke(
                new CellReadError(
                    sheetName,
                    fullRowId,
                    column.Offset,
                    column.Type.ToString(),
                    ex.Message
                )
            );
        }
    }
}
