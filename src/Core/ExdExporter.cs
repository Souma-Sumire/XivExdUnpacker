using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using Lumina;
using Lumina.Data;
using Lumina.Data.Files;
using Lumina.Data.Files.Excel;
using Lumina.Data.Structs.Excel;
using XivExdUnpacker.Decoders;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;

namespace XivExdUnpacker.src.Core;

public record CellReadError(
    string SheetName,
    string RowId,
    ushort ColumnOffset,
    string ColumnType,
    string Message
);

public class ExdExporter(bool useHexcode = true, bool includeOffset = true)
{
    private const int BufferSize = 131072;
    private readonly bool _useHexcode = useHexcode;
    private readonly bool _includeOffset = includeOffset;

    private static readonly ConcurrentDictionary<string, byte> _createdDirs = new();

    private static readonly Encoding _utf8WithBom = Encoding.UTF8;

    public void ExportSheet(
        GameData lumina,
        string sheetName,
        string outputDir,
        Language language,
        ExdSchema? schema,
        Action<CellReadError>? onCellError = null
    )
    {
        var headerFile =
            lumina.GetFile<ExcelHeaderFile>($"exd/{sheetName}.exh")
            ?? throw new Exception(Localizer.T(MsgKey.ExhReadFailed));
        Language actualLanguage =
            headerFile.Languages.Contains(Language.None) ? Language.None
            : headerFile.Languages.Contains(language) ? language
            : headerFile.Languages.Contains(Language.English) ? Language.English
            : headerFile.Languages.FirstOrDefault();

        ExportSheetWithLanguage(
            lumina,
            sheetName,
            headerFile,
            outputDir,
            actualLanguage,
            schema,
            onCellError
        );
    }

    private void ExportSheetWithLanguage(
        GameData lumina,
        string sheetName,
        ExcelHeaderFile headerFile,
        string outputDir,
        Language language,
        ExdSchema? schema,
        Action<CellReadError>? onCellError
    )
    {
        var header = headerFile.Header;
        var columns = headerFile.ColumnDefinitions;

        var fileName = sheetName.Replace('/', Path.DirectorySeparatorChar);
        var outputPath = Path.Combine(outputDir, fileName + ".csv");
        var outputFileDir = Path.GetDirectoryName(outputPath);

        if (!string.IsNullOrEmpty(outputFileDir) && !_createdDirs.ContainsKey(outputFileDir))
        {
            Directory.CreateDirectory(outputFileDir);
            _createdDirs.TryAdd(outputFileDir, 0);
        }

        using var writer = new StreamWriter(outputPath, false, _utf8WithBom, BufferSize);

        var indexedColumns = columns
            .Select((c, i) => new { Definition = c, OriginalIndex = i })
            .ToList();
        var sortedForSchema = indexedColumns.OrderBy(x => x.Definition.Offset).ToList();

        var (sortedNames, _) = GenerateColumnNamesAndTypes(
            [.. sortedForSchema.Select(x => x.Definition)],
            schema
        );

        var columnInfos = new List<ColumnInfo>();
        for (int i = 0; i < sortedForSchema.Count; i++)
        {
            var def = sortedForSchema[i].Definition;
            var originalIndex = sortedForSchema[i].OriginalIndex;
            string name;
            bool isUnknown;

            if (schema?.Fields == null || schema.Fields.Count == 0)
            {
                name = originalIndex.ToString();
                isUnknown = true;
            }
            else
            {
                name = i < sortedNames.Count ? sortedNames[i] : $"Unknown{originalIndex}";
                isUnknown = name.StartsWith("Unknown", StringComparison.OrdinalIgnoreCase);
            }

            columnInfos.Add(
                new ColumnInfo
                {
                    Definition = def,
                    OriginalIndex = originalIndex,
                    Name = name,
                    Type = def.Type.ToString(),
                    IsUnknown = isUnknown,
                }
            );
        }

        var finalColumns = columnInfos.OrderBy(x => x.OriginalIndex).ToList();
        bool isSubrow = header.Variant == ExcelVariant.Subrows;

        var keyIndexRow = new List<string> { "key" };
        for (int i = 0; i < finalColumns.Count; i++)
            keyIndexRow.Add(i.ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine(string.Join(",", keyIndexRow.Select(EscapeCsv)));

        var headerRow = new List<string> { "#" };
        foreach (var col in finalColumns)
            headerRow.Add(col.IsUnknown ? "" : col.Name);
        writer.WriteLine(string.Join(",", headerRow.Select(EscapeCsv)));

        if (_includeOffset)
        {
            var offsetRow = new List<string> { "offset" };
            foreach (var col in finalColumns)
                offsetRow.Add(
                    col.Definition.Offset.ToString(
                        System.Globalization.CultureInfo.InvariantCulture
                    )
                );
            writer.WriteLine(string.Join(",", offsetRow.Select(EscapeCsv)));
        }

        var typeRow = new List<string> { "Int32" };
        foreach (var col in finalColumns)
            typeRow.Add(col.Type);
        writer.WriteLine(string.Join(",", typeRow.Select(EscapeCsv)));

        foreach (var pageDef in headerFile.DataPages)
        {
            string exdPath =
                language == Language.None
                    ? $"exd/{sheetName}_{pageDef.StartId}.exd"
                    : $"exd/{sheetName}_{pageDef.StartId}_{LanguageUtil.GetLanguageStr(language)}.exd";

            var exdFile = lumina.GetFile<ExcelDataFile>(exdPath);
            if (exdFile == null)
                continue;

            foreach (var rowKvp in exdFile.RowData)
            {
                var rowId = rowKvp.Key;
                var rowPtr = rowKvp.Value;
                var rowDataStart = (int)(rowPtr.Offset + 6);

                if (isSubrow)
                {
                    var subrowCount = BinaryPrimitives.ReadUInt16BigEndian(
                        exdFile.Data.AsSpan((int)(rowPtr.Offset + 4), 2)
                    );
                    for (ushort subrowId = 0; subrowId < subrowCount; subrowId++)
                    {
                        writer.Write(
                            EscapeCsv(
                                string.Create(
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    $"{rowId}.{subrowId}"
                                )
                            )
                        );

                        var subrowDataStart = rowDataStart + subrowId * (header.DataOffset + 2) + 2;
                        var fullRowId = $"{rowId}.{subrowId}";

                        foreach (var col in finalColumns)
                        {
                            writer.Write(',');
                            WriteColumnValue(
                                writer,
                                sheetName,
                                fullRowId,
                                exdFile.Data,
                                subrowDataStart,
                                col.Definition,
                                header.DataOffset,
                                onCellError
                            );
                        }
                        writer.WriteLine();
                    }
                }
                else
                {
                    writer.Write(rowId);

                    foreach (var col in finalColumns)
                    {
                        writer.Write(',');
                        WriteColumnValue(
                            writer,
                            sheetName,
                            rowId.ToString(),
                            exdFile.Data,
                            rowDataStart,
                            col.Definition,
                            header.DataOffset,
                            onCellError
                        );
                    }
                    writer.WriteLine();
                }
            }
        }
    }

    private static void WriteQuotedCsvString(StreamWriter writer, string value)
    {
        writer.Write('"');
        int start = 0;
        while (start < value.Length)
        {
            int quoteIdx = value.IndexOf('"', start);
            if (quoteIdx < 0)
            {
                writer.Write(value.AsSpan(start));
                break;
            }
            if (quoteIdx > start)
                writer.Write(value.AsSpan(start, quoteIdx - start));
            writer.Write("\"\"");
            start = quoteIdx + 1;
        }
        writer.Write('"');
    }

    private static (List<string> names, List<string> types) GenerateColumnNamesAndTypes(
        ExcelColumnDefinition[] columns,
        ExdSchema? schema
    )
    {
        var names = new List<string>();
        if (schema?.Fields == null || schema.Fields.Count == 0)
            return ([], []);

        var expandedFields = new List<string>();
        foreach (var field in schema.Fields)
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
                            expandedFields.Add(
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
                        expandedFields.Add($"{field.Name}[{i}]");
                }
            }
            else
                expandedFields.Add(field.Name ?? $"Unknown{expandedFields.Count}");
        }

        for (int i = 0; i < columns.Length; i++)
            names.Add(i < expandedFields.Count ? expandedFields[i] : $"Unknown{i}");
        return (names, []);
    }

    private void WriteColumnValue(
        StreamWriter writer,
        string sheetName,
        string rowId,
        byte[] data,
        int rowDataStart,
        ExcelColumnDefinition column,
        ushort headerDataOffset,
        Action<CellReadError>? onCellError
    )
    {
        var offset = rowDataStart + column.Offset;
        try
        {
            var span = data.AsSpan();
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            Span<char> numBuf = stackalloc char[64];

            switch (column.Type)
            {
                case ExcelColumnDataType.String:
                    WriteQuotedCsvString(
                        writer,
                        ReadString(data, offset, rowDataStart, headerDataOffset)
                    );
                    break;
                case ExcelColumnDataType.Bool:
                    writer.Write(span[offset] != 0 ? "True" : "False");
                    break;
                case ExcelColumnDataType.Int8:
                    ((sbyte)span[offset]).TryFormat(numBuf, out int w1, default, culture);
                    writer.Write(numBuf[..w1]);
                    break;
                case ExcelColumnDataType.UInt8:
                    span[offset].TryFormat(numBuf, out int w2, default, culture);
                    writer.Write(numBuf[..w2]);
                    break;
                case ExcelColumnDataType.Int16:
                    BinaryPrimitives
                        .ReadInt16BigEndian(span.Slice(offset, 2))
                        .TryFormat(numBuf, out int w3, default, culture);
                    writer.Write(numBuf[..w3]);
                    break;
                case ExcelColumnDataType.UInt16:
                    BinaryPrimitives
                        .ReadUInt16BigEndian(span.Slice(offset, 2))
                        .TryFormat(numBuf, out int w4, default, culture);
                    writer.Write(numBuf[..w4]);
                    break;
                case ExcelColumnDataType.Int32:
                    BinaryPrimitives
                        .ReadInt32BigEndian(span.Slice(offset, 4))
                        .TryFormat(numBuf, out int w5, default, culture);
                    writer.Write(numBuf[..w5]);
                    break;
                case ExcelColumnDataType.UInt32:
                    BinaryPrimitives
                        .ReadUInt32BigEndian(span.Slice(offset, 4))
                        .TryFormat(numBuf, out int w6, default, culture);
                    writer.Write(numBuf[..w6]);
                    break;
                case ExcelColumnDataType.Float32:
                    BitConverter
                        .Int32BitsToSingle(BinaryPrimitives.ReadInt32BigEndian(span.Slice(offset, 4)))
                        .TryFormat(numBuf, out int w7, default, culture);
                    writer.Write(numBuf[..w7]);
                    break;
                case ExcelColumnDataType.Int64:
                    BinaryPrimitives
                        .ReadInt64BigEndian(span.Slice(offset, 8))
                        .TryFormat(numBuf, out int w8, default, culture);
                    writer.Write(numBuf[..w8]);
                    break;
                case ExcelColumnDataType.UInt64:
                    BinaryPrimitives
                        .ReadUInt64BigEndian(span.Slice(offset, 8))
                        .TryFormat(numBuf, out int w9, default, culture);
                    writer.Write(numBuf[..w9]);
                    break;
                case >= ExcelColumnDataType.PackedBool0 and <= ExcelColumnDataType.PackedBool7:
                    writer.Write(
                        (span[offset] & (1 << (column.Type - ExcelColumnDataType.PackedBool0))) != 0
                            ? "True"
                            : "False"
                    );
                    break;
                default:
                    writer.Write($"<unknown:{column.Type}>");
                    break;
            }
        }
        catch (Exception ex)
        {
            if (sheetName == "CustomTalkDefineClient" || sheetName == "QuestDefineClient")
                return;

            onCellError?.Invoke(
                new CellReadError(
                    sheetName,
                    rowId,
                    column.Offset,
                    column.Type.ToString(),
                    ex.Message
                )
            );
        }
    }

    private string ReadString(byte[] data, int offset, int rowDataStart, ushort headerDataOffset)
    {
        var stringOffset = BinaryPrimitives.ReadInt32BigEndian(data.AsSpan(offset, 4));
        if (stringOffset < 0)
            return string.Empty;

        var absoluteOffset = rowDataStart + headerDataOffset + stringOffset;
        if (absoluteOffset < 0 || absoluteOffset >= data.Length)
            return string.Empty;

        var length = 0;
        while (absoluteOffset + length < data.Length && data[absoluteOffset + length] != 0)
            length++;
        var stringMemory = data.AsMemory(absoluteOffset, length);

        return _useHexcode
            ? SeStringDecoder.DecodeHexcode(stringMemory)
            : SeStringDecoder.DecodeReadable(stringMemory);
    }

    private string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        bool needsQuotes = false;
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (c == ',' || c == '"' || c == '\n' || c == '\r')
            {
                needsQuotes = true;
                break;
            }
        }

        if (!needsQuotes)
            return value;

        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
