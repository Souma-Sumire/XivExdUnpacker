using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using Lumina;
using Lumina.Data;
using Lumina.Data.Files;
using Lumina.Data.Files.Excel;
using Lumina.Data.Structs.Excel;
using SaintCoinach.Text;
using XivExdUnpacker.Models;

namespace XivExdUnpacker.Core;

public class ExdExporter
{
    private const int BufferSize = 65536; // 64KB 缓冲区

    // 使用 ThreadLocal 缓存解码器，既保证了线程安全，又避免了每个单元格都 new 的海量分配
    private static readonly ThreadLocal<XivStringDecoder> _threadDecoder = new(() =>
        new XivStringDecoder()
    );

    // 全局静态目录缓存, 避免多个语言重复调用 Directory.Exists/CreateDirectory
    private static readonly ConcurrentDictionary<string, byte> _createdDirs = new();

    // 静态编码实例
    private static readonly UTF8Encoding _utf8WithBom = new(true);

    public void ExportSheet(
        GameData lumina,
        string sheetName,
        string outputDir,
        Language language,
        ExdSchema? schema
    )
    {
        var headerFile = lumina.GetFile<ExcelHeaderFile>($"exd/{sheetName}.exh");
        if (headerFile == null)
            throw new Exception("无法读取表头文件");

        Language actualLanguage =
            headerFile.Languages.Contains(Language.None) ? Language.None
            : headerFile.Languages.Contains(language) ? language
            : headerFile.Languages.Contains(Language.English) ? Language.English
            : headerFile.Languages.FirstOrDefault();

        ExportSheetWithLanguage(lumina, sheetName, headerFile, outputDir, actualLanguage, schema);
    }

    private void ExportSheetWithLanguage(
        GameData lumina,
        string sheetName,
        ExcelHeaderFile headerFile,
        string outputDir,
        Language language,
        ExdSchema? schema
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

        using var writer = new StreamWriter(outputPath, false, _utf8WithBom, 65536);

        var indexedColumns = columns
            .Select((c, i) => new { Definition = c, OriginalIndex = i })
            .ToList();
        var sortedForSchema = indexedColumns.OrderBy(x => x.Definition.Offset).ToList();

        var (sortedNames, _) = GenerateColumnNamesAndTypes(
            sortedForSchema.Select(x => x.Definition).ToArray(),
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

        // Header Rows
        var keyIndexRow = new List<string> { "key" };
        for (int i = 0; i < finalColumns.Count; i++)
            keyIndexRow.Add(i.ToString());
        writer.WriteLine(string.Join(",", keyIndexRow.Select(EscapeCsv)));

        var headerRow = new List<string> { "#" };
        foreach (var col in finalColumns)
            headerRow.Add(col.IsUnknown ? "" : col.Name);
        writer.WriteLine(string.Join(",", headerRow.Select(EscapeCsv)));

        var offsetRow = new List<string> { "offset" };
        foreach (var col in finalColumns)
            offsetRow.Add(col.Definition.Offset.ToString());
        writer.WriteLine(string.Join(",", offsetRow.Select(EscapeCsv)));

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
                        // 写入 ID 列
                        writer.Write(EscapeCsv($"{rowId}.{subrowId}"));

                        var subrowDataStart = rowDataStart + subrowId * (header.DataOffset + 2) + 2;
                        var fullRowId = $"{rowId}.{subrowId}";

                        foreach (var col in finalColumns)
                        {
                            writer.Write(',');
                            var value = ReadColumnValue(
                                sheetName,
                                fullRowId,
                                exdFile.Data,
                                subrowDataStart,
                                col.Definition,
                                header.DataOffset
                            );
                            WriteCsvValue(writer, value, col.Definition.Type);
                        }
                        writer.WriteLine();
                    }
                }
                else
                {
                    // 写入 ID 列
                    writer.Write(EscapeCsv(rowId.ToString()));

                    foreach (var col in finalColumns)
                    {
                        writer.Write(',');
                        var value = ReadColumnValue(
                            sheetName,
                            rowId,
                            exdFile.Data,
                            rowDataStart,
                            col.Definition,
                            header.DataOffset
                        );
                        WriteCsvValue(writer, value, col.Definition.Type);
                    }
                    writer.WriteLine();
                }
            }
        }
    }

    private void WriteCsvValue(StreamWriter writer, string value, ExcelColumnDataType type)
    {
        if (type == ExcelColumnDataType.String)
        {
            writer.Write('"');
            // 手动转义，不使用 Replace 分配新字符串
            foreach (char c in value)
            {
                if (c == '"')
                    writer.Write("\"\"");
                else
                    writer.Write(c);
            }
            writer.Write('"');
        }
        else
        {
            // 对于数值类型，EscapeCsv 是安全的
            writer.Write(EscapeCsv(value));
        }
    }

    private (List<string> names, List<string> types) GenerateColumnNamesAndTypes(
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

    private string ReadColumnValue(
        string sheetName,
        object rowId,
        byte[] data,
        int rowDataStart,
        ExcelColumnDefinition column,
        ushort headerDataOffset
    )
    {
        var offset = rowDataStart + column.Offset;
        try
        {
            var span = data.AsSpan();
            return column.Type switch
            {
                ExcelColumnDataType.String => ReadString(
                    data,
                    offset,
                    rowDataStart,
                    headerDataOffset
                ),
                ExcelColumnDataType.Bool => (span[offset] != 0).ToString(),
                ExcelColumnDataType.Int8 => ((sbyte)span[offset]).ToString(),
                ExcelColumnDataType.UInt8 => span[offset].ToString(),
                ExcelColumnDataType.Int16 => BinaryPrimitives
                    .ReadInt16BigEndian(span.Slice(offset, 2))
                    .ToString(),
                ExcelColumnDataType.UInt16 => BinaryPrimitives
                    .ReadUInt16BigEndian(span.Slice(offset, 2))
                    .ToString(),
                ExcelColumnDataType.Int32 => BinaryPrimitives
                    .ReadInt32BigEndian(span.Slice(offset, 4))
                    .ToString(),
                ExcelColumnDataType.UInt32 => BinaryPrimitives
                    .ReadUInt32BigEndian(span.Slice(offset, 4))
                    .ToString(),
                ExcelColumnDataType.Float32 => BitConverter
                    .Int32BitsToSingle(BinaryPrimitives.ReadInt32BigEndian(span.Slice(offset, 4)))
                    .ToString(),
                ExcelColumnDataType.Int64 => BinaryPrimitives
                    .ReadInt64BigEndian(span.Slice(offset, 8))
                    .ToString(),
                ExcelColumnDataType.UInt64 => BinaryPrimitives
                    .ReadUInt64BigEndian(span.Slice(offset, 8))
                    .ToString(),

                // Packed bool 类型 (位域)
                >= ExcelColumnDataType.PackedBool0 and <= ExcelColumnDataType.PackedBool7 => (
                    (span[offset] & (1 << (column.Type - ExcelColumnDataType.PackedBool0))) != 0
                ).ToString(),

                _ => $"<unknown:{column.Type}>",
            };
        }
        catch (Exception ex)
        {
            // 屏蔽已知存在 Schema 损坏或结构异常的表，避免错误日志刷屏
            if (sheetName == "CustomTalkDefineClient" || sheetName == "QuestDefineClient")
                return "";

            // 着色打印具体的定位信息
            lock (Console.Out)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write($"[Error] ");
                Console.ResetColor();
                Console.WriteLine(
                    $"读取失败 @ {sheetName} -> 行:{rowId} | 列偏移:0x{column.Offset:X} | 类型:{column.Type} | 实际地址:0x{offset:X}"
                );

                // 打印失败位置的数据
                string hexDump;
                try
                {
                    int dumpSize = Math.Min(16, data.Length - offset);
                    if (offset >= 0 && offset < data.Length && dumpSize > 0)
                        hexDump = BitConverter.ToString(data, offset, dumpSize).Replace("-", " ");
                    else
                        hexDump = "N/A (越界)";
                }
                catch
                {
                    hexDump = "获取失败";
                }

                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine($"        错误信息: {ex.Message}");
                Console.WriteLine($"        原始数据: [{hexDump}]");
                Console.ResetColor();
            }
            return "";
        }
    }

    private string ReadString(byte[] data, int offset, int rowDataStart, ushort headerDataOffset)
    {
        var stringOffset = BinaryPrimitives.ReadInt32BigEndian(data.AsSpan(offset, 4));
        var absoluteOffset = rowDataStart + headerDataOffset + stringOffset;
        var length = 0;
        while (absoluteOffset + length < data.Length && data[absoluteOffset + length] != 0)
            length++;
        var stringData = data.AsSpan(absoluteOffset, length).ToArray();
        return _threadDecoder.Value!.Decode(stringData).ToString();
    }

    private string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        // 快速扫描：如果没有特殊字符，直接返回原始引用
        bool needsQuotes = false;
        foreach (char c in value)
        {
            if (c == ',' || c == '"' || c == '\n' || c == '\r')
            {
                needsQuotes = true;
                break;
            }
        }

        if (!needsQuotes)
            return value;

        // 只有在确定需要转义时才进行替换和加引号
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
