using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Lumina;
using Lumina.Data;
using Lumina.Data.Files;
using Lumina.Data.Files.Excel;
using Lumina.Data.Structs.Excel;
using Lumina.Excel;
using SaintCoinach.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace XivExdUnpacker;

class Program
{
    // 类型简称映射表
    private static readonly Dictionary<string, string> TypeMapping = new()
    {
        { "String", "str" },
        { "Bool", "bool" },
        { "Int8", "sbyte" },
        { "UInt8", "byte" },
        { "Int16", "int16" },
        { "UInt16", "uint16" },
        { "Int32", "int32" },
        { "UInt32", "uint32" },
        { "Int64", "int64" },
        { "UInt64", "uint64" },
        { "Float32", "single" },
        { "PackedBool0", "bit&01" },
        { "PackedBool1", "bit&02" },
        { "PackedBool2", "bit&04" },
        { "PackedBool3", "bit&08" },
        { "PackedBool4", "bit&10" },
        { "PackedBool5", "bit&20" },
        { "PackedBool6", "bit&40" },
        { "PackedBool7", "bit&80" },
    };

    private static string GetMappedType(string originalType)
    {
        return TypeMapping.TryGetValue(originalType, out var mapped) ? mapped : originalType;
    }

    // 配置名称到语言的固定映射
    private static readonly Dictionary<string, Language> KeyToLanguage = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        { "en", Language.English },
        { "ja", Language.Japanese },
        { "de", Language.German },
        { "fr", Language.French },
        { "cn", Language.ChineseSimplified },
        { "ko", Language.Korean },
        { "tc", Language.ChineseTraditional },
    };

    // ======== 动态配置类 ========
    public class AppConfig
    {
        public ClientConfig? Global { get; set; }
        public Dictionary<string, ClientConfig> Clients { get; set; } = new();
    }

    public class ClientConfig
    {
        public string? Path { get; set; }
        public string? OutputDir { get; set; }
        public string? SchemaVersion { get; set; }
    }

    private static string gamePath = "";
    private static string outputDir = "";
    private static string schemaVersion = "latest";
    private static string schemaDir = "";
    private static Language exportLanguage = Language.ChineseSimplified;

    // ============================

    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.CursorVisible = false;
        Console.WriteLine("=== XivExd Unpacker (基于 Lumina) ===");

        // 1. 加载配置文件
        AppConfig config = LoadAppConfig();

        if (config.Clients.Count == 0)
        {
            Console.WriteLine("错误: 配置文件中未定义任何客户端。");
            return;
        }

        List<string>? selectedKeys = new List<string>();

        // 2. 确定要处理的客户端
        if (args.Length > 0)
        {
            if (args[0].ToLower() == "all")
                selectedKeys.AddRange(config.Clients.Keys);
            else
                selectedKeys.Add(args[0]);
        }
        else
        {
            // 模式 B: Vue-CLI 风格交互菜单
            selectedKeys = ShowInteractiveMenu(config.Clients.Keys.ToList());
        }

        if (selectedKeys == null || selectedKeys.Count == 0)
        {
            Console.WriteLine("任务已取消或未选择任何客户端。");
            Console.CursorVisible = true;
            return;
        }

        Console.CursorVisible = true;
        // 3. 执行导出
        foreach (var key in selectedKeys)
        {
            Console.WriteLine($"\n>>> 正在开始处理客户端: {key} <<<");
            RunDumpProcess(key, config);
        }

        Console.WriteLine("\n所有选定任务处理完毕。");
    }

    static List<string>? ShowInteractiveMenu(List<string> options)
    {
        int currentPos = 0;
        bool[] selected = new bool[options.Count];
        bool done = false;

        Console.WriteLine(
            "\n使用 [上下方向键] 移动, [空格] 勾选/取消, [A] 全选, [回车] 确定, [Esc] 退出\n"
        );

        int startLine = Console.CursorTop;
        int maxBuffer = Console.BufferHeight;

        // 如果菜单可能超出缓冲区，调整起点
        if (startLine + options.Count >= maxBuffer)
        {
            startLine = maxBuffer - options.Count - 1;
            if (startLine < 0)
                startLine = 0;
        }

        while (!done)
        {
            for (int i = 0; i < options.Count; i++)
            {
                int targetLine = startLine + i;
                if (targetLine < 0 || targetLine >= maxBuffer)
                    continue;

                try
                {
                    Console.SetCursorPosition(0, targetLine);
                }
                catch
                {
                    continue;
                }
                string cursor = (i == currentPos) ? ">" : " ";
                string checkbox = selected[i] ? "[*]" : "[ ]";

                if (i == currentPos)
                {
                    Console.BackgroundColor = ConsoleColor.DarkGray;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                }

                Console.Write(
                    $"{cursor} {checkbox} {options[i]}".PadRight(Console.WindowWidth - 1)
                );
                Console.ResetColor();
            }

            var key = Console.ReadKey(true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    currentPos = (currentPos - 1 + options.Count) % options.Count;
                    break;
                case ConsoleKey.DownArrow:
                    currentPos = (currentPos + 1) % options.Count;
                    break;
                case ConsoleKey.Spacebar:
                    selected[currentPos] = !selected[currentPos];
                    break;
                case ConsoleKey.A:
                    bool allSelected = selected.All(x => x);
                    for (int j = 0; j < selected.Length; j++)
                        selected[j] = !allSelected;
                    break;
                case ConsoleKey.Enter:
                    done = true;
                    break;
                case ConsoleKey.Escape:
                    return null;
            }
        }

        try
        {
            Console.SetCursorPosition(0, Math.Min(startLine + options.Count, maxBuffer - 1));
        }
        catch { }
        Console.WriteLine();
        return options.Where((t, i) => selected[i]).ToList();
    }

    static void RunDumpProcess(string clientKey, AppConfig config)
    {
        // 1. 获取配置
        config.Clients.TryGetValue(clientKey, out var client);

        // 2. 确定最终参数
        gamePath = client?.Path ?? config.Global?.Path ?? "";
        outputDir = client?.OutputDir ?? config.Global?.OutputDir ?? "./rawexd";
        schemaVersion = client?.SchemaVersion ?? config.Global?.SchemaVersion ?? "latest";

        // 3. 根据客户端 key 自动选择语言
        if (!KeyToLanguage.TryGetValue(clientKey, out exportLanguage))
        {
            exportLanguage = Language.English;
        }

        schemaDir = Path.Combine("./EXDSchema/schemas", schemaVersion);

        Console.WriteLine($"[配置]: {clientKey}");
        Console.WriteLine($"[路径]: {gamePath}");
        Console.WriteLine($"[语言]: {exportLanguage}");
        Console.WriteLine($"[Schema]: {schemaVersion}");

        if (!Directory.Exists(gamePath))
        {
            Console.WriteLine($"警告: 游戏路径不存在，跳过任务: {gamePath}");
            return;
        }

        // 加载 Schema
        Dictionary<string, ExdSchema>? schemas = null;
        if (Directory.Exists(schemaDir))
        {
            schemas = LoadSchemas(schemaDir);
            Console.WriteLine($"已加载 {schemas.Count} 个 Schema");
        }

        // 创建输出目录
        Directory.CreateDirectory(outputDir);

        try
        {
            if (Directory.Exists(outputDir))
            {
                Console.WriteLine($"正在清空输出目录: {outputDir} ...");
                try
                {
                    var dir = new DirectoryInfo(outputDir);
                    foreach (var file in dir.GetFiles())
                        file.Delete();
                    foreach (var subDir in dir.GetDirectories())
                        subDir.Delete(true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"警告: 清空目录失败 ({ex.Message})，导出将继续...");
                }
            }

            Console.WriteLine("正在初始化 Lumina...");
            var lumina = new GameData(
                gamePath,
                new LuminaOptions { DefaultExcelLanguage = exportLanguage }
            );

            Console.WriteLine("正在获取所有 Excel 表列表...");

            // 获取所有可用的 sheet 名称
            var sheetNames = lumina.Excel.SheetNames.ToList();
            Console.WriteLine($"找到 {sheetNames.Count} 个表");
            Console.WriteLine();

            int successCount = 0;
            int schemaCount = 0;
            var failedSheets = new ConcurrentBag<(string name, string error)>();
            int processedCount = 0;
            int totalSheets = sheetNames.Count;
            object consoleConsoleLock = new object();

            Console.WriteLine("开始并行导出...");

            Parallel.ForEach(
                sheetNames,
                new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                sheetName =>
                {
                    try
                    {
                        // 尝试获取 schema
                        ExdSchema? schema = null;
                        if (schemas != null)
                        {
                            var baseSheetName = sheetName.Contains('/')
                                ? sheetName.Substring(sheetName.LastIndexOf('/') + 1)
                                : sheetName;
                            schemas.TryGetValue(baseSheetName, out schema);
                        }

                        ExportSheet(lumina, sheetName, outputDir, exportLanguage, schema);
                        Interlocked.Increment(ref successCount);

                        if (schema != null)
                        {
                            Interlocked.Increment(ref schemaCount);
                        }
                    }
                    catch (Exception ex)
                    {
                        failedSheets.Add((sheetName, ex.Message));
                    }
                    finally
                    {
                        int current = Interlocked.Increment(ref processedCount);
                        // 每 10 个或完成时刷新
                        if (current % 10 == 0 || current == totalSheets)
                        {
                            lock (consoleConsoleLock)
                            {
                                Console.Write(
                                    $"\r进度: {current}/{totalSheets} ({(double)current / totalSheets:P0})   "
                                );
                            }
                        }
                    }
                }
            );

            Console.WriteLine(); // 进度条换行

            Console.WriteLine("=== 导出完成 ===");
            Console.WriteLine($"总计表数: {sheetNames.Count}");
            Console.WriteLine($"成功导出: {successCount} (其中 {schemaCount} 个包含 Schema 定义)");

            Console.WriteLine($"输出目录: {outputDir}");

            if (failedSheets.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("失败的表:");
                foreach (var (name, error) in failedSheets.Take(20))
                {
                    Console.WriteLine($"  - {name}: {error}");
                }
                if (failedSheets.Count > 20)
                {
                    Console.WriteLine($"  ... 还有 {failedSheets.Count - 20} 个");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"初始化失败: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    /// <summary>
    /// 加载全局配置
    /// </summary>
    static AppConfig LoadAppConfig()
    {
        // 优先检查当前工作目录
        var configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.yml");
        if (!File.Exists(configPath))
        {
            // 备选检查程序运行目录
            configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.yml");
        }

        if (!File.Exists(configPath))
        {
            Console.WriteLine("警告: 未找到 config.yml，使用默认配置。");
            return new AppConfig();
        }

        try
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();

            var yaml = File.ReadAllText(configPath);
            return deserializer.Deserialize<AppConfig>(yaml) ?? new AppConfig();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"无法加载 config.yml: {ex.Message}。将使用默认配置。");
            return new AppConfig();
        }
    }

    /// <summary>
    /// 加载所有 Schema 文件
    /// </summary>
    static Dictionary<string, ExdSchema> LoadSchemas(string schemaDir)
    {
        var schemas = new Dictionary<string, ExdSchema>(StringComparer.OrdinalIgnoreCase);
        var deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();

        foreach (var file in Directory.GetFiles(schemaDir, "*.yml"))
        {
            try
            {
                var yaml = File.ReadAllText(file);
                var schema = deserializer.Deserialize<ExdSchema>(yaml);
                if (schema?.Name != null)
                {
                    schemas[schema.Name] = schema;
                }
            }
            catch
            {
                // 忽略解析失败的 schema
                Console.WriteLine($"无法解析 schema: {file}");
            }
        }

        return schemas;
    }

    /// <summary>
    /// 导出单个 Sheet 为 CSV
    /// </summary>
    static void ExportSheet(
        GameData lumina,
        string sheetName,
        string outputDir,
        Language language,
        ExdSchema? schema
    )
    {
        // 获取 Excel 头文件
        var headerFile = lumina.GetFile<ExcelHeaderFile>($"exd/{sheetName}.exh");
        if (headerFile == null)
        {
            throw new Exception("无法读取表头文件");
        }

        var header = headerFile.Header;
        var columns = headerFile.ColumnDefinitions;

        // 确定要导出的语言
        Language actualLanguage;
        if (headerFile.Languages.Contains(Language.None))
        {
            actualLanguage = Language.None;
        }
        else if (headerFile.Languages.Contains(language))
        {
            actualLanguage = language;
        }
        else if (headerFile.Languages.Contains(Language.English))
        {
            actualLanguage = Language.English;
        }
        else
        {
            actualLanguage = headerFile.Languages.FirstOrDefault();
        }

        ExportSheetWithLanguage(lumina, sheetName, headerFile, outputDir, actualLanguage, schema);
    }

    static void ExportSheetWithLanguage(
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

        // 创建输出文件路径 (支持子目录)
        var fileName = sheetName.Replace('/', Path.DirectorySeparatorChar);
        var outputPath = Path.Combine(outputDir, fileName + ".csv");
        var outputFileDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputFileDir))
        {
            Directory.CreateDirectory(outputFileDir);
        }

        using var writer = new StreamWriter(outputPath, false, new UTF8Encoding(true)); // UTF-8 with BOM

        // 1. 记录原始索引并按 Offset 排序，以便正确匹配 Schema 列名
        var indexedColumns = columns
            .Select((c, i) => new { Definition = c, OriginalIndex = i })
            .ToList();
        var sortedForSchema = indexedColumns.OrderBy(x => x.Definition.Offset).ToList();

        // 2. 根据排序后的列生成列名
        var (sortedNames, _) = GenerateColumnNamesAndTypes(
            sortedForSchema.Select(x => x.Definition).ToArray(),
            schema
        );

        // 3. 将列名、类型等信息绑定到列对象上
        var columnInfos = new List<ColumnInfo>();
        for (int i = 0; i < sortedForSchema.Count; i++)
        {
            var def = sortedForSchema[i].Definition;
            var originalIndex = sortedForSchema[i].OriginalIndex;
            string name;

            bool isUnknown;
            if (schema?.Fields == null || schema.Fields.Count == 0)
            {
                // 没有 schema，直接使用原始索引作为名字，但标记为 Unknown 以输出空字符串
                name = originalIndex.ToString();
                isUnknown = true;
            }
            else
            {
                // 有 schema，从生成的排序名字列表中获取
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

        // 4. 按照原始索引恢复顺序 (排除 key 之后的所有列)
        var finalColumns = columnInfos.OrderBy(x => x.OriginalIndex).ToList();

        // 检查是否是 Subrow 类型
        bool isSubrow = header.Variant == ExcelVariant.Subrows;

        // 写入第一行: 字段名
        // 写入第一行: key 索引行 (key, 0, 1, 2, ...)
        var keyIndexRow = new List<string> { "key" };
        int totalDataColumns = finalColumns.Count;
        for (int i = 0; i < totalDataColumns; i++)
        {
            keyIndexRow.Add(i.ToString());
        }
        writer.WriteLine(string.Join(",", keyIndexRow.Select(EscapeCsv)));

        // 写入第二行: 字段名行 (原本的 key 改为 #)
        var headerRow = new List<string> { "#" };
        foreach (var col in finalColumns)
        {
            // 如果是 Unknown 字段，表头显示为空白
            headerRow.Add(col.IsUnknown ? "" : col.Name);
        }
        writer.WriteLine(string.Join(",", headerRow.Select(EscapeCsv)));

        // 写入第三行: offset
        var offsetRow = new List<string> { "offset" };
        foreach (var col in finalColumns)
        {
            offsetRow.Add(col.Definition.Offset.ToString());
        }
        writer.WriteLine(string.Join(",", offsetRow.Select(EscapeCsv)));

        // 写入第四行: type
        var typeRow = new List<string> { GetMappedType("Int32") };
        foreach (var col in finalColumns)
        {
            typeRow.Add(GetMappedType(col.Type));
        }
        writer.WriteLine(string.Join(",", typeRow.Select(EscapeCsv)));

        // 遍历所有数据页
        foreach (var pageDef in headerFile.DataPages)
        {
            // 构建数据文件路径
            string exdPath =
                language == Language.None
                    ? $"exd/{sheetName}_{pageDef.StartId}.exd"
                    : $"exd/{sheetName}_{pageDef.StartId}_{LanguageUtil.GetLanguageStr(language)}.exd";

            var exdFile = lumina.GetFile<ExcelDataFile>(exdPath);
            if (exdFile == null)
                continue;

            // 遍历所有行
            foreach (var rowKvp in exdFile.RowData)
            {
                var rowId = rowKvp.Key;
                var rowPtr = rowKvp.Value;

                // 行头结构: 4字节数据大小 + 2字节子行数
                // 数据起始于 rowPtr.Offset + 6
                var rowDataStart = (int)(rowPtr.Offset + 6);

                if (isSubrow)
                {
                    // 读取子行数量
                    var subrowCount = BinaryPrimitives.ReadUInt16BigEndian(
                        exdFile.Data.AsSpan((int)(rowPtr.Offset + 4), 2)
                    );

                    for (ushort subrowId = 0; subrowId < subrowCount; subrowId++)
                    {
                        var rowData = new List<string> { EscapeCsv($"{rowId}.{subrowId}") };

                        // 子行数据起始偏移 (每个子行前有 2 字节的子行ID)
                        var subrowDataStart = rowDataStart + subrowId * (header.DataOffset + 2) + 2;

                        foreach (var col in finalColumns)
                        {
                            var value = ReadColumnValue(
                                exdFile.Data,
                                subrowDataStart,
                                col.Definition,
                                header.DataOffset
                            );

                            if (col.Definition.Type == ExcelColumnDataType.String)
                            {
                                // String 类型强制加双引号并转义内部引号
                                rowData.Add("\"" + value.Replace("\"", "\"\"") + "\"");
                            }
                            else
                            {
                                rowData.Add(EscapeCsv(value));
                            }
                        }

                        writer.WriteLine(string.Join(",", rowData));
                    }
                }
                else
                {
                    // 标准行
                    var rowData = new List<string> { EscapeCsv(rowId.ToString()) };

                    foreach (var col in finalColumns)
                    {
                        var value = ReadColumnValue(
                            exdFile.Data,
                            rowDataStart,
                            col.Definition,
                            header.DataOffset
                        );

                        if (col.Definition.Type == ExcelColumnDataType.String)
                        {
                            // String 类型强制加双引号并转义内部引号
                            rowData.Add("\"" + value.Replace("\"", "\"\"") + "\"");
                        }
                        else
                        {
                            rowData.Add(EscapeCsv(value));
                        }
                    }

                    writer.WriteLine(string.Join(",", rowData));
                }
            }
        }
    }

    /// <summary>
    /// 生成列名和类型
    /// </summary>
    static (List<string> names, List<string> types) GenerateColumnNamesAndTypes(
        ExcelColumnDefinition[] columns,
        ExdSchema? schema
    )
    {
        var names = new List<string>();

        if (schema?.Fields == null || schema.Fields.Count == 0)
        {
            return ([], []);
        }

        // 展开 schema 字段（处理数组）
        var expandedFields = new List<string>();
        foreach (var field in schema.Fields)
        {
            if (field.Type == "array" && field.Count > 0)
            {
                if (field.Fields != null && field.Fields.Count > 0)
                {
                    // 嵌套数组或包含多个字段的数组
                    for (int i = 0; i < field.Count; i++)
                    {
                        for (int j = 0; j < field.Fields.Count; j++)
                        {
                            var subField = field.Fields[j];
                            if (field.Fields.Count == 1)
                            {
                                // 如果数组元素只有一个字段，简写为 Name[i]
                                expandedFields.Add($"{field.Name}[{i}]");
                            }
                            else
                            {
                                expandedFields.Add(
                                    $"{field.Name}[{i}].{subField.Name ?? $"Field{j}"}"
                                );
                            }
                        }
                    }
                }
                else
                {
                    // 简单数组
                    for (int i = 0; i < field.Count; i++)
                    {
                        expandedFields.Add($"{field.Name}[{i}]");
                    }
                }
            }
            else
            {
                expandedFields.Add(field.Name ?? $"Unknown{expandedFields.Count}");
            }
        }

        // 匹配列数
        for (int i = 0; i < columns.Length; i++)
        {
            if (i < expandedFields.Count)
            {
                names.Add(expandedFields[i]);
            }
            else
            {
                names.Add($"Unknown{i}");
            }
        }

        return (names, []);
    }

    /// <summary>
    /// 从原始数据读取列值
    /// </summary>
    static string ReadColumnValue(
        byte[] data,
        int rowDataStart,
        ExcelColumnDefinition column,
        ushort headerDataOffset
    )
    {
        try
        {
            var span = data.AsSpan();
            var offset = rowDataStart + column.Offset;

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
            return $"<error:{ex.Message}>";
        }
    }

    /// <summary>
    /// 读取字符串值
    /// </summary>
    static string ReadString(
        byte[] data,
        int stringFieldOffset,
        int rowDataStart,
        ushort headerDataOffset
    )
    {
        // 字符串字段存储的是相对于字符串区域的偏移量
        var stringOffset = BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(stringFieldOffset, 4));

        // 字符串区域在行数据之后，即 rowDataStart + headerDataOffset
        var stringStart = (int)(rowDataStart + headerDataOffset + stringOffset);

        if (stringStart >= data.Length)
            return "";

        // 找到字符串结尾 (null terminator)
        var stringEnd = stringStart;
        while (stringEnd < data.Length && data[stringEnd] != 0)
        {
            stringEnd++;
        }

        if (stringEnd <= stringStart)
            return "";

        // 使用 SaintCoinach 解码以保留 Hex 标签
        var length = stringEnd - stringStart;
        var buffer = new byte[length];
        Array.Copy(data, stringStart, buffer, 0, length);

        // 使用 Default 解码器
        var xivString = XivStringDecoder.Default.Decode(buffer);
        return xivString.ToString();
    }

    /// <summary>
    /// CSV 值转义
    /// </summary>
    static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        // 如果包含逗号、引号或换行，需要用引号包裹
        if (
            value.Contains(',')
            || value.Contains('"')
            || value.Contains('\n')
            || value.Contains('\r')
        )
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }
}

/// <summary>
/// EXDSchema YAML 结构
/// </summary>
public class ExdSchema
{
    public string? Name { get; set; }
    public string? DisplayField { get; set; }
    public List<ExdField>? Fields { get; set; }
}

public class ExdField
{
    public string? Name { get; set; }
    public string? Type { get; set; }
    public int Count { get; set; }
    public List<ExdField>? Fields { get; set; }
    public List<string>? Targets { get; set; }
    public string? Comment { get; set; }
}

public class ColumnInfo
{
    public required ExcelColumnDefinition Definition { get; set; }
    public int OriginalIndex { get; set; }
    public required string Name { get; set; }
    public required string Type { get; set; }
    public bool IsUnknown { get; set; }
}
