using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using Lumina;
using Lumina.Data;
using XivExdUnpacker.Core;
using XivExdUnpacker.Models;
using XivExdUnpacker.Services;

namespace XivExdUnpacker;

class Program
{
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

    record ClientExportResult
    {
        public string ClientKey { get; init; } = "";
        public string LanguageName { get; init; } = "";
        public string SchemaVersion { get; init; } = "";
        public string OutputDir { get; init; } = "";
        public int SuccessCount { get; init; }
        public int FailedCount { get; init; }
        public double ElapsedSeconds { get; init; }
    }

    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        var parsedArgs = ParseCommandLineArgs(args);

        if (parsedArgs.ShowHelp || args.Length == 0)
        {
            ShowHelp();
            return;
        }

        if (parsedArgs.Languages == null || parsedArgs.Languages.Count == 0)
        {
            Console.WriteLine("错误: 请指定要导出的语言 (使用 --language 或 -l)");
            Console.WriteLine("使用 --help 查看帮助信息");
            return;
        }

        var configService = new ConfigService();
        var config = configService.LoadConfig();

        if (config.GetClients().Count == 0)
        {
            Console.WriteLine("错误: 配置文件中未定义任何客户端。");
            return;
        }

        List<string> selectedKeys;
        List<string> filters = parsedArgs.Sheets ?? new List<string>();

        if (parsedArgs.Languages.Contains("all", StringComparer.OrdinalIgnoreCase))
        {
            selectedKeys = config.GetClients().Keys.ToList();
        }
        else
        {
            selectedKeys = parsedArgs
                .Languages.Where(c => config.GetClients().ContainsKey(c))
                .ToList();

            if (selectedKeys.Count == 0)
            {
                Console.WriteLine(
                    $"错误: 未找到指定的语言: {string.Join(", ", parsedArgs.Languages)}"
                );
                Console.WriteLine($"可用的语言: {string.Join(", ", config.GetClients().Keys)}");
                return;
            }
        }

        Console.CursorVisible = true;

        var totalStopwatch = Stopwatch.StartNew();
        var clientResults = new ConcurrentBag<(string key, ClientExportResult result)>();
        object globalConsoleLock = new object();

        var schemaCache = new ConcurrentDictionary<string, Dictionary<string, ExdSchema>>(
            StringComparer.OrdinalIgnoreCase
        );
        var gameDataPool = new Dictionary<string, GameData>(StringComparer.OrdinalIgnoreCase);

        int maxSheetParallelism =
            config.MaxSheetParallelism ?? Math.Min(Environment.ProcessorCount, 32);
        maxSheetParallelism = Math.Max(1, Math.Min(maxSheetParallelism, 128));

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.ResetColor();

        foreach (var clientKey in selectedKeys)
        {
            var result = RunDumpProcess(
                clientKey,
                config,
                filters,
                globalConsoleLock,
                maxSheetParallelism,
                schemaCache,
                gameDataPool,
                parsedArgs.HexCode,
                parsedArgs.Clear,
                parsedArgs.SkipOffset
            );
            clientResults.Add((clientKey, result));

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        totalStopwatch.Stop();

        var orderedResults = selectedKeys
            .Select(k => clientResults.FirstOrDefault(r => r.key == k).result)
            .Where(r => r != null)
            .ToList();

        if (orderedResults.Count > 0)
        {
            int wKey = Math.Max(10, orderedResults.Max(r => $"[{r.ClientKey}]".Length));
            int wSchema = Math.Max(10, orderedResults.Max(r => r.SchemaVersion.Length));
            int wSuccess = 8;
            int wFailed = 8;
            int wTime = 10;
            int wPath = orderedResults.Max(r => r.OutputDir.Length);

            int totalLineLength = Math.Min(
                wKey + wSchema + wSuccess + wFailed + wTime + 15 + wPath,
                Console.WindowWidth - 1
            );
            string lineSep = new string('=', totalLineLength);

            Console.WriteLine(lineSep);

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(
                $"{"客户端".PadRight(wKey - 3)} │ {"版本".PadRight(wSchema - 2)} │ {"成功".PadLeft(wSuccess - 2)} │ {"失败".PadLeft(wFailed - 2)} │ {"耗时".PadLeft(wTime - 2)} │ 输出目录"
            );
            Console.WriteLine(lineSep);

            foreach (var r in orderedResults)
            {
                Console.ResetColor();
                Console.Write($"{r.ClientKey}".PadRight(wKey) + " │ ");
                Console.Write(r.SchemaVersion.PadRight(wSchema) + " │ ");
                Console.Write(r.SuccessCount.ToString().PadLeft(wSuccess) + " │ ");

                if (r.FailedCount > 0)
                    Console.ForegroundColor = ConsoleColor.Red;
                else
                    Console.ResetColor();
                Console.Write(r.FailedCount.ToString().PadLeft(wFailed) + " │ ");

                Console.ResetColor();
                Console.Write((r.ElapsedSeconds.ToString("F2") + "s").PadLeft(wTime) + " │ ");

                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine(r.OutputDir);
            }

            Console.ResetColor();
            Console.WriteLine(lineSep);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(
                $"完成 | 总计: {orderedResults.Count} 个客户端 | 总耗时: {totalStopwatch.Elapsed.TotalSeconds:F2}s"
            );
            Console.ResetColor();
            Console.WriteLine(lineSep);
        }
    }

    static ClientExportResult RunDumpProcess(
        string clientKey,
        AppConfig config,
        List<string> cmdFilters,
        object globalConsoleLock,
        int maxSheetParallelism,
        ConcurrentDictionary<string, Dictionary<string, ExdSchema>> schemaCache,
        Dictionary<string, GameData> gameDataPool,
        bool useHexcode,
        bool clear,
        bool skipOffset
    )
    {
        var logBuffer = new System.Text.StringBuilder();
        var startTime = DateTime.Now;

        void LogStatus(string status, ConsoleColor color = ConsoleColor.White)
        {
            lock (globalConsoleLock)
            {
                Console.ForegroundColor = color;
                Console.WriteLine($"[{clientKey}] {status}");
                Console.ResetColor();
            }
        }

        void LogDetail(string message)
        {
            logBuffer.AppendLine($"  {message}");
        }

        config.GetClients().TryGetValue(clientKey, out var client);

        var internationalKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "en",
            "ja",
            "de",
            "fr",
        };
        bool isInternational = internationalKeys.Contains(clientKey);

        string? gamePath;
        if (isInternational)
        {
            gamePath = config.GlobalGamePath;
        }
        else
        {
            gamePath = client?.Path;
        }

        bool isDetectedPath = false;
        if (string.IsNullOrEmpty(gamePath) || !Directory.Exists(gamePath))
        {
            LogStatus("⚠ 路径未配置,尝试自动检测...", ConsoleColor.Yellow);

            var detector = new GamePathDetector();
            var detectedPath = detector.Detect(isInternational);

            if (!string.IsNullOrEmpty(detectedPath))
            {
                LogDetail($"✓ 已自动检测到路径: {detectedPath}");
                gamePath = detectedPath;
                isDetectedPath = true;
            }
            else
            {
                LogStatus("✗ 未检测到路径,跳过", ConsoleColor.Red);
                return new ClientExportResult
                {
                    SuccessCount = 0,
                    FailedCount = 0,
                    ElapsedSeconds = 0,
                };
            }
        }

        if (!string.IsNullOrEmpty(gamePath) && Directory.Exists(gamePath))
        {
            var combinedPath = Path.Combine(gamePath, "game", "sqpack");
            if (Directory.Exists(combinedPath))
            {
                gamePath = combinedPath;
            }
        }

        var outputDir = client?.OutputDir;
        if (string.IsNullOrEmpty(outputDir))
        {
            LogStatus("✗ 配置错误: 未指定输出目录", ConsoleColor.Red);
            return new ClientExportResult
            {
                SuccessCount = 0,
                FailedCount = 0,
                ElapsedSeconds = 0,
            };
        }
        string schemaVersion;
        if (isInternational)
            schemaVersion = client?.SchemaVersion ?? config.GlobalSchemaVersion ?? "latest";
        else
            schemaVersion = client?.SchemaVersion ?? "latest";

        var schemas = schemaCache.GetOrAdd(
            schemaVersion,
            version =>
            {
                var dir = Path.Combine("./EXDSchema/schemas", version);
                var service = new SchemaService();
                return service.LoadSchemas(dir);
            }
        );

        if (!KeyToLanguage.TryGetValue(clientKey, out var exportLanguage))
            exportLanguage = Language.English;

        var fullOutputDir = Path.GetFullPath(outputDir);

        if (!Directory.Exists(gamePath))
        {
            LogStatus($"✗ {exportLanguage} 路径不存在,跳过", ConsoleColor.Red);
            return new ClientExportResult
            {
                SuccessCount = 0,
                FailedCount = 0,
                ElapsedSeconds = 0,
            };
        }

        Directory.CreateDirectory(outputDir);

        var failedSheets = new ConcurrentBag<(string name, string error)>();
        int successCount = 0;

        try
        {
            var stopwatch = Stopwatch.StartNew();

            if (!gameDataPool.TryGetValue(gamePath, out var lumina))
            {
                lumina = new GameData(
                    gamePath,
                    new LuminaOptions { DefaultExcelLanguage = exportLanguage }
                );
                gameDataPool[gamePath] = lumina;
            }

            var allSheetNames = lumina.Excel.SheetNames.ToList();
            var sheetNames = new List<string>(allSheetNames);

            if (cmdFilters != null && cmdFilters.Count > 0)
            {
                sheetNames = sheetNames
                    .Where(s =>
                        cmdFilters.Any(f => s.Equals(f, StringComparison.OrdinalIgnoreCase))
                    )
                    .ToList();
            }

            bool shouldClear = clear;
            string clearStrategySource = clear ? "[命令行 --clear]" : "[保留现有文件]";

            LogDetail($"清空策略: {(shouldClear ? "是" : "否")} {clearStrategySource}");
            LogDetail($"待导出表数量: {sheetNames.Count} (总计: {allSheetNames.Count})");

            Directory.CreateDirectory(outputDir);
            if (shouldClear)
            {
                LogDetail($"正在检查并清空输出目录...");
                bool clearSuccess = ClearDirectory(outputDir, globalConsoleLock);
                if (!clearSuccess)
                {
                    return new ClientExportResult
                    {
                        SuccessCount = 0,
                        FailedCount = 0,
                        ElapsedSeconds = 0,
                    };
                }
            }

            bool includeOffset = !skipOffset;
            var exporter = new ExdExporter(useHexcode, includeOffset);

            int schemaCount = 0;
            object consoleLock = new object();

            LogDetail($"表导出并行数: {maxSheetParallelism}");
            LogStatus($"准备解包 | 客户端: {clientKey}", ConsoleColor.Cyan);

            Parallel.ForEach(
                sheetNames,
                new ParallelOptions { MaxDegreeOfParallelism = maxSheetParallelism },
                sheetName =>
                {
                    try
                    {
                        ExdSchema? schema = null;
                        var baseSheetName = sheetName.Contains('/')
                            ? sheetName.Substring(sheetName.LastIndexOf('/') + 1)
                            : sheetName;
                        schemas.TryGetValue(baseSheetName, out schema);

                        exporter.ExportSheet(lumina, sheetName, outputDir, exportLanguage, schema);
                        Interlocked.Increment(ref successCount);
                        if (schema != null)
                            Interlocked.Increment(ref schemaCount);
                    }
                    catch (Exception ex)
                    {
                        failedSheets.Add((sheetName, ex.Message));
                    }
                }
            );

            stopwatch.Stop();
            LogStatus(
                $"✓ 解包完成 | 成功: {successCount} | 失败: {failedSheets.Count} | 耗时: {stopwatch.Elapsed.TotalSeconds:F2}s",
                ConsoleColor.Green
            );

            if (failedSheets.Count > 0)
            {
                LogDetail($"\n失败详情 (前10个):");
                foreach (var (name, error) in failedSheets.Take(10))
                    LogDetail($"  - {name}: {error}");
            }

            if (successCount > 0 && isDetectedPath)
            {
                try
                {
                    LogDetail($"✓ 自动检测的路径有效,正在保存到配置文件...");

                    if (isInternational)
                    {
                        config.GlobalGamePath = gamePath;
                    }
                    else
                    {
                        if (clientKey.Equals("cn", StringComparison.OrdinalIgnoreCase))
                        {
                            config.Cn ??= new ClientConfig();
                            config.Cn.Path = gamePath;
                        }
                        else if (clientKey.Equals("ko", StringComparison.OrdinalIgnoreCase))
                        {
                            config.Ko ??= new ClientConfig();
                            config.Ko.Path = gamePath;
                        }
                        else if (clientKey.Equals("tc", StringComparison.OrdinalIgnoreCase))
                        {
                            config.Tc ??= new ClientConfig();
                            config.Tc.Path = gamePath;
                        }
                    }

                    var configService = new ConfigService();
                    configService.SaveConfig(config);
                    LogDetail($"✓ 配置已保存,下次将直接使用此路径");
                }
                catch (Exception saveEx)
                {
                    LogDetail($"警告: 保存配置失败: {saveEx.Message}");
                }
            }
            stopwatch.Stop();
            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                SchemaVersion = schemaVersion,
                OutputDir = outputDir,
                SuccessCount = successCount,
                FailedCount = failedSheets.Count,
                ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
            };
        }
        catch (Exception ex)
        {
            LogStatus($"✗ 运行失败: {ex.Message}", ConsoleColor.Red);
            return new ClientExportResult();
        }
        finally
        {
            bool hasSignificantInfo = failedSheets.Count > 0 || isDetectedPath;

            if (hasSignificantInfo && logBuffer.Length > 0)
            {
                lock (globalConsoleLock)
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine($"\n[{clientKey}] 补充信息:");
                    Console.WriteLine(logBuffer.ToString());
                    Console.ResetColor();
                }
            }
        }
    }

    static bool ClearDirectory(string path, object consoleLock)
    {
        try
        {
            var dir = new DirectoryInfo(path);
            if (!dir.Exists)
                return true;

            var allFiles = dir.GetFiles("*", SearchOption.AllDirectories).ToList();
            if (allFiles.Count == 0)
            {
                foreach (var subDir in dir.GetDirectories())
                    subDir.Delete(true);
                return true;
            }

            var nonCsvFiles = allFiles
                .Where(f => !f.Extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (nonCsvFiles.Count == 0)
            {
                foreach (var file in allFiles)
                    file.Delete();
                foreach (var subDir in dir.GetDirectories())
                    subDir.Delete(true);
                return true;
            }
            else
            {
                lock (consoleLock)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"\n⚠ 警告: 输出目录包含 {nonCsvFiles.Count} 个非 CSV 文件!");
                    Console.WriteLine($"目录: {path}");
                    Console.WriteLine("\n非 CSV 文件列表 (前20个):");
                    Console.ResetColor();

                    foreach (var file in nonCsvFiles.Take(20))
                    {
                        var relativePath = Path.GetRelativePath(path, file.FullName);
                        Console.WriteLine($"  - {relativePath}");
                    }

                    if (nonCsvFiles.Count > 20)
                        Console.WriteLine($"  ... 还有 {nonCsvFiles.Count - 20} 个文件");

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\n是否仍要清空此目录? [y/N]");
                    Console.ResetColor();
                    Console.Write("> ");

                    var response = Console.ReadLine()?.Trim().ToLower();
                    if (response == "y" || response == "yes")
                    {
                        Console.WriteLine("正在清空目录...");
                        foreach (var file in allFiles)
                            file.Delete();
                        foreach (var subDir in dir.GetDirectories())
                            subDir.Delete(true);
                        return true;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("✗ 已取消清空,将跳过此客户端");
                        Console.ResetColor();
                        return false;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            lock (consoleLock)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"清空目录失败: {ex.Message}");
                Console.ResetColor();
            }
            return false;
        }
    }

    record CommandLineArgs
    {
        public List<string>? Languages { get; init; }
        public List<string>? Sheets { get; init; }
        public bool HexCode { get; init; }
        public bool Clear { get; init; }
        public bool SkipOffset { get; init; }
        public bool ShowHelp { get; init; }
    }

    static CommandLineArgs ParseCommandLineArgs(string[] args)
    {
        var languages = new List<string>();
        var sheets = new List<string>();
        bool hexCode = false;
        bool clear = false;
        bool skipOffset = false;
        bool showHelp = false;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            switch (arg.ToLower())
            {
                case "--help":
                case "-h":
                case "/?":
                    showHelp = true;
                    break;

                case "--language":
                case "-l":
                    i++;
                    while (i < args.Length && !args[i].StartsWith("-"))
                    {
                        languages.Add(args[i]);
                        i++;
                    }
                    i--;
                    break;

                case "--sheets":
                case "-s":
                    i++;
                    while (i < args.Length && !args[i].StartsWith("-"))
                    {
                        sheets.Add(args[i]);
                        i++;
                    }
                    i--;
                    break;

                case "--hexcode":
                case "-x":
                    hexCode = true;
                    break;

                case "--clear":
                case "-c":
                    clear = true;
                    break;

                case "--skip-offset":
                    skipOffset = true;
                    break;
            }
        }

        return new CommandLineArgs
        {
            Languages = languages.Count > 0 ? languages : null,
            Sheets = sheets.Count > 0 ? sheets : null,
            HexCode = hexCode,
            Clear = clear,
            SkipOffset = skipOffset,
            ShowHelp = showHelp,
        };
    }

    static void ShowHelp()
    {
        Console.WriteLine("FFXIV EXD 数据解包工具");
        Console.WriteLine();
        Console.WriteLine("用法:");
        Console.WriteLine("  XivExdUnpacker --language <语言...> [选项]");
        Console.WriteLine();
        Console.WriteLine("必需参数:");
        Console.WriteLine(
            "  --language, -l <语言...>     指定要导出的语言 (en ja de fr cn ko tc all)"
        );
        Console.WriteLine();
        Console.WriteLine("可选参数:");
        Console.WriteLine("  --sheets, -s <表名...>       指定要导出的表名 (默认: 全部)");
        Console.WriteLine(
            "  --hexcode, -x                保留原始数据 (默认: 解码字符串为人类可读格式)"
        );
        Console.WriteLine("  --clear, -c                  导出前清空输出目录");
        Console.WriteLine("  --skip-offset                跳过 CSV 的 offset 行");
        Console.WriteLine("  --help, -h                   显示此帮助信息");
        Console.WriteLine();
        Console.WriteLine("示例:");
        Console.WriteLine("  # 导出中文的所有表 (默认解码字符串)");
        Console.WriteLine("  XivExdUnpacker --language cn");
        Console.WriteLine();
        Console.WriteLine("  # 导出英文的所有表 (保留原始数据)");
        Console.WriteLine("  XivExdUnpacker --language en --hexcode");
        Console.WriteLine();
        Console.WriteLine("  # 导出英文和日文的 Action 和 Item 表");
        Console.WriteLine("  XivExdUnpacker --language en ja --sheets Action Item");
        Console.WriteLine();
        Console.WriteLine("  # 导出所有语言，清空输出目录，跳过 offset 行");
        Console.WriteLine("  XivExdUnpacker --language all --clear --skip-offset");
        Console.WriteLine();
        Console.WriteLine("  # 使用简写");
        Console.WriteLine("  XivExdUnpacker -l cn -s Addon Quest -x -c");
    }
}
