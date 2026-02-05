using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using Lumina;
using Lumina.Data;
using XivExdUnpacker.Core;
using XivExdUnpacker.Models;
using XivExdUnpacker.Services;

namespace XivExdUnpacker.src;

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
        public string ClientVersion { get; init; } = "";
        public string ActualSchema { get; init; } = "";
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
        List<string> filters = parsedArgs.Sheets ?? [];

        if (parsedArgs.Languages.Contains("all", StringComparer.OrdinalIgnoreCase))
        {
            selectedKeys = [.. config.GetClients().Keys];
        }
        else
        {
            selectedKeys = [.. parsedArgs.Languages.Where(c => config.GetClients().ContainsKey(c))];

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
        object globalConsoleLock = new();

        var schemaCache = new ConcurrentDictionary<string, Dictionary<string, ExdSchema>>(
            StringComparer.OrdinalIgnoreCase
        );
        var gameDataPool = new Dictionary<string, GameData>(StringComparer.OrdinalIgnoreCase);

        int maxSheetParallelism =
            config.MaxSheetParallelism ?? Math.Clamp(Environment.ProcessorCount, 1, 8);
        maxSheetParallelism = Math.Clamp(maxSheetParallelism, 1, 128);

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
            int wVer = Math.Max(10, orderedResults.Max(r => r.ClientVersion.Length));
            int wSchema = Math.Max(10, orderedResults.Max(r => r.ActualSchema.Length));
            int wSuccess = 8;
            int wFailed = 8;
            int wTime = 10;
            int wPath = orderedResults.Max(r => r.OutputDir.Length);

            int totalLineLength = Math.Min(
                wKey + wVer + wSchema + wSuccess + wFailed + wTime + 18 + wPath,
                Console.WindowWidth - 1
            );
            string lineSep = new('=', totalLineLength);

            Console.WriteLine(lineSep);

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(
                $"{"客户端".PadRight(wKey - 3)} │ {"版本".PadRight(wVer - 2)} │ {"Schema".PadRight(wSchema - 2)} │ {"成功".PadLeft(wSuccess - 2)} │ {"失败".PadLeft(wFailed - 2)} │ {"耗时".PadLeft(wTime - 2)} │ 输出目录"
            );
            Console.WriteLine(lineSep);

            foreach (var r in orderedResults)
            {
                Console.ResetColor();
                Console.Write($"{r.ClientKey}".PadRight(wKey) + " │ ");
                Console.Write(r.ClientVersion.PadRight(wVer) + " │ ");
                Console.Write(r.ActualSchema.PadRight(wSchema) + " │ ");
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
        var logBuffer = new StringBuilder();
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

        void LogDetail(string message) => logBuffer.AppendLine($"  {message}");

        // 1. 获取客户端基本信息
        config.GetClients().TryGetValue(clientKey, out var client);
        if (!KeyToLanguage.TryGetValue(clientKey, out var exportLanguage))
            exportLanguage = Language.English;

        var internationalKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "en",
            "ja",
            "de",
            "fr",
        };
        bool isInternational = internationalKeys.Contains(clientKey);
        string? gamePath = isInternational ? config.GlobalGamePath : client?.Path;
        string? outputDir = client?.OutputDir;

        // 2. 路径检测与校验
        bool isDetectedPath = false;
        if (string.IsNullOrEmpty(gamePath) || !Directory.Exists(gamePath))
        {
            LogStatus("⚠ 路径未配置,尝试自动检测...", ConsoleColor.Yellow);
            gamePath = new GamePathDetector().Detect(isInternational);
            if (string.IsNullOrEmpty(gamePath))
            {
                LogStatus("✗ 未检测到路径,跳过", ConsoleColor.Red);
                return new ClientExportResult();
            }
            LogDetail($"✓ 已自动检测到路径: {gamePath}");
            isDetectedPath = true;
        }

        // 路径规范化 (sqpack 检测)
        var combinedPath = Path.Combine(gamePath, "game", "sqpack");
        if (Directory.Exists(combinedPath))
            gamePath = combinedPath;

        if (string.IsNullOrEmpty(outputDir))
        {
            LogStatus("✗ 配置错误: 未指定输出目录", ConsoleColor.Red);
            return new ClientExportResult();
        }

        if (!Directory.Exists(gamePath))
        {
            LogStatus($"✗ {exportLanguage} 路径不存在,跳过", ConsoleColor.Red);
            return new ClientExportResult();
        }

        var failedSheets = new ConcurrentBag<(string name, string error)>();
        int successCount = 0;

        try
        {
            var stopwatch = Stopwatch.StartNew();

            // 3. 初始化 Lumina
            if (!gameDataPool.TryGetValue(gamePath, out var lumina))
            {
                lumina = new GameData(
                    gamePath,
                    new LuminaOptions { DefaultExcelLanguage = exportLanguage }
                );
                gameDataPool[gamePath] = lumina;
            }

            // 4. 版本与 Schema 检测
            string? detectedVersion = null;
            try
            {
                var parentDir = Directory.GetParent(gamePath);
                var verFile =
                    parentDir != null ? Path.Combine(parentDir.FullName, "ffxivgame.ver") : null;
                if (verFile != null && File.Exists(verFile))
                    detectedVersion = File.ReadAllText(verFile).Trim();
            }
            catch { }

            if (!string.IsNullOrEmpty(detectedVersion))
                LogDetail($"✓ 自动检测到游戏版本: {detectedVersion}");

            string? schemaVersion = isInternational
                ? client?.SchemaVersion ?? config.GlobalSchemaVersion
                : client?.SchemaVersion;
            if (
                string.IsNullOrEmpty(schemaVersion)
                || schemaVersion.Equals("latest", StringComparison.OrdinalIgnoreCase)
            )
                schemaVersion = detectedVersion ?? "latest";

            string finalSchemaDir = "";
            var schemas = schemaCache.GetOrAdd(
                schemaVersion!,
                version =>
                {
                    var schemaRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "schemas");
                    var targetDir = Path.Combine(schemaRoot, version);

                    // Strategy 1: Local Cache (Fixed Version)
                    if (
                        !version.Equals("latest", StringComparison.OrdinalIgnoreCase)
                        && Directory.Exists(targetDir)
                    )
                    {
                        finalSchemaDir = targetDir;
                        return new SchemaService().LoadSchemas(targetDir);
                    }

                    // Strategy 2: Online Download
                    LogStatus($"⚠ 准备检查更新: {version}...", ConsoleColor.Yellow);
                    var onlinePath = SchemaUpdater
                        .DownloadAndExtractSchema(version, msg => LogDetail(msg), schemaRoot)
                        .Result;
                    if (!string.IsNullOrEmpty(onlinePath) && Directory.Exists(onlinePath))
                    {
                        LogStatus($"✓ Schema 获取成功: {version}", ConsoleColor.Green);
                        finalSchemaDir = onlinePath;
                        return new SchemaService().LoadSchemas(onlinePath);
                    }

                    // Strategy 3: Local Fallback
                    var latestDir = Path.Combine(schemaRoot, "latest");
                    if (Directory.Exists(latestDir))
                    {
                        LogStatus($"⚠ 在线更新失败, 回退至本地 'latest'", ConsoleColor.Yellow);
                        finalSchemaDir = latestDir;
                        return new SchemaService().LoadSchemas(latestDir);
                    }

                    finalSchemaDir = targetDir;
                    return new SchemaService().LoadSchemas(targetDir);
                }
            );

            // 5. 准备输出目录
            Directory.CreateDirectory(outputDir);
            var sheetNames = lumina.Excel.SheetNames.ToList();
            if (cmdFilters?.Count > 0)
                sheetNames =
                [
                    .. sheetNames.Where(s =>
                        cmdFilters.Any(f => s.Equals(f, StringComparison.OrdinalIgnoreCase))
                    ),
                ];

            LogDetail(
                $"清空策略: {(clear ? "是" : "否")} {(clear ? "[命令行 --clear]" : "[保留现有文件]")}"
            );
            LogDetail($"待导出表数量: {sheetNames.Count}");

            if (clear && !ClearDirectory(outputDir, globalConsoleLock))
                return new ClientExportResult();

            // 6. 核心解压逻辑
            var exporter = new ExdExporter(useHexcode, !skipOffset);
            LogDetail($"表导出并行数: {maxSheetParallelism}");
            LogStatus($"准备解包 | 客户端: {clientKey}", ConsoleColor.Cyan);

            Parallel.ForEach(
                sheetNames,
                new ParallelOptions { MaxDegreeOfParallelism = maxSheetParallelism },
                sheetName =>
                {
                    try
                    {
                        var baseSheetName = sheetName.Contains('/')
                            ? sheetName[(sheetName.LastIndexOf('/') + 1)..]
                            : sheetName;
                        schemas.TryGetValue(baseSheetName, out var schema);
                        exporter.ExportSheet(lumina, sheetName, outputDir, exportLanguage, schema);
                        Interlocked.Increment(ref successCount);
                    }
                    catch (Exception ex)
                    {
                        failedSheets.Add((sheetName, ex.Message));
                    }
                }
            );

            // 7. 善后处理
            if (successCount > 0 && isDetectedPath)
                SaveDetectedPath(gamePath);

            stopwatch.Stop();
            LogStatus(
                $"✓ 解包完成 | 成功: {successCount} | 失败: {failedSheets.Count} | 耗时: {stopwatch.Elapsed.TotalSeconds:F2}s",
                ConsoleColor.Green
            );

            if (!failedSheets.IsEmpty)
            {
                LogDetail($"\n失败详情 (前10个):");
                foreach (var (name, error) in failedSheets.Take(10))
                    LogDetail($"  - {name}: {error}");
            }

            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                ClientVersion = detectedVersion ?? "Unknown",
                ActualSchema = string.IsNullOrEmpty(finalSchemaDir)
                    ? "Cache"
                    : Path.GetFileName(
                        finalSchemaDir.TrimEnd(
                            Path.DirectorySeparatorChar,
                            Path.AltDirectorySeparatorChar
                        )
                    ),
                OutputDir = outputDir,
                SuccessCount = successCount,
                FailedCount = failedSheets.Count,
                ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
            };
        }
        catch (Exception ex)
        {
            LogStatus($"✗ 运行失败: {ex.Message}", ConsoleColor.Red);
            return new ClientExportResult { ClientKey = clientKey, ActualSchema = "Error" };
        }
        finally
        {
            if ((!failedSheets.IsEmpty || isDetectedPath) && logBuffer.Length > 0)
            {
                lock (globalConsoleLock)
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine($"\n[{clientKey}] 补充信息:\n{logBuffer}");
                    Console.ResetColor();
                }
            }
        }

        void SaveDetectedPath(string path)
        {
            try
            {
                LogDetail($"✓ 自动检测的路径有效,正在保存到配置文件...");
                if (isInternational)
                    config.GlobalGamePath = path;
                else
                {
                    var target = clientKey.ToLower() switch
                    {
                        "cn" => config.Cn ??= new(),
                        "ko" => config.Ko ??= new(),
                        "tc" => config.Tc ??= new(),
                        _ => null,
                    };
                    if (target != null)
                        target.Path = path;
                }
                new ConfigService().SaveConfig(config);
            }
            catch (Exception e)
            {
                LogDetail($"警告: 保存配置失败: {e.Message}");
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
                        var lockedFiles = new List<string>();
                        foreach (var file in allFiles)
                        {
                            try
                            {
                                file.Delete();
                            }
                            catch (IOException)
                            {
                                lockedFiles.Add(Path.GetRelativePath(path, file.FullName));
                            }
                            catch { }
                        }

                        foreach (var subDir in dir.GetDirectories())
                        {
                            try
                            {
                                subDir.Delete(true);
                            }
                            catch { }
                        }

                        if (lockedFiles.Count > 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(
                                $"\n✗ 警告: 有 {lockedFiles.Count} 个文件因被占用无法删除:"
                            );
                            foreach (var f in lockedFiles.Take(10))
                                Console.WriteLine($"  - {f}");
                            if (lockedFiles.Count > 10)
                                Console.WriteLine("  ...");
                            Console.ResetColor();
                            return false;
                        }
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
                    while (i < args.Length && !args[i].StartsWith('-'))
                    {
                        languages.Add(args[i]);
                        i++;
                    }
                    i--;
                    break;

                case "--sheets":
                case "-s":
                    i++;
                    while (i < args.Length && !args[i].StartsWith('-'))
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
