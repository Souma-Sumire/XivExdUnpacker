using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using Lumina;
using Lumina.Data;
using XivExdUnpacker.Core;
using XivExdUnpacker.Models;
using XivExdUnpacker.Services;
using XivExdUnpacker.UI;

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
        public int SuccessCount { get; init; }
        public int FailedCount { get; init; }
        public double ElapsedSeconds { get; init; }
    }

    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.CursorVisible = false;
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.ResetColor();

        var configService = new ConfigService();
        var config = configService.LoadConfig();

        if (config.GetClients().Count == 0)
        {
            Console.WriteLine("错误: 配置文件中未定义任何客户端。");
            return;
        }

        List<string>? selectedKeys;
        List<string> filters = new List<string>();

        // 交互模式
        selectedKeys = Menu.ShowInteractiveMenu(config.GetClients().Keys.ToList());

        if (selectedKeys == null || selectedKeys.Count == 0)
        {
            Console.WriteLine("任务已取消或未选择任何客户端。");
            Console.CursorVisible = true;
            return;
        }

        Console.WriteLine("提示: 成功选择 " + selectedKeys.Count + " 个客户端。");

        // 询问过滤器
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n请输入要导出的表名 (多个用空格分隔，直接回车则导出全部):");
        Console.ResetColor();
        Console.Write("> ");
        var input = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(input))
        {
            filters = input.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        Console.CursorVisible = true;

        // 并行处理多个客户端
        var totalStopwatch = Stopwatch.StartNew();
        var clientResults = new ConcurrentBag<(string key, ClientExportResult result)>();
        object globalConsoleLock = new object();

        // 创建 Schema 缓存池 (Key: 版本号, Value: 该版本的 Schema 字典)
        var schemaCache = new ConcurrentDictionary<string, Dictionary<string, ExdSchema>>(
            StringComparer.OrdinalIgnoreCase
        );

        // 获取客户端并行度配置 (建议: 默认不要超过核心数的 1/4)
        // 获取表导出并行度配置 (默认: CPU核心数)
        // 获取表导出并行度配置 (默认: CPU核心数, 但上限设为 32 以保护 I/O)
        int maxSheetParallelism =
            config.MaxSheetParallelism ?? Math.Min(Environment.ProcessorCount, 32);
        maxSheetParallelism = Math.Max(1, Math.Min(maxSheetParallelism, 128)); // 强制物理上限 128

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine($"\n[性能配置] 表导出并行数: {maxSheetParallelism}");
        Console.ResetColor();

        foreach (var clientKey in selectedKeys)
        {
            var result = RunDumpProcess(
                clientKey,
                config,
                filters,
                globalConsoleLock,
                maxSheetParallelism,
                schemaCache
            );
            clientResults.Add((clientKey, result));
        }

        totalStopwatch.Stop();

        // 按原始顺序输出汇总结果
        Console.WriteLine("\n" + new string('=', 60));
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine(
            $"所有客户端处理完毕 (总耗时: {totalStopwatch.Elapsed.TotalSeconds:F2}s)"
        );
        Console.ResetColor();
        Console.WriteLine(new string('=', 60));

        Console.WriteLine("\n" + new string('=', 60));
    }

    static ClientExportResult RunDumpProcess(
        string clientKey,
        AppConfig config,
        List<string> cmdFilters,
        object globalConsoleLock,
        int maxSheetParallelism,
        ConcurrentDictionary<string, Dictionary<string, ExdSchema>> schemaCache
    )
    {
        var logBuffer = new System.Text.StringBuilder();
        var startTime = DateTime.Now;

        // 简洁的状态日志,只在关键时刻输出
        void LogStatus(string status, ConsoleColor color = ConsoleColor.White)
        {
            lock (globalConsoleLock)
            {
                Console.ForegroundColor = color;
                Console.WriteLine($"[{clientKey}] {status}");
                Console.ResetColor();
            }
        }

        // 详细日志缓冲到 StringBuilder,最后一次性输出
        void LogDetail(string message)
        {
            logBuffer.AppendLine($"  {message}");
        }

        config.GetClients().TryGetValue(clientKey, out var client);

        // 定义国际服客户端 (共用 Global Path)
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
            // 国际服: 强制使用 Global 配置
            gamePath = config.GlobalGamePath;
        }
        else
        {
            // 区域服: 强制使用 Client 配置
            gamePath = client?.Path;
        }

        bool isDetectedPath = false;
        if (string.IsNullOrEmpty(gamePath) || !Directory.Exists(gamePath))
        {
            LogStatus("⚠ 路径未配置,尝试自动检测...", ConsoleColor.Yellow);

            var detector = new GamePathDetector();
            var detectedPath = detector.Detect(isInternational); // 传入客户端类型

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

        // 路径补全逻辑：如果给的是根目录，自动补全到 sqpack
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
        // 只有在此时才确定该客户端需要的 Schema 版本
        string schemaVersion;
        if (isInternational)
            schemaVersion = client?.SchemaVersion ?? config.GlobalSchemaVersion ?? "latest";
        else
            schemaVersion = client?.SchemaVersion ?? "latest";

        // 从缓存中获取，如果没有则加载并存入缓存
        var schemas = schemaCache.GetOrAdd(
            schemaVersion,
            version =>
            {
                LogStatus($"正在加载 Schema 版本: {version}...", ConsoleColor.DarkGray);
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

        // 在 try 外部声明,以便 finally 块访问
        var failedSheets = new ConcurrentBag<(string name, string error)>();
        int successCount = 0;

        try
        {
            var stopwatch = Stopwatch.StartNew();
            LogDetail("正在初始化 Lumina...");
            var lumina = new GameData(
                gamePath,
                new LuminaOptions { DefaultExcelLanguage = exportLanguage }
            );

            // 1. 获取所有表名并过滤
            var allSheetNames = lumina.Excel.SheetNames.ToList();
            var sheetNames = new List<string>(allSheetNames);

            // 过滤逻辑: 应用交互式/命令行 Include (白名单)
            if (cmdFilters != null && cmdFilters.Count > 0)
            {
                sheetNames = sheetNames
                    .Where(s =>
                        cmdFilters.Any(f => s.Equals(f, StringComparison.OrdinalIgnoreCase))
                    )
                    .ToList();
            }

            // 清空策略: 三态逻辑 (True=强制清空, False=强制保留, Null=智能判断)
            bool shouldClear;
            string clearStrategySource;

            if (config.ClearOutputDir.HasValue)
            {
                shouldClear = config.ClearOutputDir.Value;
                clearStrategySource = shouldClear ? "[配置强制清空]" : "[配置强制保留]";
            }
            else
            {
                // 智能模式: 如果没有过滤器(全量)则清空，有过滤器(部分)则保留
                bool isPartialExport = cmdFilters != null && cmdFilters.Count > 0;
                shouldClear = !isPartialExport;
                clearStrategySource = isPartialExport
                    ? "[智能保留 (部分导出)]"
                    : "[智能清空 (全量导出)]";
            }

            LogDetail($"清空策略: {(shouldClear ? "是" : "否")} {clearStrategySource}");
            LogDetail($"待导出表数量: {sheetNames.Count} (总计: {allSheetNames.Count})");

            Directory.CreateDirectory(outputDir);
            if (shouldClear)
            {
                LogDetail($"正在检查并清空输出目录...");
                bool clearSuccess = ClearDirectory(outputDir, globalConsoleLock);
                if (!clearSuccess)
                {
                    // 用户取消清空,跳过此客户端
                    return new ClientExportResult
                    {
                        SuccessCount = 0,
                        FailedCount = 0,
                        ElapsedSeconds = 0,
                    };
                }
            }

            var exporter = new ExdExporter();

            int schemaCount = 0;
            object consoleLock = new object();

            LogDetail($"表导出并行数: {maxSheetParallelism}");

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
                $"✓ {exportLanguage} | {schemaVersion} | {stopwatch.Elapsed.TotalSeconds:F2}s | 成功: {successCount} | 失败: {failedSheets.Count}",
                ConsoleColor.Green
            );

            if (failedSheets.Count > 0)
            {
                LogDetail($"\n失败详情 (前10个):");
                foreach (var (name, error) in failedSheets.Take(10))
                    LogDetail($"  - {name}: {error}");
            }

            // 如果使用了自动检测的路径且导出成功,自动保存到配置
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
                        // 更新对应的区域服配置
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

            return new ClientExportResult
            {
                SuccessCount = successCount,
                FailedCount = failedSheets.Count,
                ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
            };
        }
        catch (Exception ex)
        {
            LogStatus($"✗ 失败: {ex.Message}", ConsoleColor.Red);
            return new ClientExportResult
            {
                SuccessCount = 0,
                FailedCount = 0,
                ElapsedSeconds = 0,
            };
        }
        finally
        {
            // 只有在【有失败】或【路径是探测出来的】时候才显示这个详细信息块
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
                // 空目录,直接清空
                foreach (var subDir in dir.GetDirectories())
                    subDir.Delete(true);
                return true;
            }

            // 检查是否只有 .csv 文件
            var nonCsvFiles = allFiles
                .Where(f => !f.Extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (nonCsvFiles.Count == 0)
            {
                // 只有 CSV 文件,安全删除
                foreach (var file in allFiles)
                    file.Delete();
                foreach (var subDir in dir.GetDirectories())
                    subDir.Delete(true);
                return true;
            }
            else
            {
                // 有非 CSV 文件,需要用户确认
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
}
