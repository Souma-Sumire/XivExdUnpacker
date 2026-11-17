using System.Buffers;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Enumeration;
using Lumina;
using Lumina.Data;
using XivExdUnpacker.Services;
using XivExdUnpacker.src.Cli;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src.Core;

public static class ExportRunner
{
    private static readonly Dictionary<string, Language> KeyToLanguage = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        { "en", Language.English },
        { "ja", Language.Japanese },
        { "de", Language.German },
        { "fr", Language.French },
        { "chs", Language.ChineseSimplified },
        { "ko", Language.Korean },
        { "tc", Language.ChineseTraditional },
    };

    private static readonly HashSet<string> InternationalKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "en",
        "ja",
        "de",
        "fr",
    };

    private sealed class ClientPlan
    {
        public required string ClientKey { get; init; }
        public required Language ExportLanguage { get; init; }
        public required bool IsInternational { get; init; }
        public string? GamePath { get; init; }
        public bool IsDetectedPath { get; init; }
        public string OutputDir { get; init; } = "";
        public string? DetectedVersion { get; init; }
        public string SchemaVersion { get; init; } = "latest";
        public ClientExportResult? EarlyFailure { get; init; }
    }

    private sealed class SharedPayloadEntry(byte[] data, int remainingConsumers)
    {
        public readonly byte[] Data = data;
        public int RemainingConsumers = remainingConsumers;
    }

    private sealed class SharedGroupCache
    {
        public readonly ConcurrentDictionary<string, SheetLayout> Layouts = new(
            StringComparer.OrdinalIgnoreCase
        );

        public readonly ConcurrentDictionary<string, SharedPayloadEntry> LanguageNonePayloads = new(
            StringComparer.OrdinalIgnoreCase
        );
    }

    private readonly record struct PendingIoItem(
        string SheetName,
        string TargetPath,
        byte[] Data,
        int Length,
        long? ExistingSize,
        bool ReturnToPool
    );

    public static List<ClientExportResult> RunAll(
        List<string> selectedKeys,
        AppConfig config,
        List<string> cmdFilters,
        int maxSheetParallelism,
        bool useHexcode,
        bool clear,
        bool skipOffset
    )
    {
        var plans = new List<ClientPlan>(selectedKeys.Count);
        foreach (var key in selectedKeys)
            plans.Add(PrepareClientPlan(key, config));

        var validPlans = plans.Where(p => p.EarlyFailure == null).ToList();
        var gameDataPool = new ConcurrentDictionary<string, GameData>(
            StringComparer.OrdinalIgnoreCase
        );
        var schemaCache = new ConcurrentDictionary<
            string,
            (string dir, Dictionary<string, ExdSchema> schemas)
        >(StringComparer.OrdinalIgnoreCase);

        if (validPlans.Count > 0)
        {
            var schemaRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "schemas");
            Directory.CreateDirectory(schemaRoot);

            var uniqueGamePaths = validPlans
                .GroupBy(p => p.GamePath!, StringComparer.OrdinalIgnoreCase)
                .Select(g => (Path: g.Key, DefaultLang: g.First().ExportLanguage))
                .ToList();

            var uniqueSchemaVersions = validPlans
                .Select(p => p.SchemaVersion)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // 并发初始化所有 GameData 实例并同步拉取/解析 Schema，消除串行启动阻塞
            var initGameDataTask = Task.Run(() =>
            {
                Parallel.ForEach(
                    uniqueGamePaths,
                    item =>
                    {
                        var gd = new GameData(
                            item.Path,
                            new LuminaOptions
                            {
                                DefaultExcelLanguage = item.DefaultLang,
                                CacheFileResources = false,
                            }
                        );
                        gameDataPool.TryAdd(item.Path, gd);
                    }
                );
            });

            var initSchemasTask = Task.Run(() =>
            {
                Parallel.ForEach(
                    uniqueSchemaVersions,
                    ver =>
                    {
                        if (schemaCache.ContainsKey(ver))
                            return;

                        var loaded = LoadOrFetchSchemas(ver, schemaRoot, out var finalDir);
                        if (loaded.Count > 0)
                            schemaCache.TryAdd(ver, (finalDir, loaded));
                    }
                );
            });

            Task.WaitAll(initGameDataTask, initSchemasTask);
        }

        // 统计共享同一 (GameData 路径 + Schema 目录) 的客户端数量，以精准控制共享表字节的生命周期
        var groupRemainingCount = new Dictionary<(string gamePath, string schemaDir), int>();
        var groupCaches = new Dictionary<(string gamePath, string schemaDir), SharedGroupCache>();

        foreach (var plan in validPlans)
        {
            if (schemaCache.TryGetValue(plan.SchemaVersion, out var sc) && !string.IsNullOrEmpty(sc.dir))
            {
                var groupKey = (plan.GamePath!.ToLowerInvariant(), sc.dir.ToLowerInvariant());
                groupRemainingCount[groupKey] = groupRemainingCount.GetValueOrDefault(groupKey) + 1;
                if (!groupCaches.ContainsKey(groupKey))
                    groupCaches[groupKey] = new SharedGroupCache();
            }
        }

        var results = new List<ClientExportResult>(plans.Count);
        bool savedGlobalDetectedPath = false;

        foreach (var plan in plans)
        {
            if (plan.EarlyFailure != null)
            {
                results.Add(plan.EarlyFailure);
                continue;
            }

            if (
                !schemaCache.TryGetValue(plan.SchemaVersion, out var sc)
                || sc.schemas.Count == 0
                || string.IsNullOrEmpty(sc.dir)
            )
            {
                ConsoleLogger.TagError(plan.ClientKey, Localizer.T(MsgKey.SchemaNoneLoaded));
                results.Add(
                    new ClientExportResult
                    {
                        ClientKey = plan.ClientKey,
                        LanguageName = plan.ExportLanguage.ToString(),
                        ClientVersion = plan.DetectedVersion ?? "-",
                        ActualSchema = "-",
                        OutputDir = plan.OutputDir,
                        Status = ExportStatus.Failed,
                        StatusReason = Localizer.T(MsgKey.SchemaNoneLoaded),
                    }
                );
                continue;
            }

            var groupKey = (plan.GamePath!.ToLowerInvariant(), sc.dir.ToLowerInvariant());
            int remainingAfterThis = --groupRemainingCount[groupKey];
            var groupCache = groupCaches[groupKey];

            var result = ExecuteClientExport(
                plan,
                config,
                gameDataPool[plan.GamePath!],
                sc.schemas,
                GetDirName(sc.dir, plan.SchemaVersion),
                groupCache,
                remainingAfterThis,
                cmdFilters,
                maxSheetParallelism,
                useHexcode,
                clear,
                skipOffset,
                ref savedGlobalDetectedPath
            );
            results.Add(result);
        }

        return results;
    }

    public static ClientExportResult RunDumpProcess(
        string clientKey,
        AppConfig config,
        List<string> cmdFilters,
        int maxSheetParallelism,
        ConcurrentDictionary<
            string,
            (string dir, Dictionary<string, ExdSchema> schemas)
        > schemaCache,
        Dictionary<string, GameData> gameDataPool,
        bool useHexcode,
        bool clear,
        bool skipOffset
    )
    {
        var plan = PrepareClientPlan(clientKey, config);
        if (plan.EarlyFailure != null)
            return plan.EarlyFailure;

        try
        {
            if (!gameDataPool.TryGetValue(plan.GamePath!, out var lumina))
            {
                lumina = new GameData(
                    plan.GamePath!,
                    new LuminaOptions
                    {
                        DefaultExcelLanguage = plan.ExportLanguage,
                        CacheFileResources = false,
                    }
                );
                gameDataPool[plan.GamePath!] = lumina;
            }

            var schemaRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "schemas");
            Directory.CreateDirectory(schemaRoot);

            string finalSchemaDir;
            Dictionary<string, ExdSchema> schemas;
            if (schemaCache.TryGetValue(plan.SchemaVersion, out var cached))
            {
                finalSchemaDir = cached.dir;
                schemas = cached.schemas;
            }
            else
            {
                schemas = LoadOrFetchSchemas(plan.SchemaVersion, schemaRoot, out finalSchemaDir);
                if (schemas.Count > 0)
                    schemaCache.TryAdd(plan.SchemaVersion, (finalSchemaDir, schemas));
            }

            if (schemas.Count == 0 || string.IsNullOrEmpty(finalSchemaDir))
            {
                ConsoleLogger.TagError(clientKey, Localizer.T(MsgKey.SchemaNoneLoaded));
                return new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = plan.ExportLanguage.ToString(),
                    ClientVersion = plan.DetectedVersion ?? "-",
                    ActualSchema = "-",
                    OutputDir = plan.OutputDir,
                    Status = ExportStatus.Failed,
                    StatusReason = Localizer.T(MsgKey.SchemaNoneLoaded),
                };
            }

            bool savedFlag = false;
            return ExecuteClientExport(
                plan,
                config,
                lumina,
                schemas,
                GetDirName(finalSchemaDir, plan.SchemaVersion),
                new SharedGroupCache(),
                0,
                cmdFilters,
                maxSheetParallelism,
                useHexcode,
                clear,
                skipOffset,
                ref savedFlag
            );
        }
        catch (Exception ex)
        {
            ConsoleLogger.ClearProgress();
            ConsoleLogger.TagError(clientKey, Localizer.TF(MsgKey.RunFailed, ex.Message));
            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = plan.ExportLanguage.ToString(),
                ClientVersion = plan.DetectedVersion ?? "-",
                Status = ExportStatus.Failed,
                StatusReason = ex.Message,
            };
        }
    }

    private static ClientPlan PrepareClientPlan(string clientKey, AppConfig config)
    {
        config.GetClients().TryGetValue(clientKey, out var client);
        if (!KeyToLanguage.TryGetValue(clientKey, out var exportLanguage))
            exportLanguage = Language.English;

        bool isInternational = InternationalKeys.Contains(clientKey);
        string? gamePath = !string.IsNullOrEmpty(client?.Path)
            ? client.Path
            : (isInternational ? config.Global?.Path : null);
        string outputRoot = string.IsNullOrWhiteSpace(config.OutputRoot)
            ? "./rawexd"
            : config.OutputRoot;
        string outputDir = !string.IsNullOrWhiteSpace(client?.OutputDir)
            ? client.OutputDir
            : Path.Combine(outputRoot, clientKey.ToLowerInvariant());

        if (!Path.IsPathRooted(outputDir))
        {
            var baseDir =
                Path.GetDirectoryName(ConfigService.ConfigPath)
                ?? AppDomain.CurrentDomain.BaseDirectory;
            outputDir = Path.Combine(baseDir, outputDir);
        }
        outputDir = Path.GetFullPath(outputDir);

        bool isDetectedPath = false;
        if (string.IsNullOrEmpty(gamePath) || !Directory.Exists(gamePath))
        {
            ConsoleLogger.TagInfo(clientKey, Localizer.T(MsgKey.PathDetecting));
            gamePath = new GamePathDetector().Detect(isInternational);
            if (string.IsNullOrEmpty(gamePath))
            {
                ConsoleLogger.TagError(clientKey, Localizer.T(MsgKey.PathNotFoundSkipping));
                return new ClientPlan
                {
                    ClientKey = clientKey,
                    ExportLanguage = exportLanguage,
                    IsInternational = isInternational,
                    OutputDir = outputDir,
                    EarlyFailure = new ClientExportResult
                    {
                        ClientKey = clientKey,
                        LanguageName = exportLanguage.ToString(),
                        Status = ExportStatus.Skipped,
                        StatusReason = Localizer.T(MsgKey.PathNotFoundSkipping),
                    },
                };
            }
            ConsoleLogger.TagSuccess(clientKey, Localizer.TF(MsgKey.PathAutoDetected, gamePath));
            isDetectedPath = true;
        }

        var combinedPath = Path.Combine(gamePath, "game", "sqpack");
        if (Directory.Exists(combinedPath))
            gamePath = combinedPath;
        else
        {
            var directSqpack = Path.Combine(gamePath, "sqpack");
            if (Directory.Exists(directSqpack))
                gamePath = directSqpack;
        }

        if (string.IsNullOrEmpty(outputDir))
        {
            ConsoleLogger.TagError(clientKey, Localizer.T(MsgKey.ConfigMissingOutputDir));
            return new ClientPlan
            {
                ClientKey = clientKey,
                ExportLanguage = exportLanguage,
                IsInternational = isInternational,
                EarlyFailure = new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = exportLanguage.ToString(),
                    Status = ExportStatus.Skipped,
                    StatusReason = Localizer.T(MsgKey.ConfigMissingOutputDir),
                },
            };
        }

        if (!Directory.Exists(gamePath))
        {
            var msg = Localizer.TF(MsgKey.ClientPathNotExist, gamePath);
            ConsoleLogger.TagError(clientKey, msg);
            return new ClientPlan
            {
                ClientKey = clientKey,
                ExportLanguage = exportLanguage,
                IsInternational = isInternational,
                OutputDir = outputDir,
                EarlyFailure = new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = exportLanguage.ToString(),
                    Status = ExportStatus.Skipped,
                    StatusReason = msg,
                },
            };
        }

        string? detectedVersion = ReadGameVersion(gamePath);
        string? schemaVersion = isInternational
            ? client?.SchemaVersion ?? config.Global?.SchemaVersion
            : client?.SchemaVersion;
        if (
            string.IsNullOrEmpty(schemaVersion)
            || schemaVersion.Equals("latest", StringComparison.OrdinalIgnoreCase)
        )
        {
            schemaVersion = string.IsNullOrEmpty(detectedVersion) ? "latest" : detectedVersion;
        }

        return new ClientPlan
        {
            ClientKey = clientKey,
            ExportLanguage = exportLanguage,
            IsInternational = isInternational,
            GamePath = gamePath,
            IsDetectedPath = isDetectedPath,
            OutputDir = outputDir,
            DetectedVersion = detectedVersion,
            SchemaVersion = schemaVersion,
        };
    }

    private static ClientExportResult ExecuteClientExport(
        ClientPlan plan,
        AppConfig config,
        GameData lumina,
        Dictionary<string, ExdSchema> schemas,
        string actualSchemaName,
        SharedGroupCache groupCache,
        int remainingSharedConsumers,
        List<string> cmdFilters,
        int maxSheetParallelism,
        bool useHexcode,
        bool clear,
        bool skipOffset,
        ref bool savedGlobalDetectedPath
    )
    {
        var clientKey = plan.ClientKey;
        var exportLanguage = plan.ExportLanguage;
        var outputDir = plan.OutputDir;
        var detectedVersion = plan.DetectedVersion;

        try
        {
            var stopwatch = Stopwatch.StartNew();

            Directory.CreateDirectory(outputDir);
            bool isHdd = StorageHelper.IsRotationalDisk(outputDir);

            if (clear)
            {
                try
                {
                    var clearResult = ClearDirectory(clientKey, outputDir, isHdd);
                    if (clearResult == ClearResult.Cancelled)
                    {
                        return new ClientExportResult
                        {
                            ClientKey = clientKey,
                            LanguageName = exportLanguage.ToString(),
                            ClientVersion = detectedVersion ?? "-",
                            ActualSchema = actualSchemaName,
                            OutputDir = outputDir,
                            Status = ExportStatus.Skipped,
                            StatusReason = Localizer.T(MsgKey.ClearCancelled),
                        };
                    }
                    if (clearResult == ClearResult.Failed)
                    {
                        return new ClientExportResult
                        {
                            ClientKey = clientKey,
                            LanguageName = exportLanguage.ToString(),
                            ClientVersion = detectedVersion ?? "-",
                            ActualSchema = actualSchemaName,
                            OutputDir = outputDir,
                            Status = ExportStatus.Failed,
                            StatusReason = Localizer.TF(MsgKey.ClearFailed, outputDir),
                        };
                    }
                }
                catch (Exception ex)
                {
                    var err = Localizer.TF(MsgKey.PrepareOutputException, ex.Message);
                    ConsoleLogger.TagError(clientKey, err);
                    return new ClientExportResult
                    {
                        ClientKey = clientKey,
                        LanguageName = exportLanguage.ToString(),
                        ClientVersion = detectedVersion ?? "-",
                        ActualSchema = actualSchemaName,
                        OutputDir = outputDir,
                        Status = ExportStatus.Failed,
                        StatusReason = err,
                    };
                }
            }

            var sheetNames = lumina.Excel.SheetNames.ToList();
            if (cmdFilters.Count > 0)
            {
                sheetNames =
                [
                    .. sheetNames.Where(s =>
                        cmdFilters.Any(f => s.Equals(f, StringComparison.OrdinalIgnoreCase))
                    ),
                ];
            }

            if (sheetNames.Count == 0)
            {
                ConsoleLogger.TagWarn(clientKey, Localizer.T(MsgKey.NoSheetsMatched));
                return new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = exportLanguage.ToString(),
                    ClientVersion = detectedVersion ?? "-",
                    ActualSchema = actualSchemaName,
                    OutputDir = outputDir,
                    ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
                };
            }

            Dictionary<string, long>? existingSizes = clear
                ? null
                : StorageHelper.EnumerateExistingCsvSizes(outputDir);
            if (existingSizes is { Count: 0 })
                existingSizes = null;

            StorageHelper.PreCreateSubdirectories(outputDir, sheetNames);

            var exporter = new ExdExporter(useHexcode, !skipOffset);
            var failedSheets = new ConcurrentBag<(string name, string error)>();
            var cellErrors = new ConcurrentBag<CellReadError>();
            int successCount = 0;
            int completedCount = 0;
            int totalSheets = sheetNames.Count;

            if (isHdd)
            {
                // 对机械硬盘按路径聚簇排序并用受控 I/O 线程消费，消除随机寻道风暴
                sheetNames.Sort(StringComparer.OrdinalIgnoreCase);
                RunHddPipeline(
                    clientKey,
                    lumina,
                    outputDir,
                    exportLanguage,
                    schemas,
                    exporter,
                    groupCache,
                    remainingSharedConsumers,
                    sheetNames,
                    existingSizes,
                    maxSheetParallelism,
                    failedSheets,
                    cellErrors,
                    ref successCount,
                    ref completedCount,
                    totalSheets
                );
            }
            else
            {
                RunSsdParallel(
                    clientKey,
                    lumina,
                    outputDir,
                    exportLanguage,
                    schemas,
                    exporter,
                    groupCache,
                    remainingSharedConsumers,
                    sheetNames,
                    existingSizes,
                    maxSheetParallelism,
                    failedSheets,
                    cellErrors,
                    ref successCount,
                    ref completedCount,
                    totalSheets
                );
            }

            ConsoleLogger.ClearProgress();
            stopwatch.Stop();

            var displayVersion = detectedVersion ?? "-";
            if (failedSheets.IsEmpty)
            {
                ConsoleLogger.TagSuccess(
                    clientKey,
                    Localizer.TF(
                        MsgKey.ExportDone,
                        successCount,
                        displayVersion,
                        actualSchemaName,
                        stopwatch.Elapsed.TotalSeconds,
                        outputDir
                    )
                );
            }
            else
            {
                ConsoleLogger.TagWarn(
                    clientKey,
                    Localizer.TF(
                        MsgKey.ExportDonePartial,
                        successCount,
                        failedSheets.Count,
                        displayVersion,
                        actualSchemaName,
                        stopwatch.Elapsed.TotalSeconds,
                        outputDir
                    )
                );
                foreach (var (name, error) in failedSheets.Take(10))
                    ConsoleLogger.Detail($"{name}: {error}");
                if (failedSheets.Count > 10)
                    ConsoleLogger.Detail(
                        Localizer.TF(MsgKey.ExportFailedMore, failedSheets.Count - 10)
                    );
            }

            if (!cellErrors.IsEmpty)
            {
                var affectedSheets = cellErrors.Select(e => e.SheetName).Distinct().Count();
                ConsoleLogger.TagWarn(
                    clientKey,
                    Localizer.TF(MsgKey.CellErrorsSummary, affectedSheets, cellErrors.Count)
                );
                foreach (var err in cellErrors.Take(5))
                {
                    ConsoleLogger.Detail(
                        Localizer.TF(
                            MsgKey.CellErrorItem,
                            err.SheetName,
                            err.RowId,
                            err.ColumnOffset,
                            err.ColumnType,
                            err.Message
                        )
                    );
                }
            }

            if (successCount > 0 && plan.IsDetectedPath)
            {
                if (!plan.IsInternational || !savedGlobalDetectedPath)
                {
                    SaveDetectedPath(config, clientKey, plan.IsInternational, plan.GamePath!);
                    if (plan.IsInternational)
                        savedGlobalDetectedPath = true;
                }
            }

            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                ClientVersion = detectedVersion ?? "-",
                ActualSchema = actualSchemaName,
                OutputDir = outputDir,
                SuccessCount = successCount,
                FailedCount = failedSheets.Count,
                ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
            };
        }
        catch (Exception ex)
        {
            ConsoleLogger.ClearProgress();
            ConsoleLogger.TagError(clientKey, Localizer.TF(MsgKey.RunFailed, ex.Message));
            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                ClientVersion = detectedVersion ?? "-",
                Status = ExportStatus.Failed,
                StatusReason = ex.Message,
            };
        }
    }

    private static void RunSsdParallel(
        string clientKey,
        GameData lumina,
        string outputDir,
        Language exportLanguage,
        Dictionary<string, ExdSchema> schemas,
        ExdExporter exporter,
        SharedGroupCache groupCache,
        int remainingSharedConsumers,
        List<string> sheetNames,
        Dictionary<string, long>? existingSizes,
        int maxSheetParallelism,
        ConcurrentBag<(string name, string error)> failedSheets,
        ConcurrentBag<CellReadError> cellErrors,
        ref int successCount,
        ref int completedCount,
        int totalSheets
    )
    {
        int localSuccess = 0;
        int localCompleted = 0;

        Parallel.ForEach(
            Partitioner.Create(sheetNames, EnumerablePartitionerOptions.NoBuffering),
            new ParallelOptions { MaxDegreeOfParallelism = maxSheetParallelism },
            () => (Buf: new Utf8CsvBuffer(), Scratch: new byte[65536]),
            (sheetName, _, worker) =>
            {
                try
                {
                    var targetPath = Path.Combine(
                        outputDir,
                        sheetName.Replace('/', Path.DirectorySeparatorChar) + ".csv"
                    );
                    long? oldSize =
                        existingSizes != null && existingSizes.TryGetValue(targetPath, out var sz)
                            ? sz
                            : null;

                    var payload = GenerateOrGetSharedBytes(
                        lumina,
                        sheetName,
                        exportLanguage,
                        schemas,
                        exporter,
                        groupCache,
                        remainingSharedConsumers,
                        worker.Buf,
                        cellErrors.Add,
                        out var sharedBytes
                    );

                    StorageHelper.WriteIfChanged(
                        targetPath,
                        sharedBytes != null ? sharedBytes : payload,
                        oldSize,
                        worker.Scratch
                    );
                    Interlocked.Increment(ref localSuccess);
                }
                catch (Exception ex)
                {
                    failedSheets.Add((sheetName, ex.Message));
                }
                finally
                {
                    int done = Interlocked.Increment(ref localCompleted);
                    ReportProgress(clientKey, done, totalSheets);
                }
                return worker;
            },
            _ => { }
        );

        successCount = localSuccess;
        completedCount = localCompleted;
    }

    private static void RunHddPipeline(
        string clientKey,
        GameData lumina,
        string outputDir,
        Language exportLanguage,
        Dictionary<string, ExdSchema> schemas,
        ExdExporter exporter,
        SharedGroupCache groupCache,
        int remainingSharedConsumers,
        List<string> sheetNames,
        Dictionary<string, long>? existingSizes,
        int maxSheetParallelism,
        ConcurrentBag<(string name, string error)> failedSheets,
        ConcurrentBag<CellReadError> cellErrors,
        ref int successCount,
        ref int completedCount,
        int totalSheets
    )
    {
        int ioThreadCount = existingSizes != null ? 4 : 2;
        using var queue = new BlockingCollection<PendingIoItem>(boundedCapacity: 256);

        int localSuccess = 0;
        int localCompleted = 0;
        var ioWorkers = new Task[ioThreadCount];

        for (int i = 0; i < ioThreadCount; i++)
        {
            ioWorkers[i] = Task.Factory.StartNew(
                () =>
                {
                    var scratch = existingSizes != null ? new byte[65536] : [];
                    foreach (var item in queue.GetConsumingEnumerable())
                    {
                        try
                        {
                            var span = item.Data.AsSpan(0, item.Length);
                            if (item.ExistingSize.HasValue)
                                StorageHelper.WriteIfChanged(
                                    item.TargetPath,
                                    span,
                                    item.ExistingSize,
                                    scratch
                                );
                            else
                                StorageHelper.WriteFileDirect(item.TargetPath, span);

                            Interlocked.Increment(ref localSuccess);
                        }
                        catch (Exception ex)
                        {
                            failedSheets.Add((item.SheetName, ex.Message));
                        }
                        finally
                        {
                            if (item.ReturnToPool)
                                ArrayPool<byte>.Shared.Return(item.Data);
                            int done = Interlocked.Increment(ref localCompleted);
                            ReportProgress(clientKey, done, totalSheets);
                        }
                    }
                },
                TaskCreationOptions.LongRunning
            );
        }

        Parallel.ForEach(
            Partitioner.Create(sheetNames, EnumerablePartitionerOptions.NoBuffering),
            new ParallelOptions { MaxDegreeOfParallelism = maxSheetParallelism },
            () => new Utf8CsvBuffer(),
            (sheetName, _, buf) =>
            {
                try
                {
                    var targetPath = Path.Combine(
                        outputDir,
                        sheetName.Replace('/', Path.DirectorySeparatorChar) + ".csv"
                    );
                    long? oldSize =
                        existingSizes != null && existingSizes.TryGetValue(targetPath, out var sz)
                            ? sz
                            : null;

                    var span = GenerateOrGetSharedBytes(
                        lumina,
                        sheetName,
                        exportLanguage,
                        schemas,
                        exporter,
                        groupCache,
                        remainingSharedConsumers,
                        buf,
                        cellErrors.Add,
                        out var sharedBytes
                    );

                    if (sharedBytes != null)
                    {
                        queue.Add(
                            new PendingIoItem(
                                sheetName,
                                targetPath,
                                sharedBytes,
                                sharedBytes.Length,
                                oldSize,
                                ReturnToPool: false
                            )
                        );
                    }
                    else
                    {
                        var pooled = ArrayPool<byte>.Shared.Rent(span.Length);
                        try
                        {
                            span.CopyTo(pooled);
                            queue.Add(
                                new PendingIoItem(
                                    sheetName,
                                    targetPath,
                                    pooled,
                                    span.Length,
                                    oldSize,
                                    ReturnToPool: true
                                )
                            );
                        }
                        catch
                        {
                            ArrayPool<byte>.Shared.Return(pooled);
                            throw;
                        }
                    }
                }
                catch (Exception ex)
                {
                    failedSheets.Add((sheetName, ex.Message));
                    int done = Interlocked.Increment(ref localCompleted);
                    ReportProgress(clientKey, done, totalSheets);
                }

                return buf;
            },
            _ => { }
        );

        queue.CompleteAdding();
        Task.WaitAll(ioWorkers);

        successCount = localSuccess;
        completedCount = localCompleted;
    }

    private static ReadOnlySpan<byte> GenerateOrGetSharedBytes(
        GameData lumina,
        string sheetName,
        Language exportLanguage,
        Dictionary<string, ExdSchema> schemas,
        ExdExporter exporter,
        SharedGroupCache groupCache,
        int remainingSharedConsumers,
        Utf8CsvBuffer buffer,
        Action<CellReadError> onCellError,
        out byte[]? sharedBytes
    )
    {
        if (groupCache.LanguageNonePayloads.TryGetValue(sheetName, out var cached))
        {
            if (Interlocked.Decrement(ref cached.RemainingConsumers) <= 0)
                groupCache.LanguageNonePayloads.TryRemove(sheetName, out _);

            sharedBytes = cached.Data;
            return cached.Data;
        }

        var layout = groupCache.Layouts.GetOrAdd(
            sheetName,
            s =>
            {
                var baseSheetName = s.Contains('/') ? s[(s.LastIndexOf('/') + 1)..] : s;
                schemas.TryGetValue(baseSheetName, out var schema);
                return exporter.LoadLayout(lumina, s, schema);
            }
        );

        var actualLanguage = ExdExporter.ResolveLanguage(layout.HeaderFile, exportLanguage);

        buffer.Reset();
        exporter.ExportToBuffer(lumina, sheetName, actualLanguage, layout, buffer, onCellError);

        if (actualLanguage == Language.None && remainingSharedConsumers > 0)
        {
            var copy = buffer.WrittenSpan.ToArray();
            groupCache.LanguageNonePayloads.TryAdd(
                sheetName,
                new SharedPayloadEntry(copy, remainingSharedConsumers)
            );
            sharedBytes = copy;
            return copy;
        }

        sharedBytes = null;
        return buffer.WrittenSpan;
    }

    private static void ReportProgress(string clientKey, int done, int totalSheets)
    {
        ConsoleLogger.Progress(
            clientKey,
            () => Localizer.TF(MsgKey.ExportingProgress, done, totalSheets, done * 100 / totalSheets)
        );
    }

    private static string? ReadGameVersion(string gamePath)
    {
        try
        {
            var parentDir = Directory.GetParent(gamePath);
            var verFile = parentDir != null ? Path.Combine(parentDir.FullName, "ffxivgame.ver") : null;
            if (verFile != null && File.Exists(verFile))
            {
                var ver = File.ReadAllText(verFile).Trim();
                return string.IsNullOrEmpty(ver) ? null : ver;
            }
        }
        catch { }
        return null;
    }

    private static string GetDirName(string dirPath, string fallback)
    {
        if (string.IsNullOrEmpty(dirPath))
            return fallback;
        var name = Path.GetFileName(
            dirPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        );
        return string.IsNullOrEmpty(name) ? fallback : name;
    }

    private static void SaveDetectedPath(
        AppConfig config,
        string clientKey,
        bool isInternational,
        string path
    )
    {
        try
        {
            if (isInternational)
            {
                config.Global ??= new();
                config.Global.Path = path;
            }
            else
            {
                config.Clients ??= new(StringComparer.OrdinalIgnoreCase);
                if (!config.Clients.TryGetValue(clientKey, out var target) || target == null)
                {
                    target = new ClientConfig();
                    config.Clients[clientKey] = target;
                }
                target.Path = path;
            }
            ConfigService.SaveConfig(config);
            ConsoleLogger.TagSuccess(clientKey, Localizer.T(MsgKey.SavingDetectedPath));
        }
        catch (Exception ex)
        {
            ConsoleLogger.TagWarn(clientKey, Localizer.TF(MsgKey.SaveConfigWarning, ex.Message));
        }
    }

    private static Dictionary<string, ExdSchema> LoadOrFetchSchemas(
        string schemaVersion,
        string schemaRoot,
        out string finalSchemaDir
    )
    {
        finalSchemaDir = "";
        bool isLatest = schemaVersion.Equals("latest", StringComparison.OrdinalIgnoreCase);
        var targetDir = Path.Combine(schemaRoot, schemaVersion);

        if (!isLatest && Directory.Exists(targetDir))
        {
            var loaded = SchemaService.LoadSchemas(targetDir);
            if (loaded.Count > 0)
            {
                finalSchemaDir = targetDir;
                return loaded;
            }
        }

        var missingMarker = Path.Combine(schemaRoot, $".missing_{schemaVersion}");
        bool skipVersionCheck =
            !isLatest
            && File.Exists(missingMarker)
            && (DateTime.UtcNow - File.GetLastWriteTimeUtc(missingMarker)).TotalHours < 24;

        string? lastError = null;
        if (!skipVersionCheck)
        {
            ConsoleLogger.Progress("Schema", Localizer.TF(MsgKey.SchemaFetching, schemaVersion));
            var fetchResult = SchemaUpdater.DownloadAndExtractSchema(schemaVersion, schemaRoot).Result;

            if (fetchResult.Success)
            {
                var loaded = SchemaService.LoadSchemas(fetchResult.DirectoryPath!);
                if (loaded.Count > 0)
                {
                    finalSchemaDir = fetchResult.DirectoryPath!;
                    ConsoleLogger.ClearProgress();
                    if (!fetchResult.FromCache && !fetchResult.IsUpToDate)
                    {
                        ConsoleLogger.TagSuccess(
                            "Schema",
                            Localizer.TF(MsgKey.SchemaLoadedUpdated, schemaVersion, loaded.Count)
                        );
                    }
                    return loaded;
                }
            }

            lastError = fetchResult.ErrorMessage;
            if (!isLatest && fetchResult.IsNotFound)
            {
                try
                {
                    File.WriteAllText(missingMarker, DateTime.UtcNow.ToString("O"));
                }
                catch { }
            }
        }

        if (!isLatest)
        {
            ConsoleLogger.Progress("Schema", Localizer.TF(MsgKey.SchemaFetching, "latest"));
            var latestResult = SchemaUpdater.DownloadAndExtractSchema("latest", schemaRoot).Result;

            if (latestResult.Success)
            {
                var loadedLatest = SchemaService.LoadSchemas(latestResult.DirectoryPath!);
                if (loadedLatest.Count > 0)
                {
                    finalSchemaDir = latestResult.DirectoryPath!;
                    ConsoleLogger.TagWarn(
                        "Schema",
                        Localizer.TF(MsgKey.SchemaFallbackUsed, schemaVersion, loadedLatest.Count)
                    );
                    return loadedLatest;
                }
            }
            lastError ??= latestResult.ErrorMessage;
        }

        var latestLocalDir = Path.Combine(schemaRoot, "latest");
        if (Directory.Exists(latestLocalDir))
        {
            var loadedLocalLatest = SchemaService.LoadSchemas(latestLocalDir);
            if (loadedLocalLatest.Count > 0)
            {
                finalSchemaDir = latestLocalDir;
                ConsoleLogger.TagWarn(
                    "Schema",
                    Localizer.TF(
                        MsgKey.SchemaOnlineFailedLocalFallback,
                        lastError ?? "network error",
                        "latest"
                    )
                );
                return loadedLocalLatest;
            }
        }

        ConsoleLogger.ClearProgress();
        return [];
    }

    private enum ClearResult
    {
        Success,
        Cancelled,
        Failed,
    }

    private static ClearResult ClearDirectory(string clientKey, string path, bool isHdd)
    {
        if (!Directory.Exists(path))
            return ClearResult.Success;

        var allFiles = new List<string>(8192);
        var nonCsvFiles = new List<string>();
        var enumOpts = new EnumerationOptions { RecurseSubdirectories = true };
        var fsEnum = new FileSystemEnumerable<string>(
            path,
            (ref FileSystemEntry entry) => entry.ToFullPath(),
            enumOpts
        )
        {
            ShouldIncludePredicate = (ref FileSystemEntry entry) => !entry.IsDirectory,
        };

        foreach (var file in fsEnum)
        {
            allFiles.Add(file);
            if (!file.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                nonCsvFiles.Add(file);
        }

        var dir = new DirectoryInfo(path);
        if (allFiles.Count == 0)
        {
            foreach (var subDir in dir.GetDirectories())
                subDir.Delete(true);
            return ClearResult.Success;
        }

        if (nonCsvFiles.Count > 0)
        {
            bool confirmed = false;
            ConsoleLogger.WithLock(() =>
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(
                    $"[{clientKey}] ! {Localizer.TF(MsgKey.ClearWarnNonCsv, nonCsvFiles.Count, path)}"
                );
                Console.ResetColor();
                Console.WriteLine($"  {Localizer.T(MsgKey.ClearWarnListHeader)}");

                foreach (var file in nonCsvFiles.Take(20))
                    Console.WriteLine($"  - {Path.GetRelativePath(path, file)}");

                if (nonCsvFiles.Count > 20)
                    Console.WriteLine(
                        $"  {Localizer.TF(MsgKey.ClearWarnMoreFiles, nonCsvFiles.Count - 20)}"
                    );

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write($"  {Localizer.T(MsgKey.ClearConfirmPrompt)}");
                Console.ResetColor();

                var response = Console.ReadLine()?.Trim().ToLower();
                confirmed = response is "y" or "yes";
            });

            if (!confirmed)
            {
                ConsoleLogger.TagWarn(clientKey, Localizer.T(MsgKey.ClearCancelled));
                return ClearResult.Cancelled;
            }
        }

        var lockedFiles = new ConcurrentBag<string>();
        int deleteParallelism = isHdd ? 4 : Math.Max(1, Environment.ProcessorCount);

        Parallel.ForEach(
            allFiles,
            new ParallelOptions { MaxDegreeOfParallelism = deleteParallelism },
            file =>
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    lockedFiles.Add($"{Path.GetRelativePath(path, file)} ({ex.Message})");
                }
            }
        );

        foreach (var subDir in dir.GetDirectories())
        {
            try
            {
                subDir.Delete(true);
            }
            catch (Exception ex)
            {
                ConsoleLogger.TagWarn(
                    clientKey,
                    Localizer.TF(MsgKey.ClearSubdirFailed, subDir.Name, ex.Message)
                );
            }
        }

        if (!lockedFiles.IsEmpty)
        {
            ConsoleLogger.TagError(
                clientKey,
                Localizer.TF(MsgKey.ClearLockedSummary, lockedFiles.Count)
            );
            foreach (var f in lockedFiles.Take(10))
                ConsoleLogger.Detail(f);
            if (lockedFiles.Count > 10)
                ConsoleLogger.Detail("...");
            return ClearResult.Failed;
        }

        return ClearResult.Success;
    }
}
