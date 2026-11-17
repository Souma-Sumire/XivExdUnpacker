using System.Collections.Concurrent;
using System.Diagnostics;
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
                return new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = exportLanguage.ToString(),
                    Status = ExportStatus.Skipped,
                    StatusReason = Localizer.T(MsgKey.PathNotFoundSkipping),
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
            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                Status = ExportStatus.Skipped,
                StatusReason = Localizer.T(MsgKey.ConfigMissingOutputDir),
            };
        }

        if (!Directory.Exists(gamePath))
        {
            var msg = Localizer.TF(MsgKey.ClientPathNotExist, gamePath);
            ConsoleLogger.TagError(clientKey, msg);
            return new ClientExportResult
            {
                ClientKey = clientKey,
                LanguageName = exportLanguage.ToString(),
                Status = ExportStatus.Skipped,
                StatusReason = msg,
            };
        }

        string? detectedVersion = ReadGameVersion(gamePath);

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

            var schemaRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "schemas");
            Directory.CreateDirectory(schemaRoot);

            string finalSchemaDir;
            Dictionary<string, ExdSchema> schemas;
            if (schemaCache.TryGetValue(schemaVersion, out var cached))
            {
                finalSchemaDir = cached.dir;
                schemas = cached.schemas;
            }
            else
            {
                schemas = LoadOrFetchSchemas(schemaVersion, schemaRoot, out finalSchemaDir);
                if (schemas.Count > 0)
                    schemaCache.TryAdd(schemaVersion, (finalSchemaDir, schemas));
            }

            var actualSchemaName = GetDirName(finalSchemaDir, schemaVersion);
            if (schemas.Count == 0 || string.IsNullOrEmpty(finalSchemaDir))
            {
                ConsoleLogger.TagError(clientKey, Localizer.T(MsgKey.SchemaNoneLoaded));
                return new ClientExportResult
                {
                    ClientKey = clientKey,
                    LanguageName = exportLanguage.ToString(),
                    ClientVersion = detectedVersion ?? "-",
                    ActualSchema = "-",
                    OutputDir = outputDir,
                    Status = ExportStatus.Failed,
                    StatusReason = Localizer.T(MsgKey.SchemaNoneLoaded),
                };
            }

            Directory.CreateDirectory(outputDir);
            if (clear)
            {
                try
                {
                    var clearResult = ClearDirectory(clientKey, outputDir);
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

            var exporter = new ExdExporter(useHexcode, !skipOffset);
            var failedSheets = new ConcurrentBag<(string name, string error)>();
            var cellErrors = new ConcurrentBag<CellReadError>();
            int successCount = 0;
            int completedCount = 0;
            int totalSheets = sheetNames.Count;

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
                        exporter.ExportSheet(
                            lumina,
                            sheetName,
                            outputDir,
                            exportLanguage,
                            schema,
                            cellErrors.Add
                        );
                        Interlocked.Increment(ref successCount);
                    }
                    catch (Exception ex)
                    {
                        failedSheets.Add((sheetName, ex.Message));
                    }
                    finally
                    {
                        int done = Interlocked.Increment(ref completedCount);
                        int pct = done * 100 / totalSheets;
                        ConsoleLogger.Progress(
                            clientKey,
                            Localizer.TF(MsgKey.ExportingProgress, done, totalSheets, pct)
                        );
                    }
                }
            );

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
                    ConsoleLogger.Detail(Localizer.TF(MsgKey.ExportFailedMore, failedSheets.Count - 10));
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

            if (successCount > 0 && isDetectedPath)
                SaveDetectedPath(config, clientKey, isInternational, gamePath);

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

    private static ClearResult ClearDirectory(string clientKey, string path)
    {
        var dir = new DirectoryInfo(path);
        if (!dir.Exists)
            return ClearResult.Success;

        var allFiles = dir.GetFiles("*", SearchOption.AllDirectories).ToList();
        if (allFiles.Count == 0)
        {
            foreach (var subDir in dir.GetDirectories())
                subDir.Delete(true);
            return ClearResult.Success;
        }

        var nonCsvFiles = allFiles
            .Where(f => !f.Extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (nonCsvFiles.Count > 0)
        {
            bool confirmed = false;
            ConsoleLogger.WithLock(() =>
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[{clientKey}] ! {Localizer.TF(MsgKey.ClearWarnNonCsv, nonCsvFiles.Count, path)}");
                Console.ResetColor();
                Console.WriteLine($"  {Localizer.T(MsgKey.ClearWarnListHeader)}");

                foreach (var file in nonCsvFiles.Take(20))
                    Console.WriteLine($"  - {Path.GetRelativePath(path, file.FullName)}");

                if (nonCsvFiles.Count > 20)
                    Console.WriteLine($"  {Localizer.TF(MsgKey.ClearWarnMoreFiles, nonCsvFiles.Count - 20)}");

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

        var lockedFiles = new List<string>();
        foreach (var file in allFiles)
        {
            try
            {
                file.Delete();
            }
            catch (Exception ex)
            {
                lockedFiles.Add($"{Path.GetRelativePath(path, file.FullName)} ({ex.Message})");
            }
        }

        foreach (var subDir in dir.GetDirectories())
        {
            try
            {
                subDir.Delete(true);
            }
            catch (Exception ex)
            {
                ConsoleLogger.TagWarn(clientKey, Localizer.TF(MsgKey.ClearSubdirFailed, subDir.Name, ex.Message));
            }
        }

        if (lockedFiles.Count > 0)
        {
            ConsoleLogger.TagError(clientKey, Localizer.TF(MsgKey.ClearLockedSummary, lockedFiles.Count));
            foreach (var f in lockedFiles.Take(10))
                ConsoleLogger.Detail(f);
            if (lockedFiles.Count > 10)
                ConsoleLogger.Detail("...");
            return ClearResult.Failed;
        }

        return ClearResult.Success;
    }
}
