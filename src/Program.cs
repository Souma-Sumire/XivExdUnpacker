using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using Lumina;
using XivExdUnpacker.src.Cli;
using XivExdUnpacker.src.Core;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src;

class Program
{
    static void Main(string[] args)
    {
        Localizer.TryLoadFromConfigIfExists();

        if (OperatingSystem.IsWindows())
        {
            try
            {
                Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.AboveNormal;
            }
            catch
            {
                // Ignore priority elevation failure; does not affect unpacking
            }
        }
        Console.OutputEncoding = Encoding.UTF8;

        var parsedArgs = CliOptions.Parse(args);
        if (parsedArgs.ShowHelp || args.Length == 0)
        {
            CliOptions.ShowHelp();
            return;
        }

        if (parsedArgs.Languages == null || parsedArgs.Languages.Count == 0)
        {
            ConsoleLogger.TopError(Localizer.T(MsgKey.ErrorMissingLanguage));
            ConsoleLogger.TopInfo(Localizer.T(MsgKey.ErrorUseHelp));
            return;
        }

        var config = ConfigService.LoadConfig();
        if (config.GetClients().Count == 0)
        {
            ConsoleLogger.TopError(Localizer.T(MsgKey.ErrorNoClientsInConfig));
            return;
        }

        List<string> filters = parsedArgs.Sheets ?? [];
        List<string> selectedKeys;

        if (parsedArgs.Languages.Contains("all", StringComparer.OrdinalIgnoreCase))
        {
            selectedKeys = [.. config.GetClients().Keys];
        }
        else
        {
            selectedKeys = [.. parsedArgs.Languages.Where(c => config.GetClients().ContainsKey(c))];
            if (selectedKeys.Count == 0)
            {
                ConsoleLogger.TopError(
                    Localizer.TF(
                        MsgKey.ErrorLanguagesNotFound,
                        string.Join(", ", parsedArgs.Languages)
                    )
                );
                ConsoleLogger.TopInfo(
                    Localizer.TF(
                        MsgKey.AvailableLanguages,
                        string.Join(", ", config.GetClients().Keys)
                    )
                );
                return;
            }
        }

        try
        {
            if (!Console.IsOutputRedirected)
                Console.CursorVisible = true;
        }
        catch { }

        var totalStopwatch = Stopwatch.StartNew();
        var clientResults = new ConcurrentBag<(string key, ClientExportResult result)>();
        var gameDataPool = new Dictionary<string, GameData>(StringComparer.OrdinalIgnoreCase);
        var schemaCache = new ConcurrentDictionary<
            string,
            (string dir, Dictionary<string, ExdSchema> schemas)
        >(StringComparer.OrdinalIgnoreCase);

        int maxSheetParallelism = Math.Max(1, Environment.ProcessorCount);

        foreach (var clientKey in selectedKeys)
        {
            var result = ExportRunner.RunDumpProcess(
                clientKey,
                config,
                filters,
                maxSheetParallelism,
                schemaCache,
                gameDataPool,
                parsedArgs.HexCode,
                parsedArgs.Clear,
                parsedArgs.SkipOffset
            );
            clientResults.Add((clientKey, result));
        }

        totalStopwatch.Stop();

        var orderedResults = selectedKeys
            .Select(k => clientResults.FirstOrDefault(r => r.key == k).result)
            .Where(r => r != null)
            .ToList();

        ConsoleReport.PrintSummary(orderedResults, totalStopwatch.Elapsed);
    }
}
