using System.Diagnostics;
using System.Text;
using XivExdUnpacker.src.Cli;
using XivExdUnpacker.src.Core;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src;

class Program
{
    static int Main(string[] args)
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

        int exitCode = Run(args);
        InteractiveWizard.WaitForExitIfStandaloneConsole();
        return exitCode;
    }

    private static int Run(string[] args)
    {
        AppConfig config;
        CommandLineArgs parsedArgs;

        if (args.Length == 0)
        {
            if (Console.IsInputRedirected)
            {
                CliOptions.ShowHelp();
                return 0;
            }

            if (!ConfigService.Exists && !InteractiveWizard.RunConfigWizard())
                return 1;

            if (!TryLoadValidConfig(out config))
                return 1;

            var wizardArgs = InteractiveWizard.RunExportWizard(config);
            if (wizardArgs == null)
                return 1;
            parsedArgs = wizardArgs;
        }
        else
        {
            parsedArgs = CliOptions.Parse(args);
            if (parsedArgs.ShowHelp)
            {
                CliOptions.ShowHelp();
                return 0;
            }

            if (!string.IsNullOrEmpty(parsedArgs.ErrorMessage))
            {
                ConsoleLogger.TopError(parsedArgs.ErrorMessage);
                ConsoleLogger.TopInfo(Localizer.T(MsgKey.ErrorUseHelp));
                return 1;
            }

            if (parsedArgs.Languages == null || parsedArgs.Languages.Count == 0)
            {
                ConsoleLogger.TopError(Localizer.T(MsgKey.ErrorMissingLanguage));
                ConsoleLogger.TopInfo(Localizer.T(MsgKey.ErrorUseHelp));
                return 1;
            }

            if (!TryLoadValidConfig(out config))
                return 1;
        }

        List<string> filters = parsedArgs.Sheets ?? [];
        List<string> selectedKeys;

        var requestedLanguages = parsedArgs.Languages!;
        if (requestedLanguages.Contains("all", StringComparer.OrdinalIgnoreCase))
        {
            selectedKeys = [.. config.GetClients().Keys];
        }
        else
        {
            selectedKeys = [.. requestedLanguages.Where(c => config.GetClients().ContainsKey(c))];
            if (selectedKeys.Count == 0)
            {
                ConsoleLogger.TopError(
                    Localizer.TF(
                        MsgKey.ErrorLanguagesNotFound,
                        string.Join(", ", requestedLanguages)
                    )
                );
                ConsoleLogger.TopInfo(
                    Localizer.TF(
                        MsgKey.AvailableLanguages,
                        string.Join(", ", config.GetClients().Keys)
                    )
                );
                return 1;
            }
        }

        try
        {
            if (!Console.IsOutputRedirected)
                Console.CursorVisible = true;
        }
        catch { }

        var totalStopwatch = Stopwatch.StartNew();
        int maxSheetParallelism = Math.Max(1, Environment.ProcessorCount);

        var orderedResults = ExportRunner.RunAll(
            selectedKeys,
            config,
            filters,
            maxSheetParallelism,
            parsedArgs.HexCode,
            parsedArgs.Clear,
            parsedArgs.SkipOffset
        );

        totalStopwatch.Stop();

        ConsoleReport.PrintSummary(orderedResults, totalStopwatch.Elapsed);

        bool hasFailures =
            orderedResults.Count == 0
            || orderedResults.Any(r => r.Status != ExportStatus.Completed || r.FailedCount > 0);
        return hasFailures ? 1 : 0;
    }

    private static bool TryLoadValidConfig(out AppConfig config)
    {
        if (!ConfigService.TryLoadConfig(out config, out var configError, out var isNotFound))
        {
            ConsoleLogger.TopError(configError!);
            if (isNotFound)
                ConsoleLogger.TopInfo(Localizer.T(MsgKey.ConfigNotFoundHint));
            return false;
        }

        if (config.GetClients().Count == 0)
        {
            ConsoleLogger.TopError(Localizer.T(MsgKey.ErrorNoClientsInConfig));
            return false;
        }

        return true;
    }
}
