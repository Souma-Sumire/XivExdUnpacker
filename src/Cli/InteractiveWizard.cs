using System.Runtime.InteropServices;
using XivExdUnpacker.Services;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src.Cli;

public static class InteractiveWizard
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint GetConsoleProcessList([Out] uint[] processList, uint processCount);

    public static void WaitForExitIfStandaloneConsole()
    {
        if (!OperatingSystem.IsWindows() || Console.IsInputRedirected || Console.IsOutputRedirected)
            return;

        try
        {
            var processIds = new uint[2];
            if (GetConsoleProcessList(processIds, (uint)processIds.Length) == 1)
            {
                Console.WriteLine();
                Console.Write(Localizer.T(MsgKey.PressAnyKeyToExit));
                Console.ReadKey(true);
            }
        }
        catch
        {
            // Ignore console query/read failures on exit
        }
    }

    public static bool RunConfigWizard()
    {
        const string defaultUiLang = "en";
        Localizer.SetLanguage(defaultUiLang);

        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.WizardConfigTitle));
        Console.WriteLine(Localizer.T(MsgKey.WizardConfigIntro));
        Console.WriteLine();

        string uiLangInput = PromptText(
            Localizer.TF(MsgKey.WizardPromptUiLanguage, defaultUiLang),
            defaultUiLang
        );
        string uiLanguage = uiLangInput.Trim().ToLowerInvariant() is "chs" or "zh" or "zh-cn" or "cn"
            ? "chs"
            : "en";
        Localizer.SetLanguage(uiLanguage);

        string outputRoot = PromptText(
            Localizer.TF(MsgKey.WizardPromptOutputRoot, "./rawexd"),
            "./rawexd"
        );

        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.WizardScanningPaths));
        var detected = new GamePathDetector().DetectAll();
        detected.TryGetValue(ClientRegion.Global, out var detectedGlobal);
        detected.TryGetValue(ClientRegion.Chs, out var detectedChs);
        detected.TryGetValue(ClientRegion.Ko, out var detectedKo);
        detected.TryGetValue(ClientRegion.Tc, out var detectedTc);

        string? globalPath = PromptOptionalPath(
            Localizer.T(MsgKey.WizardPromptGlobalPath),
            detectedGlobal
        );
        string? chsPath = PromptOptionalPath(
            Localizer.T(MsgKey.WizardPromptChsPath),
            detectedChs
        );
        string? koPath = PromptOptionalPath(
            Localizer.T(MsgKey.WizardPromptKoPath),
            detectedKo
        );
        string? tcPath = PromptOptionalPath(
            Localizer.T(MsgKey.WizardPromptTcPath),
            detectedTc
        );

        var config = new AppConfig
        {
            UiLanguage = uiLanguage,
            OutputRoot = outputRoot,
            Global = ToClientConfig(globalPath),
            Chs = ToClientConfig(chsPath),
            Ko = ToClientConfig(koPath),
            Tc = ToClientConfig(tcPath),
        };

        if (config.GetClients().Count == 0)
        {
            Console.WriteLine();
            ConsoleLogger.TopError(Localizer.T(MsgKey.WizardConfigNoClients));
            return false;
        }

        try
        {
            ConfigService.SaveConfig(config);
            Console.WriteLine();
            ConsoleLogger.TopInfo(Localizer.TF(MsgKey.WizardConfigSaved, ConfigService.ConfigPath));
            return true;
        }
        catch (Exception ex)
        {
            ConsoleLogger.TopError(Localizer.TF(MsgKey.ConfigSaveFailed, ex.Message));
            return false;
        }
    }

    public static CommandLineArgs? RunExportWizard(AppConfig config)
    {
        var clients = config.GetClients();
        if (clients.Count == 0)
            return null;

        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.WizardExportTitle));
        Console.WriteLine(Localizer.T(MsgKey.WizardAvailableLanguages));

        foreach (var (key, client) in clients)
        {
            string status = GetClientPathStatus(client);
            Console.WriteLine($"  - {key} ({status})");
        }

        List<string> selectedLanguages;
        while (true)
        {
            Console.Write(Localizer.T(MsgKey.WizardPromptLanguages));
            string? line = Console.ReadLine();
            if (line == null)
                return null;

            var tokens = SplitTokens(line);
            if (tokens.Count == 0 || tokens.Contains("all", StringComparer.OrdinalIgnoreCase))
            {
                selectedLanguages = ["all"];
                break;
            }

            bool allValid = tokens.All(t => clients.ContainsKey(t));
            if (allValid)
            {
                selectedLanguages = tokens;
                break;
            }

            ConsoleLogger.TopError(
                Localizer.TF(MsgKey.WizardInvalidLanguages, string.Join(", ", clients.Keys))
            );
        }

        Console.Write(Localizer.T(MsgKey.WizardPromptSheets));
        string? sheetsLine = Console.ReadLine();
        if (sheetsLine == null)
            return null;
        var sheets = SplitTokens(sheetsLine);

        bool hexCode = PromptYesNo(Localizer.T(MsgKey.WizardPromptHexcode), defaultValue: false);
        bool clear = PromptYesNo(Localizer.T(MsgKey.WizardPromptClear), defaultValue: false);

        return new CommandLineArgs
        {
            Languages = selectedLanguages,
            Sheets = sheets.Count > 0 ? sheets : null,
            HexCode = hexCode,
            Clear = clear,
            SkipOffset = false,
        };
    }

    private static string? PromptOptionalPath(string label, string? detectedPath)
    {
        while (true)
        {
            Console.WriteLine();
            Console.WriteLine(label);
            if (!string.IsNullOrEmpty(detectedPath))
                Console.Write(Localizer.TF(MsgKey.WizardPathDetected, detectedPath));
            else
                Console.Write(Localizer.T(MsgKey.WizardPathNotDetected));

            string? input = Console.ReadLine()?.Trim();
            if (input == null)
                return null;

            if (string.IsNullOrEmpty(input))
                return detectedPath;

            if (input == "-" || input.Equals("skip", StringComparison.OrdinalIgnoreCase))
                return null;

            string normalized = input.Trim('"');
            if (IsValidGamePath(normalized))
                return normalized;

            ConsoleLogger.TopError(Localizer.TF(MsgKey.WizardPathInvalidWarning, normalized));
            if (!PromptYesNo(Localizer.T(MsgKey.WizardPromptRetryPath), defaultValue: true))
                return normalized;
        }
    }

    private static ClientConfig ToClientConfig(string? path)
    {
        return new ClientConfig
        {
            Path = path ?? "",
            OutputDir = "",
            SchemaVersion = "",
        };
    }

    private static string GetClientPathStatus(ClientConfig client)
    {
        if (string.IsNullOrWhiteSpace(client.Path))
            return Localizer.T(MsgKey.WizardLangUnconfigured);

        return IsValidGamePath(client.Path)
            ? Localizer.T(MsgKey.WizardLangValid)
            : Localizer.T(MsgKey.WizardLangInvalidPath);
    }

    private static bool IsValidGamePath(string path)
    {
        if (!Directory.Exists(path))
            return false;

        return Directory.Exists(Path.Combine(path, "game", "sqpack"))
            || Directory.Exists(Path.Combine(path, "sqpack"))
            || Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
                .Equals("sqpack", StringComparison.OrdinalIgnoreCase);
    }

    private static string PromptText(string prompt, string defaultValue)
    {
        Console.Write(prompt);
        string? input = Console.ReadLine()?.Trim();
        return string.IsNullOrEmpty(input) ? defaultValue : input;
    }

    private static bool PromptYesNo(string prompt, bool defaultValue)
    {
        Console.Write(prompt);
        string? input = Console.ReadLine()?.Trim();
        if (string.IsNullOrEmpty(input))
            return defaultValue;

        return input.Equals("y", StringComparison.OrdinalIgnoreCase)
            || input.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> SplitTokens(string input)
    {
        return input
            .Split([' ', '\t', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
    }
}
