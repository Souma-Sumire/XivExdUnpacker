using XivExdUnpacker.src.I18n;

namespace XivExdUnpacker.src.Cli;

public record CommandLineArgs
{
    public List<string>? Languages { get; init; }
    public List<string>? Sheets { get; init; }
    public bool HexCode { get; init; }
    public bool Clear { get; init; }
    public bool SkipOffset { get; init; }
    public bool ShowHelp { get; init; }
    public string? ErrorMessage { get; init; }
}

public static class CliOptions
{
    public static CommandLineArgs Parse(string[] args)
    {
        var languages = new List<string>();
        var sheets = new List<string>();
        bool hexCode = false;
        bool clear = false;
        bool skipOffset = false;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            switch (arg.ToLowerInvariant())
            {
                case "--help":
                case "-h":
                case "/?":
                    return new CommandLineArgs { ShowHelp = true };

                case "--language":
                case "-l":
                    if (!TryReadValues(args, ref i, languages))
                    {
                        return new CommandLineArgs
                        {
                            ErrorMessage = Localizer.TF(MsgKey.ErrorMissingOptionValue, arg),
                        };
                    }
                    break;

                case "--sheets":
                case "-s":
                    if (!TryReadValues(args, ref i, sheets))
                    {
                        return new CommandLineArgs
                        {
                            ErrorMessage = Localizer.TF(MsgKey.ErrorMissingOptionValue, arg),
                        };
                    }
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
                case "-o":
                    skipOffset = true;
                    break;

                default:
                    return new CommandLineArgs
                    {
                        ErrorMessage = Localizer.TF(MsgKey.ErrorUnknownOption, arg),
                    };
            }
        }

        return new CommandLineArgs
        {
            Languages = languages.Count > 0 ? languages : null,
            Sheets = sheets.Count > 0 ? sheets : null,
            HexCode = hexCode,
            Clear = clear,
            SkipOffset = skipOffset,
            ShowHelp = false,
        };
    }

    private static bool TryReadValues(string[] args, ref int index, List<string> target)
    {
        int startCount = target.Count;
        while (index + 1 < args.Length && !args[index + 1].StartsWith('-'))
        {
            index++;
            target.Add(args[index]);
        }
        return target.Count > startCount;
    }

    public static void ShowHelp()
    {
        Console.WriteLine(Localizer.T(MsgKey.HelpTitle));
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpUsageHeader));
        Console.WriteLine(Localizer.T(MsgKey.HelpUsageLine));
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpRequiredHeader));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionLanguage));
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionsHeader));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionSheets));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionHexcode));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionClear));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionSkipOffset));
        Console.WriteLine(Localizer.T(MsgKey.HelpOptionHelp));
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpExamplesHeader));
        Console.WriteLine(Localizer.T(MsgKey.HelpExampleCn));
        Console.WriteLine("  XivExdUnpacker --language chs");
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpExampleEnHex));
        Console.WriteLine("  XivExdUnpacker --language en --hexcode");
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpExampleEnJaSheets));
        Console.WriteLine("  XivExdUnpacker --language en ja --sheets Action Item");
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpExampleAll));
        Console.WriteLine("  XivExdUnpacker --language all --clear --skip-offset");
        Console.WriteLine();
        Console.WriteLine(Localizer.T(MsgKey.HelpExampleShort));
        Console.WriteLine("  XivExdUnpacker -l chs -s Addon Quest -x -c -o");
    }
}
