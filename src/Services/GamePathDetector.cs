using System.Collections.Concurrent;
using Lumina;
using Microsoft.Win32;

namespace XivExdUnpacker.Services;

public enum ClientRegion
{
    Global,
    Chs,
    Ko,
    Tc,
}

public class GamePathDetector
{
    private const int MaxScanDepth = 4;

    private static readonly HashSet<string> SkipDirs = new(StringComparer.OrdinalIgnoreCase)
    {
        "windows",
        "programdata",
        "system volume information",
        "recovery",
        "config.msi",
        "msocache",
        "perflogs",
        "$recycle.bin",
        "users",
        "appdata",
        "node_modules",
        "temp",
        "tmp",
        "windowsapps",
        "wpsystem",
        "common files",
        "microsoft",
        "microsoft.net",
        "windows kits",
        "reference assemblies",
    };

    private static readonly string[] GlobalSuffixes =
    [
        Path.Combine("Program Files (x86)", "SquareEnix", "FINAL FANTASY XIV - A Realm Reborn"),
        Path.Combine("Program Files", "SquareEnix", "FINAL FANTASY XIV - A Realm Reborn"),
        Path.Combine("Steam", "steamapps", "common", "FINAL FANTASY XIV Online"),
        Path.Combine("Program Files (x86)", "Steam", "steamapps", "common", "FINAL FANTASY XIV Online"),
        Path.Combine("Games", "SquareEnix", "FINAL FANTASY XIV - A Realm Reborn"),
    ];

    private static readonly string[] ChinaSuffixes =
    [
        Path.Combine("Program Files (x86)", "盛趣游戏", "最终幻想XIV"),
        Path.Combine("Program Files", "盛趣游戏", "最终幻想XIV"),
        Path.Combine("Program Files (x86)", "盛大游戏", "最终幻想XIV"),
        Path.Combine("Program Files (x86)", "上海数龙科技有限公司", "最终幻想XIV"),
        Path.Combine("Program Files (x86)", "SNDA", "FFXIV"),
        Path.Combine("WeGameApps", "rail_apps", "最终幻想14(2000340)"),
        Path.Combine("WeGame", "rail_apps", "最终幻想14(2000340)"),
        Path.Combine("Games", "FFXIV"),
        Path.Combine("Games", "最终幻想XIV"),
        "FFXIV",
        "最终幻想XIV",
    ];

    public string? Detect(bool isInternational = true)
    {
        var target = isInternational ? ClientRegion.Global : ClientRegion.Chs;
        var results = DetectAll(target);
        return results.TryGetValue(target, out var path) ? path : null;
    }

    public Dictionary<ClientRegion, string> DetectAll(
        ClientRegion? stopWhenFound = null,
        TimeSpan? scanTimeout = null
    )
    {
        var found = new Dictionary<ClientRegion, string>();
        var candidates = CollectFastCandidates();

        foreach (var candidate in candidates)
        {
            if (!IsClientRoot(candidate))
                continue;

            var region = DetectClientRegion(candidate);
            if (region != null)
                found.TryAdd(region.Value, candidate);
        }

        if (stopWhenFound != null && found.ContainsKey(stopWhenFound.Value))
            return found;

        if (found.Count < 4)
        {
            using var cts = new CancellationTokenSource(scanTimeout ?? TimeSpan.FromSeconds(5));
            var roots = GetScanDriveRoots();
            foreach (var dir in ScanClientsBfs(roots, MaxScanDepth, cts.Token))
            {
                var region = DetectClientRegion(dir);
                if (region != null)
                    found.TryAdd(region.Value, dir);
            }
        }

        return found;
    }

    private static List<string> CollectFastCandidates()
    {
        var candidates = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Add(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;
            var clean = Path.GetFullPath(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (seen.Add(clean))
                candidates.Add(clean);
        }

        if (OperatingSystem.IsWindows())
        {
            foreach (var path in GetWindowsRegistryCandidates())
                Add(path);

            foreach (var root in GetScanDriveRoots())
            {
                foreach (var suffix in GlobalSuffixes)
                    Add(Path.Combine(root, suffix));
                foreach (var suffix in ChinaSuffixes)
                    Add(Path.Combine(root, suffix));
            }
        }
        else if (OperatingSystem.IsLinux())
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrEmpty(home))
            {
                Add(Path.Combine(home, ".xlcore", "ffxiv"));
                Add(Path.Combine(home, ".steam", "steam", "steamapps", "common", "FINAL FANTASY XIV Online"));
                Add(Path.Combine(home, ".local", "share", "Steam", "steamapps", "common", "FINAL FANTASY XIV Online"));
            }
        }
        else if (OperatingSystem.IsMacOS())
        {
            Add("/Applications/FINAL FANTASY XIV ONLINE.app/Contents/SharedSupport/finalfantasyxiv/support/ffxiv");
            Add("/Applications/FINAL FANTASY XIV ONLINE.app/Contents/SharedSupport/finalfantasyxiv");

            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrEmpty(home))
            {
                Add(Path.Combine(home, "Library", "Application Support", "FINAL FANTASY XIV ONLINE", "Bottles", "published_Final_Fantasy", "drive_c", "Program Files (x86)", "SquareEnix", "FINAL FANTASY XIV - A Realm Reborn"));
                Add(Path.Combine(home, "Library", "Application Support", "XIV on Mac", "ffxiv"));
                Add(Path.Combine(home, "Library", "Application Support", "XIV on Mac", "game"));
                Add(Path.Combine(home, "Library", "Application Support", "XIV on Mac", "wineprefix", "drive_c", "Program Files (x86)", "SquareEnix", "FINAL FANTASY XIV - A Realm Reborn"));
            }
        }

        return candidates;
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static IEnumerable<string> GetWindowsRegistryCandidates()
    {
        string[] globalRegistryPaths =
        [
            @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\SquareEnix\Final Fantasy XIV - A Realm Reborn",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\SquareEnix\Final Fantasy XIV - A Realm Reborn",
        ];

        foreach (var regPath in globalRegistryPaths)
        {
            var path = Registry.GetValue(regPath, "InstallLocation", null) as string;
            if (!string.IsNullOrEmpty(path))
                yield return path;
        }

        var cnIcon =
            Registry.GetValue(
                @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\最终幻想XIV",
                "DisplayIcon",
                null
            ) as string;
        if (!string.IsNullOrEmpty(cnIcon))
        {
            var installDir = Path.GetDirectoryName(cnIcon);
            if (!string.IsNullOrEmpty(installDir))
                yield return installDir;
        }
    }

    private static List<string> GetScanDriveRoots()
    {
        if (!OperatingSystem.IsWindows())
            return ["/"];

        var roots = new List<string>();
        try
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (!drive.IsReady)
                    continue;
                if (
                    drive.DriveType is DriveType.Fixed or DriveType.Removable or DriveType.Ram
                )
                {
                    roots.Add(drive.RootDirectory.FullName);
                }
            }
        }
        catch
        {
            // Fallback to system drive if drive enumeration fails
        }

        if (roots.Count == 0)
        {
            var sysDrive = Environment.GetEnvironmentVariable("SystemDrive");
            if (!string.IsNullOrEmpty(sysDrive))
                roots.Add(sysDrive + Path.DirectorySeparatorChar);
        }

        return roots;
    }

    private static List<string> ScanClientsBfs(
        List<string> roots,
        int maxDepth,
        CancellationToken ct
    )
    {
        var found = new ConcurrentBag<string>();
        var level = roots
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        try
        {
            for (int depth = 1; depth <= maxDepth && level.Count > 0 && !ct.IsCancellationRequested; depth++)
            {
                bool isLeafLevel = depth == maxDepth;
                var next = new ConcurrentBag<string>();
                var seen = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

                Parallel.ForEach(
                    level,
                    new ParallelOptions
                    {
                        MaxDegreeOfParallelism = 8,
                        CancellationToken = ct,
                    },
                    dir =>
                    {
                        if (ct.IsCancellationRequested)
                            return;

                        DirectoryInfo[] subDirs;
                        try
                        {
                            subDirs = new DirectoryInfo(dir).GetDirectories();
                        }
                        catch
                        {
                            return;
                        }

                        bool hasBoot = false;
                        bool hasGame = false;
                        foreach (var sub in subDirs)
                        {
                            if (sub.Name.Equals("boot", StringComparison.OrdinalIgnoreCase))
                                hasBoot = true;
                            else if (sub.Name.Equals("game", StringComparison.OrdinalIgnoreCase))
                                hasGame = true;
                            if (hasBoot && hasGame)
                                break;
                        }

                        if (hasBoot && hasGame && IsClientRoot(dir))
                        {
                            found.Add(dir);
                            return;
                        }

                        foreach (var sub in subDirs)
                        {
                            if (ct.IsCancellationRequested)
                                return;

                            if ((sub.Attributes & FileAttributes.ReparsePoint) != 0)
                                continue;

                            string lower = sub.Name.ToLowerInvariant();
                            if (lower.StartsWith('$') || lower.StartsWith('.') || SkipDirs.Contains(lower))
                                continue;

                            if (isLeafLevel)
                            {
                                if (IsClientRoot(sub.FullName))
                                    found.Add(sub.FullName);
                            }
                            else if (seen.TryAdd(sub.FullName, 0))
                            {
                                next.Add(sub.FullName);
                            }
                        }
                    }
                );

                level = [.. next];
            }
        }
        catch (OperationCanceledException)
        {
            // Return paths discovered before timeout
        }

        var list = found.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        list.Sort(StringComparer.OrdinalIgnoreCase);
        return list;
    }

    private static bool IsClientRoot(string dir)
    {
        if (!Directory.Exists(Path.Combine(dir, "boot")))
            return false;

        string gameDir = Path.Combine(dir, "game");
        if (!Directory.Exists(gameDir))
            return false;

        return Directory.Exists(Path.Combine(gameDir, "sqpack"))
            && (
                !OperatingSystem.IsWindows()
                || File.Exists(Path.Combine(gameDir, "ffxiv_dx11.exe"))
            );
    }

    private static ClientRegion? DetectClientRegion(string gameDir)
    {
        string sqpackDir = Path.Combine(gameDir, "game", "sqpack");
        if (!Directory.Exists(sqpackDir))
            return null;

        try
        {
            var gd = new GameData(sqpackDir, new LuminaOptions { CacheFileResources = false });
            if (gd.FileExists("exd/addon_0_chs.exd"))
                return ClientRegion.Chs;
            if (gd.FileExists("exd/addon_0_ko.exd"))
                return ClientRegion.Ko;
            if (gd.FileExists("exd/addon_0_tc.exd"))
                return ClientRegion.Tc;
            if (
                gd.FileExists("exd/addon_0_ja.exd")
                || gd.FileExists("exd/addon_0_en.exd")
                || gd.FileExists("exd/addon_0_de.exd")
                || gd.FileExists("exd/addon_0_fr.exd")
                || gd.FileExists("exd/addon_0_cht.exd")
            )
            {
                return ClientRegion.Global;
            }
        }
        catch
        {
            // Ignore unreadable or incomplete sqpack indexes
        }

        return null;
    }
}
