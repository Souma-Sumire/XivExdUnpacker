using Microsoft.Win32;

namespace XivExdUnpacker.Services;

public class GamePathDetector
{
    public string? Detect(bool isInternational = true)
    {
        if (OperatingSystem.IsWindows())
            return DetectWindows(isInternational);
        if (OperatingSystem.IsLinux())
            return DetectLinux();
        if (OperatingSystem.IsMacOS())
            return DetectMacOS();
        return null;
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private string? DetectWindows(bool isInternational)
    {
        if (isInternational)
        {
            // 只检测国际服路径
            // 1. 尝试从注册表读取
            string[] registryPaths =
            {
                @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\SquareEnix\Final Fantasy XIV - A Realm Reborn",
                @"HKEY_LOCAL_MACHINE\SOFTWARE\SquareEnix\Final Fantasy XIV - A Realm Reborn",
            };
            foreach (var regPath in registryPaths)
            {
                var path = Registry.GetValue(regPath, "InstallLocation", null) as string;
                if (!string.IsNullOrEmpty(path))
                {
                    if (Directory.Exists(Path.Combine(path, "game", "sqpack")))
                        return path;
                }
            }

            // 2. 尝试默认路径
            var defaultPath =
                "C:/Program Files (x86)/SquareEnix/FINAL FANTASY XIV - A Realm Reborn";
            if (Directory.Exists(Path.Combine(defaultPath, "game", "sqpack")))
                return defaultPath;
        }
        else
        {
            // 只检测区域服路径 (国服/韩服/台服)
            // 国服路径 (盛趣/数龙)
            var cnPath =
                Registry.GetValue(
                    @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\最终幻想XIV",
                    "DisplayIcon",
                    null
                ) as string;
            if (!string.IsNullOrEmpty(cnPath))
            {
                var installDir = Path.GetDirectoryName(cnPath);
                if (
                    !string.IsNullOrEmpty(installDir)
                    && Directory.Exists(Path.Combine(installDir, "game", "sqpack"))
                )
                {
                    return installDir;
                }
            }

            // 国服默认路径
            var cnDefaultPath = "C:/Program Files (x86)/上海数龙科技有限公司/最终幻想XIV";
            if (Directory.Exists(Path.Combine(cnDefaultPath, "game", "sqpack")))
                return cnDefaultPath;
        }

        return null;
    }

    private string? DetectLinux()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (string.IsNullOrEmpty(home))
            return null;

        string[] paths =
        {
            Path.Combine(home, ".xlcore", "ffxiv"), // XIVLauncher Core
            Path.Combine(
                home,
                ".steam",
                "steam",
                "steamapps",
                "common",
                "FINAL FANTASY XIV Online"
            ), // Standard Steam
            Path.Combine(
                home,
                ".local",
                "share",
                "Steam",
                "steamapps",
                "common",
                "FINAL FANTASY XIV Online"
            ), // Steam Deck / Distro specific
        };

        foreach (var path in paths)
        {
            if (Directory.Exists(Path.Combine(path, "game", "sqpack")))
                return path;
        }
        return null;
    }

    private string? DetectMacOS()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        string[] paths =
        {
            "/Applications/FINAL FANTASY XIV ONLINE.app/Contents/SharedSupport/finalfantasyxiv", // Official Client
            Path.Combine(home, "Library", "Application Support", "XIV on Mac", "ffxiv"), // XIV on Mac
        };

        foreach (var path in paths)
        {
            if (Directory.Exists(Path.Combine(path, "game", "sqpack")))
                return path;
        }
        return null;
    }
}
