using XivExdUnpacker.src.Cli;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace XivExdUnpacker.src.Services;

public class ConfigService
{
    private static readonly string[] InternationalKeys = ["en", "ja", "de", "fr"];

    private sealed class LegacyAppConfig
    {
        public string? UiLanguage { get; set; }
        public string? OutputRoot { get; set; }
        public ClientConfig? Global { get; set; }
        public ClientConfig? Chs { get; set; }
        public ClientConfig? Ko { get; set; }
        public ClientConfig? Tc { get; set; }
        public Dictionary<string, ClientConfig?>? Clients { get; set; }
    }

    public static readonly string ConfigPath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "config.yml"
    );

    public static bool Exists => File.Exists(ConfigPath);

    public static bool TryLoadConfig(
        out AppConfig config,
        out string? errorMessage,
        out bool isNotFound
    )
    {
        if (!File.Exists(ConfigPath))
        {
            config = new AppConfig();
            errorMessage = Localizer.T(MsgKey.ConfigNotFound);
            isNotFound = true;
            return false;
        }

        try
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            var raw =
                deserializer.Deserialize<LegacyAppConfig>(File.ReadAllText(ConfigPath))
                ?? new LegacyAppConfig();

            bool needsMigration = raw.Clients != null;
            var cfg = MigrateConfig(raw);
            Localizer.SetLanguage(cfg.UiLanguage);

            if (needsMigration)
            {
                try
                {
                    SaveConfig(cfg);
                    ConsoleLogger.TopInfo(Localizer.T(MsgKey.ConfigMigrated));
                }
                catch (Exception saveEx)
                {
                    ConsoleLogger.TopError(Localizer.TF(MsgKey.ConfigSaveFailed, saveEx.Message));
                }
            }

            config = cfg;
            errorMessage = null;
            isNotFound = false;
            return true;
        }
        catch (Exception ex)
        {
            config = new AppConfig();
            errorMessage = Localizer.TF(MsgKey.ConfigParseFailed, ex.Message);
            isNotFound = false;
            return false;
        }
    }

    private static AppConfig MigrateConfig(LegacyAppConfig raw)
    {
        if (raw.Clients == null)
        {
            return new AppConfig
            {
                UiLanguage = raw.UiLanguage,
                OutputRoot = raw.OutputRoot,
                Global = EnsureAllFields(raw.Global),
                Chs = EnsureAllFields(raw.Chs),
                Ko = EnsureAllFields(raw.Ko),
                Tc = EnsureAllFields(raw.Tc),
            };
        }

        var legacyClients = new Dictionary<string, ClientConfig?>(
            raw.Clients,
            StringComparer.OrdinalIgnoreCase
        );

        var global = raw.Global;
        foreach (var intlKey in InternationalKeys)
        {
            if (!legacyClients.TryGetValue(intlKey, out var intlClient) || !HasAnySetting(intlClient))
                continue;

            global ??= new ClientConfig();
            global.Path ??= intlClient!.Path;
            global.SchemaVersion ??= intlClient!.SchemaVersion;
        }

        return new AppConfig
        {
            UiLanguage = raw.UiLanguage ?? "en",
            OutputRoot = raw.OutputRoot ?? "./rawexd",
            Global = EnsureAllFields(global),
            Chs = EnsureAllFields(PickRegionConfig(raw.Chs, legacyClients, "chs")),
            Ko = EnsureAllFields(PickRegionConfig(raw.Ko, legacyClients, "ko")),
            Tc = EnsureAllFields(PickRegionConfig(raw.Tc, legacyClients, "tc")),
        };
    }

    private static ClientConfig EnsureAllFields(ClientConfig? client)
    {
        return new ClientConfig
        {
            Path = client?.Path ?? "",
            OutputDir = client?.OutputDir ?? "",
            SchemaVersion = client?.SchemaVersion ?? "",
        };
    }

    private static ClientConfig? PickRegionConfig(
        ClientConfig? direct,
        Dictionary<string, ClientConfig?> legacyClients,
        string key
    )
    {
        if (HasAnySetting(direct))
            return direct;

        return legacyClients.TryGetValue(key, out var legacy) && HasAnySetting(legacy)
            ? legacy
            : null;
    }

    private static bool HasAnySetting(ClientConfig? client)
    {
        return client != null
            && (
                !string.IsNullOrWhiteSpace(client.Path)
                || !string.IsNullOrWhiteSpace(client.OutputDir)
                || !string.IsNullOrWhiteSpace(client.SchemaVersion)
            );
    }

    public static void SaveConfig(AppConfig config)
    {
        var serializer = new SerializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();

        bool isZh =
            !string.IsNullOrWhiteSpace(config.UiLanguage)
            && config.UiLanguage.Trim().ToLowerInvariant() is "chs" or "zh" or "zh-cn" or "cn";

        string Scalar(string? value)
        {
            var line = serializer.Serialize(new { v = value ?? "" }).TrimEnd('\r', '\n');
            return line.StartsWith("v: ", StringComparison.Ordinal) ? line[3..] : "\"\"";
        }

        var sb = new System.Text.StringBuilder();
        if (isZh)
        {
            sb.AppendLine("# XivExdUnpacker 配置文件（完整字段说明如下，留空表示使用默认行为或未启用）");
            sb.AppendLine();
            sb.AppendLine("# 界面语言：\"en\"（英文，默认）或 \"chs\"（简体中文）");
            sb.AppendLine($"uiLanguage: {Scalar(config.UiLanguage ?? "en")}");
            sb.AppendLine();
            sb.AppendLine("# 默认输出根目录；未单独指定 outputDir 时，各语言导出至 <outputRoot>/<语言>");
            sb.AppendLine($"outputRoot: {Scalar(config.OutputRoot ?? "./rawexd")}");
            sb.AppendLine();
            sb.AppendLine("# 国际服客户端（同时覆盖 en、ja、de、fr 四种语言）");
            AppendClientSection(sb, "global", config.Global, Scalar, isZh: true, isGlobal: true);
            sb.AppendLine();
            sb.AppendLine("# 国服客户端（chs）");
            AppendClientSection(sb, "chs", config.Chs, Scalar, isZh: true, isGlobal: false);
            sb.AppendLine();
            sb.AppendLine("# 韩服客户端（ko，未配置可留空）");
            AppendClientSection(sb, "ko", config.Ko, Scalar, isZh: true, isGlobal: false);
            sb.AppendLine();
            sb.AppendLine("# 繁中客户端（tc，未配置可留空）");
            AppendClientSection(sb, "tc", config.Tc, Scalar, isZh: true, isGlobal: false);
        }
        else
        {
            sb.AppendLine("# XivExdUnpacker configuration (leave optional fields empty for default behavior)");
            sb.AppendLine();
            sb.AppendLine("# UI language: \"en\" (default) or \"chs\" (Simplified Chinese)");
            sb.AppendLine($"uiLanguage: {Scalar(config.UiLanguage ?? "en")}");
            sb.AppendLine();
            sb.AppendLine("# Default output root directory; each language exports to <outputRoot>/<lang> unless outputDir is set");
            sb.AppendLine($"outputRoot: {Scalar(config.OutputRoot ?? "./rawexd")}");
            sb.AppendLine();
            sb.AppendLine("# Global client (covers en, ja, de, fr)");
            AppendClientSection(sb, "global", config.Global, Scalar, isZh: false, isGlobal: true);
            sb.AppendLine();
            sb.AppendLine("# Simplified Chinese client (chs)");
            AppendClientSection(sb, "chs", config.Chs, Scalar, isZh: false, isGlobal: false);
            sb.AppendLine();
            sb.AppendLine("# Korean client (ko, leave empty if unused)");
            AppendClientSection(sb, "ko", config.Ko, Scalar, isZh: false, isGlobal: false);
            sb.AppendLine();
            sb.AppendLine("# Traditional Chinese client (tc, leave empty if unused)");
            AppendClientSection(sb, "tc", config.Tc, Scalar, isZh: false, isGlobal: false);
        }

        File.WriteAllText(ConfigPath, sb.ToString(), System.Text.Encoding.UTF8);
    }

    private static void AppendClientSection(
        System.Text.StringBuilder sb,
        string sectionName,
        ClientConfig? client,
        Func<string?, string> scalar,
        bool isZh,
        bool isGlobal
    )
    {
        sb.AppendLine($"{sectionName}:");
        if (isZh)
        {
            sb.AppendLine("  # 游戏安装根目录（包含 boot 与 game 目录）；留空表示未启用");
            sb.AppendLine($"  path: {scalar(client?.Path)}");
            sb.AppendLine(
                isGlobal
                    ? "  # 可选输出根目录覆盖；设置后导出至 <outputDir>/<语言>，留空则使用 outputRoot"
                    : "  # 可选输出目录覆盖（直接导出到此目录）；留空则使用 <outputRoot>/<语言>"
            );
            sb.AppendLine($"  outputDir: {scalar(client?.OutputDir)}");
            sb.AppendLine("  # 可选 Schema 版本（如 \"latest\" 或具体版本号）；留空则自动读取 ffxivgame.ver，失败时回退 latest");
            sb.AppendLine($"  schemaVersion: {scalar(client?.SchemaVersion)}");
        }
        else
        {
            sb.AppendLine("  # Game installation root directory (contains boot and game); leave empty to disable");
            sb.AppendLine($"  path: {scalar(client?.Path)}");
            sb.AppendLine(
                isGlobal
                    ? "  # Optional output root override (exports to <outputDir>/<lang>); leave empty to use outputRoot"
                    : "  # Optional output directory override (exports directly here); leave empty to use <outputRoot>/<lang>"
            );
            sb.AppendLine($"  outputDir: {scalar(client?.OutputDir)}");
            sb.AppendLine("  # Optional schema version (e.g. \"latest\" or specific version); empty = auto-detect via ffxivgame.ver, fallback to latest");
            sb.AppendLine($"  schemaVersion: {scalar(client?.SchemaVersion)}");
        }
    }
}

public class SchemaService
{
    private static readonly IDeserializer YamlDeserializer = new DeserializerBuilder()
        .WithNamingConvention(CamelCaseNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .Build();

    public static Dictionary<string, ExdSchema> LoadSchemas(string schemaDir)
    {
        var schemas = new System.Collections.Concurrent.ConcurrentDictionary<string, ExdSchema>(
            StringComparer.OrdinalIgnoreCase
        );

        if (!Directory.Exists(schemaDir))
            return new Dictionary<string, ExdSchema>(StringComparer.OrdinalIgnoreCase);

        var files = Directory.GetFiles(schemaDir, "*.yml");

        Parallel.ForEach(
            files,
            new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
            file =>
            {
                try
                {
                    var content = File.ReadAllText(file);
                    var schema = YamlDeserializer.Deserialize<ExdSchema>(content);
                    if (schema?.Name != null)
                        schemas.TryAdd(schema.Name, schema);
                }
                catch
                {
                    // Skip malformed schema files silently; valid schemas still apply
                }
            }
        );

        return new Dictionary<string, ExdSchema>(schemas, StringComparer.OrdinalIgnoreCase);
    }
}
