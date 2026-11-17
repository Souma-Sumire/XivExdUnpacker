using XivExdUnpacker.src.Cli;
using XivExdUnpacker.src.I18n;
using XivExdUnpacker.src.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace XivExdUnpacker.src.Services;

public class ConfigService
{
    public static readonly string ConfigPath = FindConfigPath();

    private static string FindConfigPath()
    {
        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "XivExdUnpacker.csproj")))
                return Path.Combine(dir.FullName, "config.yml");
            dir = dir.Parent;
        }

        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.yml");
    }

    public static AppConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath))
        {
            ConsoleLogger.TopError(Localizer.T(MsgKey.ConfigNotFound));
            ConsoleLogger.TopInfo(Localizer.T(MsgKey.ConfigNotFoundHint));
            Environment.Exit(1);
        }

        try
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            var cfg =
                deserializer.Deserialize<AppConfig>(File.ReadAllText(ConfigPath))
                ?? new AppConfig();
            Localizer.SetLanguage(cfg.UiLanguage);
            return cfg;
        }
        catch (Exception ex)
        {
            ConsoleLogger.TopError(Localizer.TF(MsgKey.ConfigParseFailed, ex.Message));
            Environment.Exit(1);
            return new AppConfig();
        }
    }

    public static void SaveConfig(AppConfig config)
    {
        var serializer = new SerializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();
        File.WriteAllText(ConfigPath, serializer.Serialize(config), System.Text.Encoding.UTF8);
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
