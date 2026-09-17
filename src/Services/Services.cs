using System.Diagnostics;
using System.Text.Json;
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
            {
                return Path.Combine(dir.FullName, "config.yml");
            }
            dir = dir.Parent;
        }

        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.yml");
    }

    public static AppConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath))
        {
            Console.WriteLine("错误: 未找到 config.yml。");
            Console.WriteLine(
                "请将 config.yml.example 复制并重命名为 config.yml，然后根据您的路径进行配置。"
            );
            Console.WriteLine("\n按任意键退出...");
            if (!Console.IsInputRedirected)
            {
                Console.ReadKey();
            }
            Environment.Exit(1);
        }

        try
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            return deserializer.Deserialize<AppConfig>(File.ReadAllText(ConfigPath))
                ?? new AppConfig();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"错误: 无法解析 config.yml: {ex.Message}");
            Console.WriteLine("\n按任意键退出...");
            if (!Console.IsInputRedirected)
            {
                Console.ReadKey();
            }
            Environment.Exit(1);
            return new AppConfig();
        }
    }

    public static void SaveConfig(AppConfig config)
    {
        try
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            File.WriteAllText(ConfigPath, serializer.Serialize(config), System.Text.Encoding.UTF8);
            Console.WriteLine("配置已成功保存至 config.yml");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"错误: 无法保存 config.yml: {ex.Message}");
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
        string version = Path.GetFileName(schemaDir);
        Console.WriteLine($"[Schema] 正在解析 YAML 定义，请稍候 ({version})...");
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
                catch (Exception ex)
                {
                    Console.WriteLine($"[Schema] 加载 {Path.GetFileName(file)} 失败: {ex.Message}");
                }
            }
        );

        return new Dictionary<string, ExdSchema>(schemas, StringComparer.OrdinalIgnoreCase);
    }
}
