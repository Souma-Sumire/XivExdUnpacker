using System.Text.Json;
using XivExdUnpacker.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace XivExdUnpacker.Services;

public class ConfigService
{
    public AppConfig LoadConfig()
    {
        var configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.yml");
        if (!File.Exists(configPath))
            configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.yml");

        if (!File.Exists(configPath))
        {
            Console.WriteLine("错误: 未找到 config.yml。");
            Console.WriteLine(
                "请将 config.yml.example 复制并重命名为 config.yml，然后根据您的路径进行配置。"
            );
            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
            Environment.Exit(1);
        }

        try
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            return deserializer.Deserialize<AppConfig>(File.ReadAllText(configPath))
                ?? new AppConfig();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"错误: 无法解析 config.yml: {ex.Message}");
            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
            Environment.Exit(1);
            return new AppConfig();
        }
    }

    public void SaveConfig(AppConfig config)
    {
        var configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.yml");
        try
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            File.WriteAllText(configPath, serializer.Serialize(config), System.Text.Encoding.UTF8);
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

    public Dictionary<string, ExdSchema> LoadSchemas(string schemaDir)
    {
        string version = Path.GetFileName(schemaDir);
        string cacheDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".cache");
        string cacheFile = Path.Combine(cacheDir, $"schema_{version}.json");

        if (File.Exists(cacheFile))
        {
            try
            {
                using var fs = File.OpenRead(cacheFile);
                var cached = JsonSerializer.Deserialize<Dictionary<string, ExdSchema>>(fs);
                if (cached != null)
                {
                    Console.WriteLine($"[Schema] 已从缓存加载定义 ({version})");
                    return cached;
                }
            }
            catch { }
        }

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
                catch { }
            }
        );

        var result = new Dictionary<string, ExdSchema>(schemas, StringComparer.OrdinalIgnoreCase);

        if (result.Count > 0)
        {
            try
            {
                if (!Directory.Exists(cacheDir))
                    Directory.CreateDirectory(cacheDir);
                string tempFile = cacheFile + ".tmp";
                using (var fs = File.Create(tempFile))
                {
                    JsonSerializer.Serialize(fs, result);
                }
                if (File.Exists(cacheFile))
                    File.Delete(cacheFile);
                File.Move(tempFile, cacheFile);
            }
            catch { }
        }

        return result;
    }
}
