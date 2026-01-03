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
            File.WriteAllText(configPath, serializer.Serialize(config));
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
    public Dictionary<string, ExdSchema> LoadSchemas(string schemaDir)
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
                    var deserializer = new DeserializerBuilder()
                        .WithNamingConvention(CamelCaseNamingConvention.Instance)
                        .IgnoreUnmatchedProperties()
                        .Build();
                    var schema = deserializer.Deserialize<ExdSchema>(File.ReadAllText(file));
                    if (schema?.Name != null)
                        schemas.TryAdd(schema.Name, schema);
                }
                catch { }
            }
        );

        return new Dictionary<string, ExdSchema>(schemas, StringComparer.OrdinalIgnoreCase);
    }
}
