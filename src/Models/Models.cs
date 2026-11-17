namespace XivExdUnpacker.src.Models;

public class AppConfig
{
    private static readonly string[] InternationalKeys = ["en", "ja", "de", "fr"];

    public string? UiLanguage { get; set; }
    public string? OutputRoot { get; set; }
    public ClientConfig? Global { get; set; }
    public ClientConfig? Chs { get; set; }
    public ClientConfig? Ko { get; set; }
    public ClientConfig? Tc { get; set; }

    public Dictionary<string, ClientConfig> GetClients()
    {
        var dict = new Dictionary<string, ClientConfig>(StringComparer.OrdinalIgnoreCase);

        if (IsConfigured(Global))
        {
            foreach (var key in InternationalKeys)
                dict[key] = Global!;
        }

        if (IsConfigured(Chs))
            dict["chs"] = Chs!;
        if (IsConfigured(Ko))
            dict["ko"] = Ko!;
        if (IsConfigured(Tc))
            dict["tc"] = Tc!;

        return dict;
    }

    private static bool IsConfigured(ClientConfig? client)
    {
        return client != null
            && (
                !string.IsNullOrWhiteSpace(client.Path)
                || !string.IsNullOrWhiteSpace(client.OutputDir)
                || !string.IsNullOrWhiteSpace(client.SchemaVersion)
            );
    }

    public ClientConfig GetOrCreateRegionClient(string clientKey)
    {
        return clientKey.ToLowerInvariant() switch
        {
            "chs" => Chs ??= new ClientConfig(),
            "ko" => Ko ??= new ClientConfig(),
            "tc" => Tc ??= new ClientConfig(),
            _ => Global ??= new ClientConfig(),
        };
    }
}

public class ClientConfig
{
    public string? Path { get; set; }
    public string? OutputDir { get; set; }
    public string? SchemaVersion { get; set; }
}

public class ExdSchema
{
    public string? Name { get; set; }
    public List<SchemaField>? Fields { get; set; }
}

public class SchemaField
{
    public string? Name { get; set; }
    public string? Type { get; set; }
    public int Count { get; set; }
    public List<SchemaField>? Fields { get; set; }
}

public class ColumnInfo
{
    public Lumina.Data.Structs.Excel.ExcelColumnDefinition Definition { get; set; }
    public int OriginalIndex { get; set; }
    public string Name { get; set; } = "";
    public string Type { get; set; } = "";
    public bool IsUnknown { get; set; }
}
