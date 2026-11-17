namespace XivExdUnpacker.src.Models;

public class AppConfig
{
    public string? UiLanguage { get; set; }
    public string? OutputRoot { get; set; }
    public GlobalClientConfig? Global { get; set; }
    public Dictionary<string, ClientConfig?>? Clients { get; set; }

    public Dictionary<string, ClientConfig> GetClients()
    {
        var dict = new Dictionary<string, ClientConfig>(StringComparer.OrdinalIgnoreCase);
        if (Clients == null)
            return dict;

        foreach (var (key, value) in Clients)
        {
            if (string.IsNullOrWhiteSpace(key))
                continue;
            dict[key] = value ?? new ClientConfig();
        }
        return dict;
    }
}

public class GlobalClientConfig
{
    public string? Path { get; set; }
    public string? SchemaVersion { get; set; }
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
