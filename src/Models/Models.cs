using Lumina.Data;

namespace XivExdUnpacker.Models;

public class AppConfig
{
    public string? GlobalGamePath { get; set; }
    public string? GlobalSchemaVersion { get; set; }
    public List<string>? GlobalExclude { get; set; }

    public int? MaxSheetParallelism { get; set; }

    public ClientConfig? En { get; set; }
    public ClientConfig? Ja { get; set; }
    public ClientConfig? De { get; set; }
    public ClientConfig? Fr { get; set; }
    public ClientConfig? Cn { get; set; }
    public ClientConfig? Ko { get; set; }
    public ClientConfig? Tc { get; set; }

    public Dictionary<string, ClientConfig> GetClients()
    {
        var dict = new Dictionary<string, ClientConfig>(StringComparer.OrdinalIgnoreCase);
        if (En != null)
            dict["en"] = En;
        if (Ja != null)
            dict["ja"] = Ja;
        if (De != null)
            dict["de"] = De;
        if (Fr != null)
            dict["fr"] = Fr;
        if (Cn != null)
            dict["cn"] = Cn;
        if (Ko != null)
            dict["ko"] = Ko;
        if (Tc != null)
            dict["tc"] = Tc;
        return dict;
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
