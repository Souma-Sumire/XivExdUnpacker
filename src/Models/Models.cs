using Lumina.Data;

namespace XivExdUnpacker.Models;

public class AppConfig
{
    // 全局设置
    public string? GlobalGamePath { get; set; }
    public string? GlobalSchemaVersion { get; set; }
    public bool? ClearOutputDir { get; set; }
    public List<string>? GlobalExclude { get; set; }

    // 性能设置
    public int? MaxClientParallelism { get; set; } // 同时处理的客户端数量 (默认: 客户端数量)
    public int? MaxSheetParallelism { get; set; } // 每个客户端内部导出表的并行数 (默认: CPU核心数)

    // 客户端配置 (扁平化)
    public ClientConfig? En { get; set; }
    public ClientConfig? Ja { get; set; }
    public ClientConfig? De { get; set; }
    public ClientConfig? Fr { get; set; }
    public ClientConfig? Cn { get; set; }
    public ClientConfig? Ko { get; set; }
    public ClientConfig? Tc { get; set; }

    // 辅助方法: 生成字典以便遍历
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
