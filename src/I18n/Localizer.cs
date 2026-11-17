using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using XivExdUnpacker.src.Models;
using XivExdUnpacker.src.Services;

namespace XivExdUnpacker.src.I18n;

public static class Localizer
{
    private static string _language = "en";

    public static string Language => _language;

    public static void SetLanguage(string? language)
    {
        if (string.IsNullOrWhiteSpace(language))
        {
            _language = "en";
            return;
        }

        var normalized = language.Trim().ToLowerInvariant();
        _language = normalized is "chs" or "zh" or "zh-cn" or "zh_cn" or "cn" ? "zh" : "en";
    }

    public static void TryLoadFromConfigIfExists()
    {
        try
        {
            if (!File.Exists(ConfigService.ConfigPath))
                return;

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            var cfg = deserializer.Deserialize<AppConfig>(File.ReadAllText(ConfigService.ConfigPath));
            if (cfg != null)
                SetLanguage(cfg.UiLanguage);
        }
        catch
        {
            // Ignore errors during early UI language detection
        }
    }

    public static string T(MsgKey key)
    {
        var map = _language == "zh" ? ZhMessages.Map : EnMessages.Map;
        if (map.TryGetValue(key, out var val))
            return val;
        return EnMessages.Map.TryGetValue(key, out var fallback) ? fallback : key.ToString();
    }

    public static string TF(MsgKey key, params object?[] args) => string.Format(T(key), args);
}
