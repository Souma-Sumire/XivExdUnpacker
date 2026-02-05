using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;

namespace XivExdUnpacker.Services;

public static class SchemaUpdater
{
    private static readonly HttpClient _httpClient;
    private const string LATEST_ZIP_URL =
        "https://github.com/xivdev/EXDSchema/archive/refs/heads/latest.zip";
    private const string LATEST_COMMIT_API =
        "https://api.github.com/repos/xivdev/EXDSchema/commits/latest";

    static SchemaUpdater()
    {
        _httpClient = new HttpClient();
        // GitHub API requires a User-Agent header
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("XivExdUnpacker/1.0");
    }

    public static string? TryGetLocalCachedSchema(string version, string cacheDir)
    {
        var path = Path.Combine(cacheDir, version);
        if (Directory.Exists(path))
            return path;

        return null;
    }

    private static async Task<string?> GetLatestCommitHash(Action<string> log)
    {
        try
        {
            var response = await _httpClient.GetAsync(LATEST_COMMIT_API);
            if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                log("检测到 GitHub API 速率限制 (Rate Limit)。将尝试使用本地缓存。");
                return null;
            }
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("sha").GetString();
        }
        catch (Exception ex)
        {
            log($"检查 Github 更新失败: {ex.Message}");
            return null;
        }
    }

    public static async Task<string?> DownloadAndExtractSchema(
        string version,
        Action<string> log,
        string outputRoot
    )
    {
        try
        {
            var targetUrl = "";
            var targetDirName = "";
            string? remoteHash = null;

            if (version.Equals("latest", StringComparison.OrdinalIgnoreCase))
            {
                targetUrl = LATEST_ZIP_URL;
                targetDirName = "latest";

                remoteHash = await GetLatestCommitHash(log);
                var latestCachePath = Path.Combine(outputRoot, "latest");
                var versionFile = Path.Combine(outputRoot, "latest.version");

                if (!string.IsNullOrEmpty(remoteHash) && File.Exists(versionFile))
                {
                    var localHash = (await File.ReadAllTextAsync(versionFile)).Trim();
                    bool isHashMatch = string.Equals(
                        localHash,
                        remoteHash.Trim(),
                        StringComparison.OrdinalIgnoreCase
                    );
                    bool hasContent =
                        Directory.Exists(latestCachePath)
                        && Directory.GetFiles(latestCachePath, "*.yml").Length > 0;

                    if (isHashMatch && hasContent)
                    {
                        log("本地 'latest' 已是最新且内容完整，跳过下载。");
                        return latestCachePath;
                    }
                }
            }
            else
            {
                // Try version branch
                targetUrl =
                    $"https://github.com/xivdev/EXDSchema/archive/refs/heads/ver/{version}.zip";
                targetDirName = version;
            }

            var extractPath = Path.Combine(outputRoot, targetDirName);
            var tempExtractPath = extractPath + ".tmp";

            log($"正在下载定义: {targetUrl}");

            // Download
            var zipPath = Path.Combine(outputRoot, $"{targetDirName}.zip");
            var zipBytes = await _httpClient.GetByteArrayAsync(targetUrl);
            await File.WriteAllBytesAsync(zipPath, zipBytes);

            // Clean previous temp if exists
            if (Directory.Exists(tempExtractPath))
                Directory.Delete(tempExtractPath, true);

            Directory.CreateDirectory(tempExtractPath);

            log($"下载完成，正在解压...");

            // Extract to TEMP
            ZipFile.ExtractToDirectory(zipPath, tempExtractPath);
            File.Delete(zipPath); // Cleanup zip early

            // Flatten the structure in TEMP
            var extractedDirs = Directory.GetDirectories(tempExtractPath);
            if (extractedDirs.Length == 1)
            {
                var innerDir = extractedDirs[0];
                foreach (var file in Directory.GetFiles(innerDir))
                {
                    var dest = Path.Combine(tempExtractPath, Path.GetFileName(file));
                    if (File.Exists(dest))
                        File.Delete(dest);
                    File.Move(file, dest);
                }
                foreach (var dir in Directory.GetDirectories(innerDir))
                {
                    var dest = Path.Combine(tempExtractPath, Path.GetFileName(dir));
                    if (Directory.Exists(dest))
                        Directory.Delete(dest, true);
                    Directory.Move(dir, dest);
                }
                Directory.Delete(innerDir, true);
            }

            Directory.Move(tempExtractPath, extractPath);

            // Save Commit Hash in root schemas dir if available
            if (
                !string.IsNullOrEmpty(remoteHash)
                && version.Equals("latest", StringComparison.OrdinalIgnoreCase)
            )
            {
                await File.WriteAllTextAsync(
                    Path.Combine(outputRoot, "latest.version"),
                    remoteHash
                );
            }

            return extractPath;
        }
        catch (Exception ex)
        {
            log($"在线更新定义失败: {ex.Message}");
            // Cleanup partial directory if exists
            var extractPath = Path.Combine(
                outputRoot,
                version.Equals("latest", StringComparison.OrdinalIgnoreCase) ? "latest" : version
            );
            var tempExtractPath = extractPath + ".tmp";
            try
            {
                if (Directory.Exists(tempExtractPath))
                    Directory.Delete(tempExtractPath, true);
            }
            catch { }
            return null;
        }
    }
}
