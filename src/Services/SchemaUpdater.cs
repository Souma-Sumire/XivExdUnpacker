using System.Collections.Concurrent;
using System.IO.Compression;
using System.Net;
using System.Text;

namespace XivExdUnpacker.src.Services;

public record SchemaFetchResult(
    string? DirectoryPath,
    bool FromCache = false,
    bool IsUpToDate = false,
    bool IsNotFound = false,
    string? ErrorMessage = null
)
{
    public bool Success => !string.IsNullOrEmpty(DirectoryPath) && Directory.Exists(DirectoryPath);
}

public static class SchemaUpdater
{
    private static readonly HttpClient _httpClient;
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> _downloadLocks = new(
        StringComparer.OrdinalIgnoreCase
    );
    private static readonly SemaphoreSlim _refsLock = new(1, 1);
    private static Dictionary<string, string>? _cachedRemoteRefs;

    private const string LATEST_ZIP_URL =
        "https://github.com/xivdev/EXDSchema/archive/refs/heads/latest.zip";
    private const string GIT_INFO_REFS_URL =
        "https://github.com/xivdev/EXDSchema.git/info/refs?service=git-upload-pack";

    static SchemaUpdater()
    {
        _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("XivExdUnpacker/1.0");
    }

    private static async Task<(Dictionary<string, string>? refs, string? error)> GetRemoteRefsAsync()
    {
        if (_cachedRemoteRefs != null)
            return (_cachedRemoteRefs, null);

        await _refsLock.WaitAsync();
        try
        {
            if (_cachedRemoteRefs != null)
                return (_cachedRemoteRefs, null);

            var response = await _httpClient.GetAsync(GIT_INFO_REFS_URL);
            response.EnsureSuccessStatusCode();
            var bytes = await response.Content.ReadAsByteArrayAsync();
            var parsed = ParseGitUploadPackRefs(bytes);
            _cachedRemoteRefs = parsed;
            return (parsed, null);
        }
        catch (Exception ex)
        {
            return (null, ex.Message);
        }
        finally
        {
            _refsLock.Release();
        }
    }

    private static Dictionary<string, string> ParseGitUploadPackRefs(byte[] data)
    {
        var refs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        int pos = 0;
        while (pos + 4 <= data.Length)
        {
            var lenHex = Encoding.ASCII.GetString(data, pos, 4);
            if (!int.TryParse(lenHex, System.Globalization.NumberStyles.HexNumber, null, out int pktLen))
                break;

            pos += 4;
            if (pktLen == 0)
                continue;

            int payloadLen = pktLen - 4;
            if (payloadLen <= 0 || pos + payloadLen > data.Length)
                break;

            var line = Encoding.UTF8.GetString(data, pos, payloadLen);
            pos += payloadLen;

            if (line.StartsWith('#'))
                continue;

            int nulIdx = line.IndexOf('\0');
            if (nulIdx >= 0)
                line = line[..nulIdx];

            line = line.TrimEnd('\r', '\n');
            int spaceIdx = line.IndexOf(' ');
            if (spaceIdx > 0 && spaceIdx + 1 < line.Length)
            {
                var sha = line[..spaceIdx];
                var refName = line[(spaceIdx + 1)..];
                refs[refName] = sha;
            }
        }
        return refs;
    }

    public static async Task<SchemaFetchResult> DownloadAndExtractSchema(
        string version,
        string outputRoot
    )
    {
        bool isLatest = version.Equals("latest", StringComparison.OrdinalIgnoreCase);
        var targetDirName = isLatest ? "latest" : version;
        var lockKey = Path.Combine(outputRoot, targetDirName);
        var sem = _downloadLocks.GetOrAdd(lockKey, _ => new SemaphoreSlim(1, 1));
        await sem.WaitAsync();

        var zipPath = Path.Combine(outputRoot, $"{targetDirName}.zip");
        var extractPath = Path.Combine(outputRoot, targetDirName);
        var tempExtractPath = extractPath + ".tmp";

        try
        {
            if (!isLatest && Directory.Exists(extractPath) && Directory.GetFiles(extractPath, "*.yml").Length > 0)
                return new SchemaFetchResult(extractPath, FromCache: true);

            var (remoteRefs, refsError) = await GetRemoteRefsAsync();
            string? remoteHash = null;

            if (isLatest)
            {
                if (remoteRefs != null && remoteRefs.TryGetValue("refs/heads/latest", out var sha))
                    remoteHash = sha;

                var versionFile = Path.Combine(outputRoot, "latest.version");
                bool hasContent =
                    Directory.Exists(extractPath)
                    && Directory.GetFiles(extractPath, "*.yml").Length > 0;

                if (!string.IsNullOrEmpty(remoteHash) && File.Exists(versionFile))
                {
                    var localHash = (await File.ReadAllTextAsync(versionFile)).Trim();
                    if (
                        string.Equals(localHash, remoteHash, StringComparison.OrdinalIgnoreCase)
                        && hasContent
                    )
                    {
                        return new SchemaFetchResult(
                            extractPath,
                            FromCache: true,
                            IsUpToDate: true
                        );
                    }
                }

                if (remoteRefs == null)
                {
                    return new SchemaFetchResult(
                        null,
                        ErrorMessage: refsError ?? "Failed to query remote refs"
                    );
                }
            }
            else
            {
                var expectedRef = $"refs/heads/ver/{version}";
                if (remoteRefs != null && !remoteRefs.ContainsKey(expectedRef))
                {
                    return new SchemaFetchResult(null, IsNotFound: true, ErrorMessage: "Branch not found");
                }
            }

            string targetUrl = isLatest
                ? LATEST_ZIP_URL
                : $"https://github.com/xivdev/EXDSchema/archive/refs/heads/ver/{version}.zip";

            Directory.CreateDirectory(outputRoot);

            var response = await _httpClient.GetAsync(targetUrl);
            if (!response.IsSuccessStatusCode)
            {
                return new SchemaFetchResult(
                    null,
                    IsNotFound: response.StatusCode == HttpStatusCode.NotFound,
                    ErrorMessage: $"HTTP {(int)response.StatusCode} {response.StatusCode}"
                );
            }

            var zipBytes = await response.Content.ReadAsByteArrayAsync();
            await File.WriteAllBytesAsync(zipPath, zipBytes);

            if (Directory.Exists(tempExtractPath))
                Directory.Delete(tempExtractPath, true);

            Directory.CreateDirectory(tempExtractPath);
            ZipFile.ExtractToDirectory(zipPath, tempExtractPath);
            File.Delete(zipPath);

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

            if (Directory.Exists(extractPath))
                Directory.Delete(extractPath, true);
            Directory.Move(tempExtractPath, extractPath);

            if (!string.IsNullOrEmpty(remoteHash) && isLatest)
            {
                await File.WriteAllTextAsync(
                    Path.Combine(outputRoot, "latest.version"),
                    remoteHash
                );
            }

            return new SchemaFetchResult(extractPath);
        }
        catch (Exception ex)
        {
            try
            {
                if (File.Exists(zipPath))
                    File.Delete(zipPath);
                if (Directory.Exists(tempExtractPath))
                    Directory.Delete(tempExtractPath, true);
            }
            catch { }

            return new SchemaFetchResult(null, ErrorMessage: ex.Message);
        }
        finally
        {
            sem.Release();
        }
    }
}
