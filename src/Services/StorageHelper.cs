using System.IO.Enumeration;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace XivExdUnpacker.src.Services;

internal static class StorageHelper
{
    public static StringComparer PathComparer =>
        OperatingSystem.IsLinux() ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

    public static bool IsRotationalDisk(string path)
    {
        if (!OperatingSystem.IsWindows())
            return false;

        try
        {
            var root = Path.GetPathRoot(Path.GetFullPath(path));
            if (string.IsNullOrEmpty(root) || root.StartsWith(@"\\", StringComparison.Ordinal))
                return false;

            var driveLetter = root.TrimEnd('\\');
            if (driveLetter.Length != 2 || driveLetter[1] != ':')
                return false;

            using SafeFileHandle handle = CreateFileW(
                @"\\.\" + driveLetter,
                0,
                FileShare.ReadWrite | FileShare.Delete,
                IntPtr.Zero,
                3,
                0,
                IntPtr.Zero
            );

            if (handle.IsInvalid)
                return false;

            var query = new STORAGE_PROPERTY_QUERY
            {
                PropertyId = 7, // StorageDeviceSeekPenaltyProperty
                QueryType = 0,
            };

            var result = new DEVICE_SEEK_PENALTY_DESCRIPTOR();
            bool ok = DeviceIoControl(
                handle,
                0x002D1400, // IOCTL_STORAGE_QUERY_PROPERTY
                ref query,
                Marshal.SizeOf<STORAGE_PROPERTY_QUERY>(),
                ref result,
                Marshal.SizeOf<DEVICE_SEEK_PENALTY_DESCRIPTOR>(),
                out _,
                IntPtr.Zero
            );

            return ok && result.IncursSeekPenalty != 0;
        }
        catch
        {
            return false;
        }
    }

    public static Dictionary<string, long> EnumerateExistingCsvSizes(string outputDir)
    {
        var sizes = new Dictionary<string, long>(8192, PathComparer);
        if (!Directory.Exists(outputDir))
            return sizes;

        var enumOpts = new EnumerationOptions { RecurseSubdirectories = true };
        var fsEnum = new FileSystemEnumerable<(string path, long size)>(
            outputDir,
            (ref FileSystemEntry entry) => (entry.ToFullPath(), entry.Length),
            enumOpts
        )
        {
            ShouldIncludePredicate = (ref FileSystemEntry entry) =>
                !entry.IsDirectory && entry.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase),
        };

        foreach (var (fullPath, size) in fsEnum)
            sizes[fullPath] = size;

        return sizes;
    }

    public static void PreCreateSubdirectories(string outputDir, IReadOnlyList<string> sheetNames)
    {
        var created = new HashSet<string>(PathComparer);
        foreach (var sheetName in sheetNames)
        {
            int slash = sheetName.LastIndexOf('/');
            if (slash <= 0)
                continue;

            var subRel = sheetName[..slash].Replace('/', Path.DirectorySeparatorChar);
            if (created.Add(subRel))
                Directory.CreateDirectory(Path.Combine(outputDir, subRel));
        }
    }

    public static void WriteIfChanged(
        string targetPath,
        ReadOnlySpan<byte> payload,
        long? existingSize,
        byte[] readScratch
    )
    {
        if (existingSize.HasValue && existingSize.Value == payload.Length)
        {
            try
            {
                if (FileBytesEqual(targetPath, payload, readScratch))
                    return;
            }
            catch
            {
                // 读取对比失败时回退到直接覆盖写入以保证输出正确
            }
        }

        WriteFileDirect(targetPath, payload);
    }

    public static void WriteFileDirect(string targetPath, ReadOnlySpan<byte> payload)
    {
        // 小于 4KB 的文件在 NTFS 中常驻 MFT，跳过 preallocationSize 避免额外元数据调用
        long prealloc = payload.Length >= 4096 ? payload.Length : 0;
        try
        {
            WriteHandle(targetPath, payload, prealloc);
        }
        catch (DirectoryNotFoundException)
        {
            var dir = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            WriteHandle(targetPath, payload, prealloc);
        }
    }

    private static void WriteHandle(string targetPath, ReadOnlySpan<byte> payload, long prealloc)
    {
        using var handle = File.OpenHandle(
            targetPath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            FileOptions.SequentialScan,
            preallocationSize: prealloc
        );
        RandomAccess.Write(handle, payload, 0);
    }

    private static bool FileBytesEqual(string path, ReadOnlySpan<byte> expected, byte[] scratch)
    {
        using var fs = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            4096,
            FileOptions.SequentialScan
        );

        int offset = 0;
        while (offset < expected.Length)
        {
            int toRead = Math.Min(scratch.Length, expected.Length - offset);
            int read = fs.Read(scratch, 0, toRead);
            if (read == 0)
                return false;
            if (!scratch.AsSpan(0, read).SequenceEqual(expected.Slice(offset, read)))
                return false;
            offset += read;
        }

        return fs.ReadByte() == -1;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct STORAGE_PROPERTY_QUERY
    {
        public int PropertyId;
        public int QueryType;
        public int AdditionalParameters;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DEVICE_SEEK_PENALTY_DESCRIPTOR
    {
        public uint Version;
        public uint Size;
        public byte IncursSeekPenalty;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(
        string lpFileName,
        uint dwDesiredAccess,
        FileShare dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile
    );

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(
        SafeFileHandle hDevice,
        uint dwIoControlCode,
        ref STORAGE_PROPERTY_QUERY lpInBuffer,
        int nInBufferSize,
        ref DEVICE_SEEK_PENALTY_DESCRIPTOR lpOutBuffer,
        int nOutBufferSize,
        out uint lpBytesReturned,
        IntPtr lpOverlapped
    );
}
