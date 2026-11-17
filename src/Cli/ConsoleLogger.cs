namespace XivExdUnpacker.src.Cli;

public static class ConsoleLogger
{
    private static readonly object _lock = new();
    private static int _lastProgressLength;
    private static long _lastProgressTicks;
    private static readonly long _progressIntervalTicks = TimeSpan.FromMilliseconds(80).Ticks;

    public static void TagInfo(string tag, string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            WriteTag(tag, ConsoleColor.Cyan);
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }

    public static void TagSuccess(string tag, string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            WriteTag(tag, ConsoleColor.Cyan);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("✓ ");
            Console.ResetColor();
            Console.WriteLine(message);
        }
    }

    public static void TagWarn(string tag, string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            WriteTag(tag, ConsoleColor.Cyan);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("! ");
            Console.ResetColor();
            Console.WriteLine(message);
        }
    }

    public static void TagError(string tag, string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            WriteTag(tag, ConsoleColor.Cyan);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("✗ ");
            Console.ResetColor();
            Console.WriteLine(message);
        }
    }

    public static void Detail(string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"  - {message}");
            Console.ResetColor();
        }
    }

    public static void TopInfo(string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            Console.WriteLine(message);
        }
    }

    public static void TopError(string message)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }

    public static void Progress(string tag, string message)
    {
        if (Console.IsOutputRedirected)
            return;

        long now = Environment.TickCount64 * TimeSpan.TicksPerMillisecond;
        lock (_lock)
        {
            if (now - _lastProgressTicks < _progressIntervalTicks && _lastProgressLength > 0)
                return;

            _lastProgressTicks = now;
            var tagPrefix = FormatTag(tag);
            var line = $"{tagPrefix} {message}";
            int pad = Math.Max(0, _lastProgressLength - GetVisualWidth(line));
            Console.Write("\r");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write($"{tagPrefix} ");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write($"{message}{new string(' ', pad)}");
            Console.ResetColor();
            _lastProgressLength = GetVisualWidth(line);
        }
    }

    public static void ClearProgress()
    {
        lock (_lock)
            ClearProgressUnsafe();
    }

    public static void WithLock(Action action)
    {
        lock (_lock)
        {
            ClearProgressUnsafe();
            action();
        }
    }

    private static string FormatTag(string tag)
    {
        var bracketed = $"[{tag}]";
        return bracketed.Length < 5 ? bracketed.PadRight(5) : bracketed;
    }

    private static void WriteTag(string tag, ConsoleColor color)
    {
        Console.ForegroundColor = color;
        Console.Write($"{FormatTag(tag)} ");
        Console.ResetColor();
    }

    private static void ClearProgressUnsafe()
    {
        if (_lastProgressLength <= 0 || Console.IsOutputRedirected)
            return;

        Console.Write($"\r{new string(' ', _lastProgressLength)}\r");
        _lastProgressLength = 0;
        _lastProgressTicks = 0;
    }

    public static int GetVisualWidth(string s)
    {
        if (string.IsNullOrEmpty(s))
            return 0;

        int width = 0;
        foreach (char c in s)
            width += c > 127 ? 2 : 1;
        return width;
    }
}
