namespace XivExdUnpacker.UI;

public static class Menu
{
    public static List<string>? ShowInteractiveMenu(List<(string Key, string Path)> options)
    {
        if (options.Count == 0)
            return new List<string>();

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("=== 请选择要导出的客户端 ===");
        Console.ResetColor();

        for (int i = 0; i < options.Count; i++)
        {
            var (key, path) = options[i];
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write($"  [{i + 1}] ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write($"{key.PadRight(4)}");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($" | {path}");
        }

        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n输入方式: 125 (连写) | 1-4 (范围) | en ja (代号) | all (全部)");
        Console.ResetColor();
        Console.Write("> ");

        string? input = Console.ReadLine()?.Trim().ToLower();
        if (string.IsNullOrEmpty(input))
            return null;

        // 全选快捷逻辑
        if (input == "all" || input == "*")
            return options.Select(o => o.Key).ToList();

        var selectedKeys = new List<string>();

        // 统一使用分词解析，不仅支持 1-4，也兼容 125 连写
        // 分隔符：空格, 逗号, 分号
        var parts = input.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

        // 如果只有一个词且全是数字，尝试按位拆分 (支持 125)
        if (parts.Length == 1 && parts[0].All(char.IsDigit) && !parts[0].Contains("-"))
        {
            foreach (char c in parts[0])
            {
                int index = c - '0';
                if (index >= 1 && index <= options.Count)
                    selectedKeys.Add(options[index - 1].Key);
            }
        }
        else
        {
            foreach (var part in parts)
            {
                // 处理范围: 1-4
                if (part.Contains("-"))
                {
                    var rangeParts = part.Split('-', StringSplitOptions.RemoveEmptyEntries);
                    if (
                        rangeParts.Length == 2
                        && int.TryParse(rangeParts[0], out int start)
                        && int.TryParse(rangeParts[1], out int end)
                    )
                    {
                        int min = Math.Min(start, end);
                        int max = Math.Max(start, end);
                        for (int i = min; i <= max; i++)
                        {
                            if (i >= 1 && i <= options.Count)
                                selectedKeys.Add(options[i - 1].Key);
                        }
                    }
                }
                // 处理单编号
                else if (int.TryParse(part, out int index))
                {
                    if (index >= 1 && index <= options.Count)
                        selectedKeys.Add(options[index - 1].Key);
                }
                // 处理代号 (en, ja...)
                else
                {
                    var match = options.FirstOrDefault(o =>
                        o.Key.Equals(part, StringComparison.OrdinalIgnoreCase)
                    );
                    if (match.Key != null)
                        selectedKeys.Add(match.Key);
                }
            }
        }

        if (selectedKeys.Count == 0)
            return null;
        return selectedKeys.Distinct().ToList();
    }
}
