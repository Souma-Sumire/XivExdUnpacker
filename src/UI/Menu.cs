namespace XivExdUnpacker.UI;

public static class Menu
{
    public static List<string>? ShowInteractiveMenu(List<(string Key, string SubText)> options)
    {
        int currentPos = 0;
        bool[] selected = new bool[options.Count];
        bool done = false;

        Console.WriteLine(
            "\n使用 [上下方向键] 移动, [空格] 勾选/取消, [A] 全选, [回车] 确定, [Esc] 退出\n"
        );

        // 预先输出空行占位，确保屏幕有足够空间，并处理潜在的滚动
        for (int i = 0; i < options.Count; i++)
            Console.WriteLine(new string(' ', Console.WindowWidth - 1));

        // 计算菜单起始行（从当前位置回溯）
        int startLine = Console.CursorTop - options.Count;
        if (startLine < 0)
            startLine = 0;

        Console.CursorVisible = false;

        while (!done)
        {
            for (int i = 0; i < options.Count; i++)
            {
                int targetLine = startLine + i;
                try
                {
                    Console.SetCursorPosition(0, targetLine);
                }
                catch
                {
                    continue;
                }

                string cursor = (i == currentPos) ? ">" : " ";
                string checkbox = selected[i] ? "[*]" : "[ ]";
                var (key, subText) = options[i];

                if (i == currentPos)
                {
                    Console.BackgroundColor = ConsoleColor.DarkGray;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                }

                string content = $"{cursor} {checkbox} {key.PadRight(4)} | {subText}";
                Console.Write(content.PadRight(Console.WindowWidth - 1));
                Console.ResetColor();
            }

            var inputKey = Console.ReadKey(true);
            switch (inputKey.Key)
            {
                case ConsoleKey.UpArrow:
                    currentPos = (currentPos - 1 + options.Count) % options.Count;
                    break;
                case ConsoleKey.DownArrow:
                    currentPos = (currentPos + 1) % options.Count;
                    break;
                case ConsoleKey.Spacebar:
                    selected[currentPos] = !selected[currentPos];
                    break;
                case ConsoleKey.A:
                    bool allSelected = selected.All(x => x);
                    for (int j = 0; j < selected.Length; j++)
                        selected[j] = !allSelected;
                    break;
                case ConsoleKey.Enter:
                    done = true;
                    break;
                case ConsoleKey.Escape:
                    return null;
            }
        }
        Console.WriteLine();
        return options.Where((t, i) => selected[i]).Select(x => x.Key).ToList();
    }
}
