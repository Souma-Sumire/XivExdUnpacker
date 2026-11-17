using XivExdUnpacker.src.I18n;

namespace XivExdUnpacker.src.Cli;

public enum ExportStatus
{
    Completed,
    Skipped,
    Failed,
}

public record ClientExportResult
{
    public string ClientKey { get; init; } = "";
    public string LanguageName { get; init; } = "";
    public string ClientVersion { get; init; } = "-";
    public string ActualSchema { get; init; } = "-";
    public string OutputDir { get; init; } = "";
    public int SuccessCount { get; init; }
    public int FailedCount { get; init; }
    public double ElapsedSeconds { get; init; }
    public ExportStatus Status { get; init; } = ExportStatus.Completed;
    public string? StatusReason { get; init; }
}

public static class ConsoleReport
{
    public static void PrintSummary(List<ClientExportResult> orderedResults, TimeSpan totalElapsed)
    {
        if (orderedResults.Count == 0)
            return;

        ConsoleLogger.WithLock(() =>
        {
            int totalFailed = orderedResults.Sum(r => r.FailedCount);
            bool hasIssues =
                totalFailed > 0 || orderedResults.Any(r => r.Status != ExportStatus.Completed);

            Console.WriteLine();
            Console.ForegroundColor = hasIssues ? ConsoleColor.Yellow : ConsoleColor.Green;
            Console.WriteLine(
                totalFailed > 0
                    ? Localizer.TF(
                        MsgKey.ReportSummary,
                        orderedResults.Count,
                        totalFailed,
                        totalElapsed.TotalSeconds
                    )
                    : Localizer.TF(
                        MsgKey.ReportSummaryAllOk,
                        orderedResults.Count,
                        totalElapsed.TotalSeconds
                    )
            );
            Console.ResetColor();
        });
    }
}
