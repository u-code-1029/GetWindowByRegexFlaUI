namespace GetWindowByRegexPattern.Services
{
    public sealed class AutomationSettings
    {
        public string ExecutablePath { get; set; } = string.Empty;
        public string ExecutableArguments { get; set; } = string.Empty;
        public string WindowTitlePattern { get; set; } = @"^Tool\s+Con(n)?trol.*";
        public string Backend { get; set; } = "Auto"; // Auto | UIA2 | UIA3
        public int WaitTimeoutMs { get; set; } = 10_000;
        public int PollIntervalMs { get; set; } = 200;

        public int SplashMaxWidth { get; set; } = 600;
        public int SplashMaxHeight { get; set; } = 400;
        public int SplashDurationMs { get; set; } = 5000;
    }
}