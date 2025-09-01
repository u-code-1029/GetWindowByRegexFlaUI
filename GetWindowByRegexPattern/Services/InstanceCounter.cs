using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA2;
using FlaUI.UIA3;
using Serilog;
using System.Text.RegularExpressions;
using Window = FlaUI.Core.AutomationElements.Window;

namespace GetWindowByRegexPattern.Services
{
    public class InstanceCounter
    {
        public static int CountInstancesByTitleRegex(Regex titleRegex, ILogger? log = null, bool distinctByProcess = true)
        {
            using var automation = Environment.OSVersion.Version >= new Version(10, 0)
                ? (AutomationBase)new UIA3Automation()
                : new UIA2Automation();

            log?.Debug("Counting instances with backend: {Backend}", automation.GetType().Name);

            var desktop = automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                                 .Select(e => e.AsWindow())
                                 .Where(IsTopLevelCandidate)
                                 .Where(w => titleRegex.IsMatch(w.Title ?? ""))
                                 .ToList();

            log?.Information("Matched {Count} window(s) by title regex.", windows.Count);

            if (!distinctByProcess)
                return windows.Count;

            var distinctProcessIds = windows
                .Select(w => w.Properties.ProcessId.TryGetValue(out var pid) ? pid : -1)
                .Where(pid => pid > 0)
                .Distinct()
                .Count();

            log?.Information("Found {Count} distinct process instance(s).", distinctProcessIds);
            return distinctProcessIds;
        }

        private static bool IsTopLevelCandidate(Window w)
        {
            try
            {
                if (w.Properties.IsOffscreen.ValueOrDefault) return false;
                if (!w.IsEnabled) return false;

                var wp = w.Patterns.Window.PatternOrDefault;
                if (wp != null && wp.WindowVisualState.ValueOrDefault == WindowVisualState.Minimized)
                    return false;

                var r = w.BoundingRectangle;
                if (r.Width < 100 || r.Height < 50) return false;

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
