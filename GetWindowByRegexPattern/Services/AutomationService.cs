using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA2;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Window = FlaUI.Core.AutomationElements.Window;

namespace GetWindowByRegexPattern.Services
{
    public class AutomationService
    {
        private readonly ILogger<AutomationService> _log;
        private readonly AutomationSettings _cfg;
        private AutomationBase? _automation;
        private readonly Regex _titleRegex;

        public AutomationService(IOptions<AutomationSettings> cfg, ILogger<AutomationService> log)
        {
            _cfg = cfg.Value;
            _log = log;
            _titleRegex = new Regex(_cfg.WindowTitlePattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
            _log.LogInformation("AutomationService initialized. Backend={Backend}, Pattern={Pattern}", _cfg.Backend, _cfg.WindowTitlePattern);
        }

        public void RunOnce()
        {
            ValidateConfig();

            var instances = InstanceCounter.CountInstancesByTitleRegex(_titleRegex);
            _log.LogInformation("Detected {Count} running instance(s) by title pattern.", instances);

            var app = EnsureApplicationRunning();

            try
            {
                using (_automation = CreateAutomation(_cfg.Backend))
                {
                    _log.LogInformation("Automation backend created: {BackendType}", _automation.GetType().Name);

                    var window = WaitForWindowByRegex(
                        timeout: TimeSpan.FromMilliseconds(_cfg.WaitTimeoutMs),
                        poll: TimeSpan.FromMilliseconds(_cfg.PollIntervalMs));

                    if (window == null)
                    {
                        _log.LogError("No window found matching pattern: {Pattern}", _cfg.WindowTitlePattern);
                        throw new InvalidOperationException("Target window not found.");
                    }

                    _log.LogInformation("Target window acquired. Title=\"{Title}\", Handle={Handle}, PID={Pid}",
                        window.Title,
                        window.Properties.NativeWindowHandle.ValueOrDefault,
                        window.Properties.ProcessId.ValueOrDefault);

                    window.Focus();
                    _log.LogInformation("Window focused.");
                    // Add your next automation steps here; keep logging key steps as needed.
                }
            }
            finally
            {
                _log.LogInformation("AutomationService finished RunOnce.");
            }
        }

        private void ValidateConfig()
        {
            if (string.IsNullOrWhiteSpace(_cfg.ExecutablePath))
            {
                _log.LogError("ExecutablePath is not configured.");
                throw new ArgumentException("ExecutablePath is required.");
            }

            if (!File.Exists(_cfg.ExecutablePath))
            {
                _log.LogError("Executable not found at path: {Path}", _cfg.ExecutablePath);
                throw new FileNotFoundException("Executable not found.", _cfg.ExecutablePath);
            }

            _log.LogInformation("Configuration validated. ExecutablePath={Path}", _cfg.ExecutablePath);
        }

        private AutomationBase CreateAutomation(string backend) =>
            backend?.ToUpperInvariant() switch
            {
                "UIA2" => LogCreate(new UIA2Automation()),
                "UIA3" => LogCreate(new UIA3Automation()),
                _ => Environment.OSVersion.Version >= new Version(10, 0)
                    ? LogCreate(new UIA3Automation())
                    : LogCreate(new UIA2Automation())
            };

        private AutomationBase LogCreate(AutomationBase a)
        {
            _log.LogDebug("Creating automation backend instance: {Type}", a.GetType().Name);
            return a;
        }

        private Window? WaitForWindowByRegex(TimeSpan timeout, TimeSpan poll)
        {
            var start = DateTime.UtcNow;
            _log.LogInformation("Begin waiting for window. Timeout={Timeout}ms, Poll={Poll}ms",
                timeout.TotalMilliseconds, poll.TotalMilliseconds);

            do
            {
                var candidates = FindTopLevelWindows()
                    .Select(e => e.AsWindow())
                    .Where(w => _titleRegex.IsMatch(w?.Title ?? string.Empty))
                    .ToList();

                if (candidates.Count > 0)
                {
                    _log.LogDebug("Found {Count} candidate window(s) by title pattern.", candidates.Count);
                }

                foreach (var win in candidates)
                {
                    if (!IsLikelySplash(win, start))
                    {
                        _log.LogInformation("Selected window: \"{Title}\" [{W}x{H}]",
                            win.Title, win.BoundingRectangle.Width, win.BoundingRectangle.Height);
                        return win;
                    }

                    _log.LogDebug("Skipping likely splash window: \"{Title}\" [{W}x{H}]",
                        win.Title, win.BoundingRectangle.Width, win.BoundingRectangle.Height);
                }

                System.Threading.Thread.Sleep(poll);
            }
            while (DateTime.UtcNow - start < timeout);

            _log.LogWarning("Timed out while waiting for target window.");
            return null;
        }

        private AutomationElement[] FindTopLevelWindows()
        {
            var desktop = _automation!.GetDesktop();
            var wins = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
            _log.LogDebug("Enumerated {Count} top-level window(s).", wins.Length);
            return wins;
        }

        private bool IsLikelySplash(Window w, DateTime startUtc)
        {
            try
            {
                var bounds = w.BoundingRectangle;
                var withinSplashWindow = (DateTime.UtcNow - startUtc).TotalMilliseconds < _cfg.SplashDurationMs;
                var smallEnough = bounds.Width <= _cfg.SplashMaxWidth && bounds.Height <= _cfg.SplashMaxHeight;

                if (smallEnough && withinSplashWindow)
                {
                    return true;
                }

                var wp = w.Patterns.Window.PatternOrDefault;
                if (wp != null && wp.WindowVisualState.ValueOrDefault == WindowVisualState.Minimized)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "Heuristic splash check failed; ignoring window.");
            }

            return false;
        }

        private FlaUI.Core.Application EnsureApplicationRunning()
        {
            var exePath = _cfg.ExecutablePath;
            var procName = Path.GetFileNameWithoutExtension(exePath);

            _log.LogInformation("Checking for existing process: {ProcName}", procName);

            var procs = Process.GetProcessesByName(procName)
                .Where(p =>
                {
                    try
                    {
                        return string.Equals(p.MainModule?.FileName, exePath, StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        return false;
                    }
                })
                .ToArray();

            if (procs.Length > 0)
            {
                _log.LogInformation("Attaching to existing process PID={Pid}", procs[0].Id);
                return FlaUI.Core.Application.Attach(procs[0]);
            }

            _log.LogInformation("No existing process found. Launching: {Path} {Args}", exePath, _cfg.ExecutableArguments);

            var app = string.IsNullOrWhiteSpace(_cfg.ExecutableArguments)
                ? FlaUI.Core.Application.Launch(exePath)
                : FlaUI.Core.Application.Launch(exePath, _cfg.ExecutableArguments);

            _log.LogInformation("Launched process. PID={Pid}", app.ProcessId);
            return app;
        }

        public void Dispose()
        {
            if (_automation != null)
            {
                _log.LogDebug("Disposing automation backend: {Type}", _automation.GetType().Name);
                _automation.Dispose();
                _automation = null;
            }
        }
    }
}
