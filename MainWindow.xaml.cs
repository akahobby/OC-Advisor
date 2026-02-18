using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Navigation;

namespace OcAdvisor;

public partial class MainWindow : Window
{
    private bool _uiReady;
    private List<Recommendation> _all = new();
    private List<Recommendation> _disclaimer = new();
    private ICollectionView? _view;
    private string _lastText = "";
    private string _lastJson = "";
    private SystemContext? _lastCtx;
    private RamCalcConfig _ramCfg = RamCalcConfigStore.Load();
    private CpuGpuCalcConfig _cpuGpuCfg = CpuGpuCalcConfigStore.Load();
    private Dictionary<string, bool> _groupExpanded = UiStateStore.Load();

    private OverclockAssumptions _assumptions = new OverclockAssumptions();


    public MainWindow()
    {
        InitializeComponent();

        // Defaults for assumptions (UI may not exist in older layouts; wrap in try)
        try
        {
            // 240mm AIO + Average silicon default
            if (CoolingTypeBox != null) CoolingTypeBox.SelectedIndex = 1;
            if (SiliconQualityBox != null) SiliconQualityBox.SelectedIndex = 1;
            UpdateAssumptionsUi();
        }
        catch { }


        // Version/build stamp
        try
        {
            var asm = Assembly.GetExecutingAssembly();
            var v = asm.GetName().Version;
            var ver = (v == null) ? "1.0" : $"{v.Major}.{v.Minor}";
            var build = File.GetLastWriteTime(asm.Location);
            TxtBuildStamp.Text = $"v{ver} • build {build:yyyy-MM-dd}";
        }
        catch { }
        // Don't let an exception during detection kill the app before the UI is visible.
        Loaded += async (_, __) =>
        {
            _uiReady = true;
            // Apply view once so toggles/search don't crash before data exists.
            ApplyView();
            await GenerateAsyncSafe();
        };}

    private async Task GenerateAsyncSafe()
    {
        try
        {
            await GenerateAsync();
        }
        catch (Exception ex)
        {
            // Keep the app alive and show a helpful error.
            SetStatus("Detection failed.");
            TxtSummary.Text = "Detection failed. Click Generate to retry.";
            MessageBox.Show(ex.ToString(), "OC Advisor - Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task GenerateAsync()
    {
        BtnGenerate.IsEnabled = false;
        try
        {

            var result = await Task.Run(() => OcAdvisorBackend.Generate(TuningProfile.Competitive));

            _lastCtx = result.ctx;
            _lastText = result.text;
            _lastJson = result.json;

            TxtSummary.Text =
                $"{result.ctx.OsCaption} ({result.ctx.OsVersion})\n" +
                (!string.IsNullOrWhiteSpace(result.ctx.ActivePowerPlan) ? $"Power plan: {result.ctx.ActivePowerPlan}\n" : "") +
                $"CPU: {result.ctx.CpuName}\n" +
                $"GPU: {result.ctx.PrimaryGpuName}\n" +
                $"RAM: {result.ctx.TotalRamGb:0.#} GB ({result.ctx.MemoryType})" +
                (result.ctx.EffectiveMemMts.HasValue ? $"\nRAM Rate: {result.ctx.EffectiveMemMts} MT/s" : "");

            
            var recs = (result.recs ?? new List<Recommendation>())
                .Where(r => !string.Equals(r.Section, "STORAGE", StringComparison.OrdinalIgnoreCase))
                .Select(r =>
                {
                    // Move the "TESTING" bucket into a dedicated Disclaimer panel instead of the main grid
                    if (string.Equals(r.Section, "TESTING", StringComparison.OrdinalIgnoreCase))
                        r.Section = "Disclaimer";
                    return r;
                })
                .ToList();

            _disclaimer = recs
                .Where(r => string.Equals(r.Section, "Disclaimer", StringComparison.OrdinalIgnoreCase))
                .ToList();

            _all = recs
                .Where(r => !string.Equals(r.Section, "Disclaimer", StringComparison.OrdinalIgnoreCase))
                .ToList();

            try
            {
                if (DisclaimerItems != null)
                    DisclaimerItems.ItemsSource = _disclaimer;
            }
            catch { }
// If CPU/GPU OC calculators are enabled, remove overlapping "basic" recommendations
            // so the advice lives in one place (mirrors RAM Calc behavior).
            if (_cpuGpuCfg.Enabled)
            {
                var gpuDup = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Fan curve",
                    "OC steps",
                    "Power / Core / Memory",
                    "Undervolt"
                };

                var cpuDup = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "CO",
                    "Max CPU Boost Clock Override",
                    "PBO Scalar",
                    "Thermal Throttle Limit"
                };

                _all = _all.Where(r =>
                {
                    var sec = r.Section ?? "";
                    var set = r.Setting ?? "";
                    if (string.Equals(sec, "GPU", StringComparison.OrdinalIgnoreCase) && gpuDup.Contains(set)) return false;
                    if (string.Equals(sec, "CPU", StringComparison.OrdinalIgnoreCase) && cpuDup.Contains(set)) return false;
                    return true;
                }).ToList();
            }

// Append RAM timing calculator output (if configured). Never break main recommendations if it fails.
            if (_ramCfg.Enabled && _lastCtx != null)
            {
                try { _all.AddRange(RamTimingCalculator.BuildRecommendations(_lastCtx, _ramCfg)); }
                catch { }
            }

            // Append CPU/GPU OC calculator output (if configured). Never break main recommendations if it fails.
            if (_cpuGpuCfg.Enabled && _lastCtx != null)
            {
                try
                {
                    var calcRows = CpuGpuOcCalculator.BuildRecommendations(_lastCtx, _cpuGpuCfg);
                    InsertCpuGpuCalcRowsInPlace(_all, calcRows);
                }
                catch { }
            }

            // Final: enforce a stable section order (mirrors RAM calc behavior).
            // This prevents GPU from falling below BIOS when the "basic" GPU rows are removed
            // and only calculator rows remain.
            try
            {
                var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                {
                    ["CPU"] = 0,
                    ["MEMORY"] = 1, // internal name used by backend
                    ["RAM"] = 1,    // safety if any rows already use RAM
                    ["GPU"] = 2,
                    ["BIOS"] = 3,
                    ["DISPLAY"] = 4,
                    ["Disclaimer"] = 5
                };

                // stamp current order so we keep row order inside each section
                var indexed = _all.Select((r, i) => (r, i)).ToList();
                _all = indexed
                    .OrderBy(x => order.TryGetValue(x.r.Section ?? string.Empty, out var p) ? p : 999)
                    .ThenBy(x => x.i)
                    .Select(x => x.r)
                    .ToList();
            }
            catch { }

            // Add assumption-based meta + compatibility checks (and keep it deduped)
            try
            {
                var a = _assumptions ?? new OverclockAssumptions();
                var assumedV = CompatibilityAnalyzer.AssumedCpuVcore(a);
                var risk = StabilityCalculator.DetermineRisk(assumedV);
                var score = StabilityCalculator.Calculate(a.Cooling, a.Silicon, assumedV);

                // Update header UI
                try { UpdateAssumptionsUi(); } catch { }

                _all.Add(new Recommendation
                {
                    Section = "SYSTEM",
                    Setting = "Assumptions",
                    Value = $"{(CoolingTypeBox?.Text ?? "240mm AIO")} • {(SiliconQualityBox?.Text ?? "Average")}",
                    Path = "OC Advisor",
                    Reason = "Used for confidence scoring and sanity checks only (does not auto-apply).",
                    WarningLevel = "Info"
                });

                _all.Add(new Recommendation
                {
                    Section = "SYSTEM",
                    Setting = "Risk level",
                    Value = risk,
                    Path = "OC Advisor",
                    Reason = "Based on assumed voltage/cooling headroom; still validate with stress tests.",
                    WarningLevel = risk == "Safe Daily" ? "Info" : (risk == "Aggressive Daily" ? "Warning" : "Danger")
                });

                _all.Add(new Recommendation
                {
                    Section = "SYSTEM",
                    Setting = "Stability confidence",
                    Value = $"{score}%",
                    Path = "OC Advisor",
                    Reason = "Heuristic score that helps you decide how conservative to be. Not a guarantee.",
                    WarningLevel = score >= 80 ? "Info" : (score >= 65 ? "Warning" : "Danger")
                });

                _all.Add(new Recommendation
                {
                    Section = "SYSTEM",
                    Setting = "Estimated uplift",
                    Value = PerformanceEstimator.Estimate(_lastCtx ?? result.ctx, a),
                    Path = "OC Advisor",
                    Reason = "Rough estimate. Real gains depend on silicon, cooling, and stability limits.",
                    WarningLevel = "Info"
                });

                // Compatibility analyzer warnings
                _all.AddRange(CompatibilityAnalyzer.Build(_lastCtx ?? result.ctx, a));

                // Add validation checklist once (in Disclaimer)
                _disclaimer.Add(new Recommendation
                {
                    Section = "Disclaimer",
                    Setting = "Required validation",
                    Value = "Cinebench R23 loop 30min • OCCT 1hr CPU+RAM • TM5 Extreme (memory)",
                    Path = "Testing",
                    Reason = "Validate one change at a time. If unstable: back off GPU memory → GPU core → CPU CO/boost → RAM/FCLK.",
                    WarningLevel = "Info"
                });
}
            catch { }

            // Final dedupe pass (prevents accidental duplicates across calculators/sections)
            _all = RecommendationDedupe.Dedupe(_all);

            // Keep disclaimer panel deduped too
            _disclaimer = RecommendationDedupe.Dedupe(_disclaimer ?? new List<Recommendation>());
            try { if (DisclaimerItems != null) DisclaimerItems.ItemsSource = _disclaimer; } catch { }


            // Rebuild text/json after UI-side modifications so copy/save includes everything.
            try
            {
                if (_lastCtx != null)
                {
                    _lastText = ReportBuilder.BuildText(_lastCtx, _all);
                    _lastJson = ReportBuilder.BuildJson(_lastCtx, _all);
                }
            }
            catch { }

            ApplyView();

        }
        finally
        {
            BtnGenerate.IsEnabled = true;
        }
    }

    private async void BtnGenerate_Click(object sender, RoutedEventArgs e) => await GenerateAsync();

    private void BtnRamCalc_Click(object sender, RoutedEventArgs e)
    {
        // Seed calculator with best-effort detected hints (DIMM count / rank) when available.
        var seed = _ramCfg;
        if (_lastCtx != null)
        {
            seed = new RamCalcConfig
            {
                Enabled = _ramCfg.Enabled,
                Die = _ramCfg.Die,
                Rank = _ramCfg.Rank,
                DimmCount = _ramCfg.DimmCount,
                TargetMts = _ramCfg.TargetMts,
                Goal = _ramCfg.Goal,
                MaxDramVoltage = _ramCfg.MaxDramVoltage
            };

            if (_lastCtx.DetectedDimmCount.HasValue && (_lastCtx.DetectedDimmCount.Value == 2 || _lastCtx.DetectedDimmCount.Value == 4))
                seed.DimmCount = _lastCtx.DetectedDimmCount.Value;

            if (_lastCtx.DetectedRankHint == "DualRank") seed.Rank = RamRank.DualRank;
            else if (_lastCtx.DetectedRankHint == "SingleRank") seed.Rank = RamRank.SingleRank;

            if (_lastCtx.EffectiveMemMts.HasValue)
                seed.TargetMts = _lastCtx.EffectiveMemMts.Value;
        }

        var dlg = new RamCalcDialog(seed, _lastCtx) { Owner = this };
        var ok = dlg.ShowDialog();
        if (ok == true)
        {
            _ramCfg = dlg.Config;
            RamCalcConfigStore.Save(_ramCfg);

            // Rebuild table with current ctx if available; otherwise user can hit Generate.
            if (_lastCtx != null)
            {
                // Remove previous RAM calc rows
                _all = _all.Where(r => !(r.Setting ?? "").StartsWith("RAM Calc:", StringComparison.OrdinalIgnoreCase)).ToList();
                if (_ramCfg.Enabled)
                    _all.AddRange(RamTimingCalculator.BuildRecommendations(_lastCtx, _ramCfg));
                ApplyView();
            }
        }
    }

    private void BtnCpuGpuCalc_Click(object sender, RoutedEventArgs e)
    {
        // Seed with best-effort detected hints.
        var seed = _cpuGpuCfg;
        if (_lastCtx != null)
            seed = CpuGpuOcCalculator.SeedFromContext(_lastCtx, _cpuGpuCfg);

        var dlg = new CpuGpuCalcDialog(seed, _lastCtx) { Owner = this };
        var ok = dlg.ShowDialog();
        if (ok == true)
        {
            _cpuGpuCfg = dlg.Config;
            CpuGpuCalcConfigStore.Save(_cpuGpuCfg);

            if (_lastCtx != null)
            {
                // Remove previous calc rows
                _all = _all.Where(r => !((r.Setting ?? "").StartsWith("CPU Calc:", StringComparison.OrdinalIgnoreCase)
                                     || (r.Setting ?? "").StartsWith("GPU Calc:", StringComparison.OrdinalIgnoreCase))).ToList();
                if (_cpuGpuCfg.Enabled)
                {
                    var calcRows = CpuGpuOcCalculator.BuildRecommendations(_lastCtx, _cpuGpuCfg);
                    InsertCpuGpuCalcRowsInPlace(_all, calcRows);
                }
                ApplyView();
            }
        }
    }

    /// <summary>
    /// Mirrors the RAM Calc behavior: CPU/GPU calc rows replace their respective section content in-place,
    /// preserving the original section ordering in the grouped table (CPU stays where CPU was, GPU stays where GPU was).
    /// </summary>
    private static void InsertCpuGpuCalcRowsInPlace(List<Recommendation> list, IEnumerable<Recommendation> calcRows)
    {
        if (list == null) return;
        if (calcRows == null) return;

        var rows = calcRows.ToList();
        if (rows.Count == 0) return;

        var cpuRows = rows.Where(r => string.Equals(r.Section, "CPU", StringComparison.OrdinalIgnoreCase)).ToList();
        var gpuRows = rows.Where(r => string.Equals(r.Section, "GPU", StringComparison.OrdinalIgnoreCase)).ToList();
        var otherRows = rows.Where(r => !string.Equals(r.Section, "CPU", StringComparison.OrdinalIgnoreCase)
                                     && !string.Equals(r.Section, "GPU", StringComparison.OrdinalIgnoreCase)).ToList();

        // Capture insertion points BEFORE inserts so group order stays stable.
        int cpuIndex = list.FindIndex(r => string.Equals(r.Section, "CPU", StringComparison.OrdinalIgnoreCase));
        int gpuIndex = list.FindIndex(r => string.Equals(r.Section, "GPU", StringComparison.OrdinalIgnoreCase));

        // Fallbacks: if a section doesn't exist yet, append to the end.
        if (cpuIndex < 0) cpuIndex = list.Count;
        if (gpuIndex < 0) gpuIndex = list.Count;

        // Insert in ascending order, adjusting the later index for the earlier insert.
        if (cpuIndex <= gpuIndex)
        {
            if (cpuRows.Count > 0)
            {
                list.InsertRange(cpuIndex, cpuRows);
                if (gpuIndex >= cpuIndex) gpuIndex += cpuRows.Count;
            }
            if (gpuRows.Count > 0)
                list.InsertRange(gpuIndex, gpuRows);
        }
        else
        {
            if (gpuRows.Count > 0)
            {
                list.InsertRange(gpuIndex, gpuRows);
                if (cpuIndex >= gpuIndex) cpuIndex += gpuRows.Count;
            }
            if (cpuRows.Count > 0)
                list.InsertRange(cpuIndex, cpuRows);
        }

        // Any non CPU/GPU calc rows (rare) go to the end.
        if (otherRows.Count > 0)
            list.AddRange(otherRows);
    }

    private void BtnCopy_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(_lastText))
            Clipboard.SetText(_lastText);
    }

    
    private void BtnCopyCpu_Click(object sender, RoutedEventArgs e) => CopySectionToClipboard("CPU");
    private void BtnCopyRam_Click(object sender, RoutedEventArgs e) => CopySectionToClipboard("RAM");
    private void BtnCopyGpu_Click(object sender, RoutedEventArgs e) => CopySectionToClipboard("GPU");


    private static bool SectionMatches(string? rowSection, string requested)
    {
        if (string.IsNullOrWhiteSpace(requested) || string.IsNullOrWhiteSpace(rowSection)) return false;

        // UI uses "RAM" label, backend uses "MEMORY"
        if (requested.Equals("RAM", StringComparison.OrdinalIgnoreCase))
            return rowSection.Equals("RAM", StringComparison.OrdinalIgnoreCase) || rowSection.Equals("MEMORY", StringComparison.OrdinalIgnoreCase);

        return rowSection.Equals(requested, StringComparison.OrdinalIgnoreCase);
    }

    private void CopySectionToClipboard(string section)
    {
        if (_all == null || _all.Count == 0) return;

        var lines = new List<string>();
        lines.Add(section.ToUpperInvariant());
        foreach (var r in _all.Where(x => SectionMatches(x.Section, section)))
        {
            // Mirror RAM-calc compact feel: Title: Value
            if (string.IsNullOrWhiteSpace(r.Title)) continue;
            var val = string.IsNullOrWhiteSpace(r.Value) ? "" : r.Value.Trim();
            lines.Add($"- {r.Title}: {val}");
        }

        var text = string.Join(Environment.NewLine, lines).Trim();
        if (!string.IsNullOrWhiteSpace(text))
            Clipboard.SetText(text);
    }

private void BtnSaveTxt_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastText)) return;

        var dlg = new SaveFileDialog
        {
            Filter = "Text File (*.txt)|*.txt",
            FileName = $"OC_Recommendations_{System.Environment.MachineName}.txt"
        };
        if (dlg.ShowDialog() == true)
            File.WriteAllText(dlg.FileName, _lastText);
    }

    private void BtnSaveJson_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastJson)) return;

        var dlg = new SaveFileDialog
        {
            Filter = "JSON File (*.json)|*.json",
            FileName = $"OC_Recommendations_{System.Environment.MachineName}.json"
        };
        if (dlg.ShowDialog() == true)
            File.WriteAllText(dlg.FileName, _lastJson);
    }

private void SectionToggle_Changed(object sender, RoutedEventArgs e)
    {
        if (!_uiReady) return;
        ApplyView();
    }



    // ComboBox.SelectionChanged uses SelectionChangedEventArgs (derived from RoutedEventArgs).
    // Using RoutedEventArgs here can cause an InvalidCastException during XAML event hookup.
    private void Assumptions_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!_uiReady) return;
        UpdateAssumptionsUi();
    }

    private void UpdateAssumptionsUi()
    {
        _assumptions = ReadAssumptionsFromUi();

        // Assumptions impact generated recommendations; refresh the view/output.
        ApplyView();
    }


    private OverclockAssumptions ReadAssumptionsFromUi()
    {
        CoolingType cooling = CoolingType.Aio240;
        SiliconQuality silicon = SiliconQuality.Average;

        try
        {
            var c = (CoolingTypeBox?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "240mm AIO";
            cooling = c switch
            {
                "Air Cooler" => CoolingType.AirCooler,
                "360mm AIO" => CoolingType.Aio360,
                "Custom Loop" => CoolingType.CustomLoop,
                _ => CoolingType.Aio240
            };

            var s = (SiliconQualityBox?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Average";
            silicon = s switch
            {
                "Conservative" => SiliconQuality.Conservative,
                "Above Average" => SiliconQuality.AboveAverage,
                "Golden Sample" => SiliconQuality.GoldenSample,
                _ => SiliconQuality.Average
            };
        }
        catch { }

        return new OverclockAssumptions { Cooling = cooling, Silicon = silicon };
    }

    private void ApplyView()
    {
        if (!_uiReady) return;
        // During initial load, controls can fire events before ItemsSource exists.
        // Avoid NullReference crashes by bailing out until data is ready.
        // Build the filtered list first (so count is correct and grouping works).
        var q = ""; // search removed
        var qLower = q.ToLowerInvariant();

        bool showBios = true;
        bool showTesting = true;

        IEnumerable<Recommendation> seq = _all;
        if (!showBios) seq = seq.Where(r => !string.Equals(r.Section, "BIOS", StringComparison.OrdinalIgnoreCase));
        if (!showTesting) seq = seq.Where(r => !string.Equals(r.Section, "TESTING", StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrWhiteSpace(qLower))
        {
            seq = seq.Where(r =>
                (r.Section ?? "").ToLowerInvariant().Contains(qLower) ||
                (r.Setting ?? "").ToLowerInvariant().Contains(qLower) ||
                (r.Value ?? "").ToLowerInvariant().Contains(qLower) ||
                (r.Path ?? "").ToLowerInvariant().Contains(qLower) ||
                (r.Reason ?? "").ToLowerInvariant().Contains(qLower));
        }

        var list = seq.ToList();

        var view = new ListCollectionView(list);

        bool group = true;
        if (group)
        {
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(Recommendation.Section)));
        }

        _view = view;
        GridRecs.ItemsSource = _view;

    }

    private void SetStatus(string message)
    {
    }

    private void BtnOpenCpuZ_Click(object sender, RoutedEventArgs e)
    {
        OpenUri("https://www.cpuid.com/softwares/cpu-z.html");
    }

    private void BtnOpenEventViewer_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("eventvwr.msc") { UseShellExecute = true });
        }
        catch
        {
            // Fallback: open help page if launching MMC fails.
            OpenUri("https://learn.microsoft.com/windows/security/threat-protection/auditing/event-viewer");
        }

        // NOTE: keep this method properly closed; other handlers follow.
    }
    
    private void BtnOpenAfterburner_Click(object sender, RoutedEventArgs e)
    {
        // Try to launch MSI Afterburner if installed; otherwise open the official page.
        try
        {
            var candidates = new List<string>();

            // Common install locations
            var pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            candidates.Add(Path.Combine(pf86, "MSI Afterburner", "MSIAfterburner.exe"));
            candidates.Add(Path.Combine(pf, "MSI Afterburner", "MSIAfterburner.exe"));

            // Some bundles install under "MSI Afterburner" with different casing
            candidates.Add(Path.Combine(pf86, "MSI Afterburner", "MSI Afterburner.exe"));

            // If user has a custom location, check Start Menu shortcuts quickly
            var startMenu = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
            var userStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);

            foreach (var sm in new[] { startMenu, userStartMenu })
            {
                if (Directory.Exists(sm))
                {
                    try
                    {
                        var lnk = Directory.EnumerateFiles(sm, "*Afterburner*.lnk", SearchOption.AllDirectories).FirstOrDefault();
                        if (!string.IsNullOrWhiteSpace(lnk))
                        {
                            // Let Windows resolve the shortcut.
                            Process.Start(new ProcessStartInfo(lnk) { UseShellExecute = true });
                            return;
                        }
                    }
                    catch { /* ignore */ }
                }
            }

            foreach (var exe in candidates.Distinct())
            {
                if (File.Exists(exe))
                {
                    Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true });
                    return;
                }
            }

            // Fallback: open official page
            OpenUri("https://www.msi.com/Landing/afterburner");
        }
        catch
        {
            OpenUri("https://www.msi.com/Landing/afterburner");
        }
    }


    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        // Support both web URLs and shell targets like eventvwr.msc
        if (e.Uri != null)
            OpenUri(e.Uri.OriginalString);
        e.Handled = true;
    }

    private static void OpenUri(string target)
    {
        if (string.IsNullOrWhiteSpace(target)) return;
        try
        {
            Process.Start(new ProcessStartInfo(target) { UseShellExecute = true });
        }
        catch
        {
            // ignore
        }
    }

    private void GroupExpander_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Expander exp) return;

        // Group name comes from CollectionViewGroup.Name
        var name = exp.DataContext?.GetType().GetProperty("Name")?.GetValue(exp.DataContext)?.ToString();
        if (string.IsNullOrWhiteSpace(name)) return;

        exp.Tag = name;

        // Default: Disclaimer collapsed; everything else expanded.
        var defaultExpanded = !string.Equals(name, "Disclaimer", StringComparison.OrdinalIgnoreCase);

        if (_groupExpanded.TryGetValue(name, out var saved))
            exp.IsExpanded = saved;
        else
            exp.IsExpanded = defaultExpanded;
    }

    private void GroupExpander_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Expander exp) return;
        var name = exp.Tag?.ToString();
        if (string.IsNullOrWhiteSpace(name)) return;

        _groupExpanded[name] = true;
        UiStateStore.Save(_groupExpanded);
    }

    private void GroupExpander_Collapsed(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Expander exp) return;
        var name = exp.Tag?.ToString();
        if (string.IsNullOrWhiteSpace(name)) return;

        _groupExpanded[name] = false;
        UiStateStore.Save(_groupExpanded);
    }

    private static class UiStateStore
    {
        private static string Dir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OCAdvisor");

        private static string PathFile => System.IO.Path.Combine(Dir, "ui_state.json");

        public static Dictionary<string, bool> Load()
        {
            try
            {
                if (!File.Exists(PathFile)) return new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                var json = File.ReadAllText(PathFile);
                var dict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, bool>>(json);
                return dict != null ? new Dictionary<string, bool>(dict, StringComparer.OrdinalIgnoreCase)
                                    : new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public static void Save(Dictionary<string, bool> state)
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var json = System.Text.Json.JsonSerializer.Serialize(state, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(PathFile, json);
            }
            catch { }
        }
    }


}
