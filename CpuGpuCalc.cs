using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;

namespace OcAdvisor;

// ----------------------------
// CPU/GPU OC Calculator (recommendations only)
// Works across models by outputting tiers + step plans + safety rails.
// ----------------------------

internal enum OcGoalTier { Safe, Balanced, Competitive }
internal enum OcCoolingTier { Basic, Decent, Strong }
internal enum OcNoisePreference { Quiet, Normal, Loud }
internal enum DeviceClass { Desktop, Laptop }

internal sealed class CpuGpuCalcConfig
{
    public bool Enabled { get; set; } = false;

    // CPU
    public CpuVendor CpuVendor { get; set; } = CpuVendor.Unknown;
    public DeviceClass CpuClass { get; set; } = DeviceClass.Desktop;

    // GPU
    public GpuVendor GpuVendor { get; set; } = GpuVendor.Unknown;
    public DeviceClass GpuClass { get; set; } = DeviceClass.Desktop;
    public bool AutoFanCurve { get; set; } = true;

    // Shared
    public OcGoalTier Goal { get; set; } = OcGoalTier.Balanced;
    public OcCoolingTier Cooling { get; set; } = OcCoolingTier.Decent;
    public OcNoisePreference Noise { get; set; } = OcNoisePreference.Normal;
}

internal static class CpuGpuCalcConfigStore
{
    private static string ConfigPath
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OCAdvisor");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "cpu_gpu_calc.json");
        }
    }

    public static CpuGpuCalcConfig Load()
    {
        try
        {
            if (!File.Exists(ConfigPath)) return new CpuGpuCalcConfig();
            var json = File.ReadAllText(ConfigPath);
            var cfg = JsonSerializer.Deserialize<CpuGpuCalcConfig>(json);
            return cfg ?? new CpuGpuCalcConfig();
        }
        catch
        {
            return new CpuGpuCalcConfig();
        }
    }

    public static void Save(CpuGpuCalcConfig cfg)
    {
        try
        {
            var json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
        }
        catch
        {
            // ignore
        }
    }
}

internal static class CpuGpuOcCalculator
{
    // --- Public integration helpers ---
    public static CpuGpuCalcConfig SeedFromContext(SystemContext ctx, CpuGpuCalcConfig current)
    {
        var seeded = new CpuGpuCalcConfig
        {
            Enabled = current.Enabled,
            Goal = current.Goal,
            Cooling = current.Cooling,
            Noise = current.Noise,
            AutoFanCurve = current.AutoFanCurve,
            CpuClass = current.CpuClass,
            GpuClass = current.GpuClass,
            CpuVendor = GuessCpuVendor(ctx.CpuName) ?? current.CpuVendor,
            GpuVendor = GuessGpuVendor(ctx.PrimaryGpuName) ?? current.GpuVendor,
        };

        // If we can infer laptop-ish naming (Max-Q / Laptop / Mobile), clamp device class.
        if (LooksLaptopGpu(ctx.PrimaryGpuName)) seeded.GpuClass = DeviceClass.Laptop;
        if (LooksLaptopCpu(ctx.CpuName)) seeded.CpuClass = DeviceClass.Laptop;

        return seeded;
    }

    public static List<Recommendation> BuildRecommendations(SystemContext ctx, CpuGpuCalcConfig cfg)
    {
        var recs = new List<Recommendation>();
        if (!cfg.Enabled) return recs;

        var cpu = CalculateCpu(cfg, ctx);
        recs.AddRange(ToRecommendations("CPU", "CPU Calc:", cpu));

        var gpu = CalculateGpu(cfg, ctx);
        recs.AddRange(ToRecommendations("GPU", "GPU Calc:", gpu));

        return recs;
    }

    // --- CPU calculator ---
    private static CalcBlock CalculateCpu(CpuGpuCalcConfig cfg, SystemContext ctx)
    {
        var vendor = cfg.CpuVendor;
        var goal = ClampGoalForCpu(cfg.Goal, cfg.CpuClass, cfg.Cooling);

        var block = new CalcBlock { Title = "CPU" };

        if (vendor == CpuVendor.AMD)
        {
            // Mirror RAM-calc style: compact numeric outputs (one row per tunable).
            // CO start by goal, then clamp for quiet/basic.
            int coStart = goal switch
            {
                OcGoalTier.Safe => -10,
                OcGoalTier.Balanced => -15,
                _ => -20
            };

            if (cfg.Noise == OcNoisePreference.Quiet) coStart = Math.Max(coStart, -15);
            if (cfg.Cooling == OcCoolingTier.Basic) coStart = Math.Max(coStart, -15);

            // Target CO is a "tune toward" value that stays model-agnostic.
            int coTarget = (goal == OcGoalTier.Safe) ? -10 : -15;

            // Boost override shown as a conditional value, not a separate step plan.
            string boost;
            if (goal == OcGoalTier.Competitive && cfg.Cooling == OcCoolingTier.Strong && cfg.Noise != OcNoisePreference.Quiet)
                boost = "+200 MHz if temps allow";
            else
                boost = "0 MHz";

            block.Values.Add(new CalcValue
            {
                Label = "Curve Optimizer",
                Value = $"All-core {coStart} \u2192 {coTarget} (step 5)",
                Reason = "Gaming-first tuning; adjust in small steps.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "PBO",
                Value = "Enabled (Auto/Motherboard limits)",
                Reason = "Allows opportunistic boost when thermals allow.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Boost Override",
                Value = boost,
                Reason = "Remove first if instability or temp spikes appear.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Thermal Target",
                Value = "85\u201390\u00B0C sustained",
                Reason = "Helps avoid boost oscillation and voltage overshoot.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Rollback Rule",
                Value = "BSOD/WHEA \u2192 reduce CO \u00B7 Temp spikes \u2192 remove boost override first",
                Reason = "Most CO instability is transient overshoot.",
                WarningLevel = "Info"
            });
        }
        else if (vendor == CpuVendor.Intel)
        {
            // Intel: recommend a safe undervolt envelope + temp/power targeting.
            double uvStart = goal switch
            {
                OcGoalTier.Safe => -0.050,
                OcGoalTier.Balanced => -0.070,
                _ => -0.090
            };

            if (cfg.Noise == OcNoisePreference.Quiet) uvStart = Math.Max(uvStart, -0.070);
            if (cfg.Cooling == OcCoolingTier.Basic) uvStart = Math.Max(uvStart, -0.070);

            block.Values.Add(new CalcValue
            {
                Label = "Undervolt (if available)",
                Value = $"Start {uvStart:0.000} V (step 0.010)",
                Reason = "Improves sustained boost by reducing heat/power spikes.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Power / Turbo Limits",
                Value = "Tune for <90\u00B0C sustained",
                Reason = "If undervolt is blocked, temperature targeting is the universal path.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Thermal Target",
                Value = "85\u201390\u00B0C sustained",
                Reason = "Avoids throttling and unstable boost behavior.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Rollback Rule",
                Value = "Crashes \u2192 reduce undervolt \u00B7 Temp spikes \u2192 lower power limits",
                Reason = "Change one variable at a time.",
                WarningLevel = "Info"
            });
        }
        else
        {
            // Generic fallback
            block.Values.Add(new CalcValue
            {
                Label = "Method",
                Value = "Temperature / power targeting",
                Reason = "Universal approach when platform capabilities are unknown.",
                WarningLevel = "Info"
            });

            block.Values.Add(new CalcValue
            {
                Label = "Thermal Target",
                Value = "<90\u00B0C sustained",
                Reason = "Keep temps controlled before increasing aggressiveness.",
                WarningLevel = "Info"
            });
        }

        return block;
    }

    // --- GPU calculator ---
    private static CalcBlock CalculateGpu(CpuGpuCalcConfig cfg, SystemContext ctx)
    {
        var vendor = cfg.GpuVendor;
        var goal = ClampGoalForGpu(cfg.Goal, cfg.GpuClass, cfg.Cooling);

        var block = new CalcBlock { Title = "GPU" };

        // Envelope defaults by goal (model-agnostic)
        int powerMin, powerMax, coreStart, coreCap, memStart, memCapLow, memCapHigh;

        switch (goal)
        {
            case OcGoalTier.Safe:
                powerMin = 0; powerMax = 0;
                coreStart = 0; coreCap = 0;
                memStart = 250; memCapLow = 500; memCapHigh = 500;
                break;

            case OcGoalTier.Balanced:
                powerMin = 5; powerMax = 10;
                coreStart = 30; coreCap = 60;
                memStart = 500; memCapLow = 750; memCapHigh = 750;
                break;

            default: // Competitive
                powerMin = 10; powerMax = 15;
                coreStart = 50; coreCap = 120;
                memStart = 750; memCapLow = 1000; memCapHigh = 2000;
                break;
        }

        // Quiet/basic clamp (favor efficiency)
        if (cfg.Noise == OcNoisePreference.Quiet || cfg.Cooling == OcCoolingTier.Basic)
        {
            powerMax = Math.Min(powerMax, 5);
            coreCap = Math.Min(coreCap, 60);
            memCapHigh = Math.Min(memCapHigh, 1000);
            if (goal == OcGoalTier.Competitive) memStart = Math.Min(memStart, 500);
        }

        string powerValue = powerMax == 0 ? "0%" : $"+{powerMin}\u2013{powerMax}%";
        block.Values.Add(new CalcValue
        {
            Label = "Power Limit",
            Value = powerValue,
            Reason = "Prevents boost throttling when temps are controlled.",
            WarningLevel = "Info"
        });

        string coreValue = coreCap == 0 ? "0 MHz" : $"+{coreStart} MHz \u2192 ~+{coreCap} MHz";
        block.Values.Add(new CalcValue
        {
            Label = "Core Offset",
            Value = coreValue,
            Reason = "Tune gradually; instability usually shows as driver resets.",
            WarningLevel = "Info"
        });

        string memCapText = (memCapLow == memCapHigh) ? $"~+{memCapHigh} MHz" : $"~+{memCapLow}\u2013{memCapHigh} MHz";
        string memValue = $"+{memStart} MHz \u2192 {memCapText}";
        block.Values.Add(new CalcValue
        {
            Label = "Memory Offset",
            Value = memValue,
            Reason = "VRAM OC is the most common cause of artifacts or timeouts.",
            WarningLevel = "Info"
        });

        // Fan strategy mirrored to your style (auto/custom curve is still "automatic")
        string fanMode = cfg.AutoFanCurve ? "Automatic" : "Manual";
        string fanRamp = goal switch
        {
            OcGoalTier.Competitive => "aggressive ramp >65\u00B0C",
            OcGoalTier.Balanced => "mild ramp >70\u00B0C",
            _ => "stock curve"
        };

        block.Values.Add(new CalcValue
        {
            Label = "Fan Strategy",
            Value = $"{fanMode} \u00B7 {fanRamp}",
            Reason = "Stable temps keep clocks steady and reduce frametime spikes.",
            WarningLevel = "Info"
        });

        block.Values.Add(new CalcValue
        {
            Label = "Temperature Target",
            Value = "Core \u226480\u201385\u00B0C \u00B7 Hotspot \u226490\u201395\u00B0C",
            Reason = "Above this, boost oscillation and instability increase.",
            WarningLevel = "Info"
        });

        block.Values.Add(new CalcValue
        {
            Label = "Rollback Rule",
            Value = "Artifacts \u2192 lower memory \u00B7 Reset \u2192 lower core \u00B7 Still unstable \u2192 lower power",
            Reason = "Correct rollback order avoids chasing the wrong limit.",
            WarningLevel = "Info"
        });

        return block;
    }

    // --- Internal structures ---
    private sealed class CalcBlock
    {
        public string Title { get; set; } = "";
        public string Primary { get; set; } = "";
        public List<CalcValue> Values { get; } = new();
        public List<string> Steps { get; } = new();
        public List<string> StopRules { get; } = new();
        public List<string> Notes { get; } = new();
    }

    private sealed class CalcValue
    {
        public string Label { get; set; } = "";
        public string Value { get; set; } = "";
        public string Reason { get; set; } = "Recommended starting value.";
        public string WarningLevel { get; set; } = "Info"; // Info/Warning/Danger
    }

    private static List<Recommendation> ToRecommendations(string section, string prefix, CalcBlock block)
    {
        // Mirror RAM calc formatting:
        // - One row per tunable
        // - Short labels
        // - Numeric/compact values
        // - No "Method", "Step", "Stop Rule", "Note" spam rows
        var outRecs = new List<Recommendation>();
        const string path = "Calculator";

        foreach (var v in block.Values.Where(v => !string.IsNullOrWhiteSpace(v.Label) && !string.IsNullOrWhiteSpace(v.Value)))
        {
            outRecs.Add(new Recommendation
            {
                Section = section,
                Setting = $"{prefix} {v.Label}".Trim(),
                Value = v.Value,
                Path = path,
                Reason = v.Reason ?? "",
                WarningLevel = string.IsNullOrWhiteSpace(v.WarningLevel) ? "Info" : v.WarningLevel
            });
        }

        return outRecs;
    }

    // --- Goal clamps ---
    private static OcGoalTier ClampGoalForCpu(OcGoalTier requested, DeviceClass cls, OcCoolingTier cooling)
    {
        var g = (int)requested;
        if (cls == DeviceClass.Laptop) g = Math.Min(g, (int)OcGoalTier.Balanced);
        if (cooling == OcCoolingTier.Basic) g = Math.Min(g, (int)OcGoalTier.Balanced);
        return (OcGoalTier)Math.Max((int)OcGoalTier.Safe, Math.Min((int)OcGoalTier.Competitive, g));
    }

    private static OcGoalTier ClampGoalForGpu(OcGoalTier requested, DeviceClass cls, OcCoolingTier cooling)
    {
        var g = (int)requested;
        if (cls == DeviceClass.Laptop) g = Math.Min(g, (int)OcGoalTier.Balanced);
        if (cooling == OcCoolingTier.Basic) g = Math.Min(g, (int)OcGoalTier.Balanced);
        return (OcGoalTier)Math.Max((int)OcGoalTier.Safe, Math.Min((int)OcGoalTier.Competitive, g));
    }

    // --- Vendor heuristics ---
    private static CpuVendor? GuessCpuVendor(string cpuName)
    {
        var s = (cpuName ?? "").ToLowerInvariant();
        if (s.Contains("amd") || s.Contains("ryzen") || s.Contains("threadripper")) return CpuVendor.AMD;
        if (s.Contains("intel") || s.Contains("core(tm)") || s.Contains("xeon")) return CpuVendor.Intel;
        return null;
    }

    private static GpuVendor? GuessGpuVendor(string gpuName)
    {
        var s = (gpuName ?? "").ToLowerInvariant();
        if (s.Contains("nvidia") || s.Contains("geforce") || s.Contains("rtx") || s.Contains("gtx")) return GpuVendor.NVIDIA;
        if (s.Contains("amd") || s.Contains("radeon") || s.Contains("rx ")) return GpuVendor.AMD;
        return null;
    }

    private static bool LooksLaptopGpu(string gpuName)
    {
        var s = (gpuName ?? "").ToLowerInvariant();
        return s.Contains("laptop") || s.Contains("mobile") || s.Contains("max-q") || s.Contains("notebook");
    }

    private static bool LooksLaptopCpu(string cpuName)
    {
        var s = (cpuName ?? "").ToLowerInvariant();
        return s.Contains("mobile") || s.Contains("u-") || s.Contains("hs") || s.Contains("h ");
    }
}

// ----------------------------
// UI dialog (code-only, matches style of RamCalcDialog)
// ----------------------------
internal sealed class CpuGpuCalcDialog : Window
{
    public CpuGpuCalcConfig Config { get; private set; }

    private readonly CheckBox _enabled;
    private readonly ComboBox _goal;
    private readonly ComboBox _cooling;
    private readonly ComboBox _noise;
    private readonly ComboBox _cpuVendor;
    private readonly ComboBox _cpuClass;
    private readonly ComboBox _gpuVendor;
    private readonly ComboBox _gpuClass;
    private readonly CheckBox _autoFan;
    private readonly TextBox _preview;

    private readonly SystemContext? _ctx;

    public CpuGpuCalcDialog(CpuGpuCalcConfig initial, SystemContext? detected = null)
    {
        _ctx = detected;

        Title = "CPU/GPU OC Calculators";
        Width = 720;
        Height = 540;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#0F0F14");
        Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0");

        // Dark dropdowns + readable text (prevents light popup + washed-out items)
        Resources[System.Windows.SystemColors.WindowBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1C2230");
        Resources[System.Windows.SystemColors.WindowTextBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0");
        Resources[System.Windows.SystemColors.HighlightBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2A3346");
        Resources[System.Windows.SystemColors.HighlightTextBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0");
        Resources[System.Windows.SystemColors.ControlBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#151922");
        Resources[System.Windows.SystemColors.ControlTextBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0");
        Resources[System.Windows.SystemColors.ControlLightBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1C2230");
        Resources[System.Windows.SystemColors.ControlLightLightBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1C2230");
        Resources[System.Windows.SystemColors.ControlDarkBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#0F1115");
        Resources[System.Windows.SystemColors.ControlDarkDarkBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#0F1115");
        Resources[System.Windows.SystemColors.GrayTextBrushKey] = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#AEB6C6");

        var comboStyle = new Style(typeof(ComboBox));
        comboStyle.Setters.Add(new Setter(ComboBox.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EDEDED")));
        comboStyle.Setters.Add(new Setter(ComboBox.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#000000")));
        comboStyle.Setters.Add(new Setter(ComboBox.BorderBrushProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2A3346")));
        comboStyle.Setters.Add(new Setter(ComboBox.BorderThicknessProperty, new Thickness(1)));
        Resources[typeof(ComboBox)] = comboStyle;

        var comboItemStyle = new Style(typeof(ComboBoxItem));
        comboItemStyle.Setters.Add(new Setter(ComboBoxItem.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1C2230")));
        comboItemStyle.Setters.Add(new Setter(ComboBoxItem.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0")));
        var hi = new Trigger { Property = ComboBoxItem.IsHighlightedProperty, Value = true };
        hi.Setters.Add(new Setter(ComboBoxItem.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2A3346")));
        comboItemStyle.Triggers.Add(hi);
        Resources[typeof(ComboBoxItem)] = comboItemStyle;


        Config = new CpuGpuCalcConfig
        {
            Enabled = initial.Enabled,
            Goal = initial.Goal,
            Cooling = initial.Cooling,
            Noise = initial.Noise,
            CpuVendor = initial.CpuVendor,
            CpuClass = initial.CpuClass,
            GpuVendor = initial.GpuVendor,
            GpuClass = initial.GpuClass,
            AutoFanCurve = initial.AutoFanCurve
        };

        var root = new Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Content = root;

        _enabled = new CheckBox { Content = "Enable CPU/GPU OC calculators", IsChecked = Config.Enabled, Margin = new Thickness(0, 0, 0, 10) };
        root.Children.Add(_enabled);

        var form = new Grid { Margin = new Thickness(0, 0, 0, 10) };
        Grid.SetRow(form, 1);
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (int i = 0; i < 8; i++) form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.Children.Add(form);

        int r = 0;
        AddRow(form, r++, "Goal", _goal = MakeCombo(typeof(OcGoalTier), Config.Goal));
        AddRow(form, r++, "Cooling", _cooling = MakeCombo(typeof(OcCoolingTier), Config.Cooling));
        AddRow(form, r++, "Noise", _noise = MakeCombo(typeof(OcNoisePreference), Config.Noise));
        AddRow(form, r++, "CPU Vendor", _cpuVendor = MakeCombo(typeof(CpuVendor), Config.CpuVendor));
        AddRow(form, r++, "CPU Class", _cpuClass = MakeCombo(typeof(DeviceClass), Config.CpuClass));
        AddRow(form, r++, "GPU Vendor", _gpuVendor = MakeCombo(typeof(GpuVendor), Config.GpuVendor));
        AddRow(form, r++, "GPU Class", _gpuClass = MakeCombo(typeof(DeviceClass), Config.GpuClass));

        _autoFan = new CheckBox { Content = "GPU fans on automatic control (stock or custom curve)", IsChecked = Config.AutoFanCurve, Margin = new Thickness(0, 6, 0, 0) };
        Grid.SetRow(_autoFan, r);
        Grid.SetColumn(_autoFan, 1);
        form.Children.Add(new TextBlock { Text = "Fan Mode", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 8, 10, 0) });
        Grid.SetRow(form.Children[^1], r);
        Grid.SetColumn(form.Children[^1], 0);
        form.Children.Add(_autoFan);
        r++;

        _preview = new TextBox
        {
            IsReadOnly = true,
            TextWrapping = TextWrapping.Wrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#101018"),
            Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0"),
            BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2A2A35"),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0, 0, 0, 10)
        };
        Grid.SetRow(_preview, 2);
        root.Children.Add(_preview);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        Grid.SetRow(buttons, 3);
        root.Children.Add(buttons);

        var btnCopy = new Button { Content = "Copy Preview", Width = 130, Margin = new Thickness(0, 0, 8, 0) };
        btnCopy.Click += (_, __) => { if (!string.IsNullOrWhiteSpace(_preview.Text)) Clipboard.SetText(_preview.Text); };
        buttons.Children.Add(btnCopy);

        var btnOk = new Button { Content = "OK", Width = 130, Margin = new Thickness(0, 0, 8, 0) };
        btnOk.Click += (_, __) => { SaveBack(); DialogResult = true; Close(); };
        buttons.Children.Add(btnOk);

        var btnCancel = new Button { Content = "Cancel", Width = 130 };
        btnCancel.Click += (_, __) => { DialogResult = false; Close(); };
        buttons.Children.Add(btnCancel);

        // Premium button styling (primary/secondary)
        try
        {
            btnCopy.Style = (Style)Application.Current.FindResource("SecondaryButton");
            btnCancel.Style = (Style)Application.Current.FindResource("SecondaryButton");
            btnOk.Style = (Style)Application.Current.FindResource("PrimaryButton");
        }
        catch { /* ignore if styles missing */ }

        // Re-render preview on change
        _enabled.Checked += (_, __) => RenderPreview();
        _enabled.Unchecked += (_, __) => RenderPreview();
        _goal.SelectionChanged += (_, __) => RenderPreview();
        _cooling.SelectionChanged += (_, __) => RenderPreview();
        _noise.SelectionChanged += (_, __) => RenderPreview();
        _cpuVendor.SelectionChanged += (_, __) => RenderPreview();
        _cpuClass.SelectionChanged += (_, __) => RenderPreview();
        _gpuVendor.SelectionChanged += (_, __) => RenderPreview();
        _gpuClass.SelectionChanged += (_, __) => RenderPreview();
        _autoFan.Checked += (_, __) => RenderPreview();
        _autoFan.Unchecked += (_, __) => RenderPreview();

        RenderPreview();
    }

    private void SaveBack()
    {
        Config.Enabled = _enabled.IsChecked == true;
        Config.Goal = (OcGoalTier)(_goal.SelectedItem ?? OcGoalTier.Balanced);
        Config.Cooling = (OcCoolingTier)(_cooling.SelectedItem ?? OcCoolingTier.Decent);
        Config.Noise = (OcNoisePreference)(_noise.SelectedItem ?? OcNoisePreference.Normal);
        Config.CpuVendor = (CpuVendor)(_cpuVendor.SelectedItem ?? CpuVendor.Unknown);
        Config.CpuClass = (DeviceClass)(_cpuClass.SelectedItem ?? DeviceClass.Desktop);
        Config.GpuVendor = (GpuVendor)(_gpuVendor.SelectedItem ?? GpuVendor.Unknown);
        Config.GpuClass = (DeviceClass)(_gpuClass.SelectedItem ?? DeviceClass.Desktop);
        Config.AutoFanCurve = _autoFan.IsChecked == true;
    }

    private void RenderPreview()
    {
        SaveBack();
        if (!Config.Enabled)
        {
            _preview.Text = "Disabled. Enable the calculators to generate CPU/GPU tuning plans.";
            return;
        }

        // Build a text preview from the same calculator used to append rows.
        // We don't depend on sensors; this is model-agnostic.
        SystemContext fakeCtx;
        if (_ctx != null)
        {
            fakeCtx = _ctx;
        }
        else
        {
            fakeCtx = new SystemContext { CpuName = "(Detected CPU)" };
            fakeCtx.Gpus.Add(new GpuInfo { Name = "(Detected GPU)", Vendor = GpuVendor.Unknown });
        }
        var sb = new StringBuilder();
        sb.AppendLine("CPU/GPU OC Calculator Preview");
        sb.AppendLine();
        sb.AppendLine($"Goal: {Config.Goal} | Cooling: {Config.Cooling} | Noise: {Config.Noise}");
        sb.AppendLine();
        sb.AppendLine($"CPU: {(string.IsNullOrWhiteSpace(fakeCtx.CpuName) ? "(unknown)" : fakeCtx.CpuName)}");
        sb.AppendLine($"GPU: {(string.IsNullOrWhiteSpace(fakeCtx.PrimaryGpuName) ? "(unknown)" : fakeCtx.PrimaryGpuName)}");
        sb.AppendLine();

        foreach (var rec in CpuGpuOcCalculator.BuildRecommendations(fakeCtx, Config)
                 .Where(r => (r.Setting ?? "").StartsWith("CPU Calc:") || (r.Setting ?? "").StartsWith("GPU Calc:"))
                 .Take(60))
        {
            sb.AppendLine($"• {rec.Value}");
        }

        _preview.Text = sb.ToString();
    }

    private static void AddRow(Grid grid, int row, string label, FrameworkElement editor)
    {
        var lbl = new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 10, 0) };
        Grid.SetRow(lbl, row);
        Grid.SetColumn(lbl, 0);
        grid.Children.Add(lbl);

        Grid.SetRow(editor, row);
        Grid.SetColumn(editor, 1);
        editor.Margin = new Thickness(0, 4, 0, 0);
        grid.Children.Add(editor);
    }

    private static ComboBox MakeCombo(Type enumType, object selected)
    {
        var items = Enum.GetValues(enumType).Cast<object>().ToList();
        var cb = new ComboBox
        {
            ItemsSource = items,
            SelectedItem = selected,
            Height = 30,
            MinWidth = 220
        };
        return cb;
    }
}
