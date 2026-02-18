using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;

namespace OcAdvisor;

internal enum RamDieType
{
    Unknown,
    Samsung_BDie,
    Samsung_CDie,
    Hynix_CJR,
    Hynix_DJR,
    Micron_EDie,
    Micron_RevB,
    Hynix_ADie,
    Hynix_MDie,
    Samsung_DDie,
    Micron_DDie
}

internal enum RamGoal
{
    Safe,
    Balanced,
    Fast,
    Extreme
}

internal enum RamRank
{
    SingleRank,
    DualRank
}

internal sealed class RamCalcConfig
{
    public bool Enabled { get; set; } = false;
    public RamDieType Die { get; set; } = RamDieType.Unknown;
    public RamRank Rank { get; set; } = RamRank.SingleRank;
    public int DimmCount { get; set; } = 2; // 2 or 4
    public int TargetMts { get; set; } = 3600;
    public RamGoal Goal { get; set; } = RamGoal.Safe;
    public double MaxDramVoltage { get; set; } = 1.40; // user cap
}

internal static class RamCalcConfigStore
{
    private static string ConfigPath
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OCAdvisor");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "ram_calc.json");
        }
    }

    public static RamCalcConfig Load()
    {
        try
        {
            if (!File.Exists(ConfigPath)) return new RamCalcConfig();
            var json = File.ReadAllText(ConfigPath);
            var cfg = JsonSerializer.Deserialize<RamCalcConfig>(json);
            return cfg ?? new RamCalcConfig();
        }
        catch
        {
            return new RamCalcConfig();
        }
    }

    public static void Save(RamCalcConfig cfg)
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

internal sealed class RamCalcDialog : Window
{
    public RamCalcConfig Config { get; private set; }

    private readonly CheckBox _enabled;
    private readonly ComboBox _die;
    private readonly ComboBox _rank;
    private readonly ComboBox _dimms;
    private readonly ComboBox _goal;
    private readonly ComboBox _mts;
    private readonly TextBox _vmax;

    private readonly Border _confidencePill;
    private readonly TextBlock _confidenceText;
    private readonly TextBlock _autoHint;

    public RamCalcDialog(RamCalcConfig initial, SystemContext? detected = null)
    {
        Title = "RAM Timing Calculator";
        Width = 520;
        Height = 420;
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
        comboStyle.Setters.Add(new Setter(ComboBox.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1C2230")));
        comboStyle.Setters.Add(new Setter(ComboBox.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EAEAF0")));
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


        Config = new RamCalcConfig
        {
            Enabled = initial.Enabled,
            Die = initial.Die,
            Rank = initial.Rank,
            DimmCount = initial.DimmCount,
            TargetMts = initial.TargetMts,
            Goal = initial.Goal,
            MaxDramVoltage = initial.MaxDramVoltage
        };

        // Apply best-effort detection hints (user can override).
        if (detected != null)
        {
            if (detected.DetectedDimmCount.HasValue && (detected.DetectedDimmCount.Value == 2 || detected.DetectedDimmCount.Value == 4))
                Config.DimmCount = detected.DetectedDimmCount.Value;

            if (string.Equals(detected.DetectedRankHint, "DualRank", StringComparison.OrdinalIgnoreCase))
                Config.Rank = RamRank.DualRank;
            else if (string.Equals(detected.DetectedRankHint, "SingleRank", StringComparison.OrdinalIgnoreCase))
                Config.Rank = RamRank.SingleRank;

            if (detected.EffectiveMemMts.HasValue)
                Config.TargetMts = detected.EffectiveMemMts.Value;
        }

        var root = new Grid { Margin = new Thickness(16) };
        // 0: enabled, 1: form, 2: confidence/hints, 3: note, 4: buttons
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        _enabled = new CheckBox { Content = "Enable RAM timing calculator", Margin = new Thickness(0, 0, 0, 12), IsChecked = Config.Enabled };
        root.Children.Add(_enabled);

        var form = new Grid();
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        for (int i = 0; i < 6; i++)
            form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        int r = 0;
        var ctx = detected;
        AddRow(form, r++, "Memory IC / Die", _die = Combo(RamDieType.Unknown, Enum.GetValues(typeof(RamDieType)).Cast<RamDieType>()));
        AddRow(form, r++, "Rank", _rank = Combo(Config.Rank, Enum.GetValues(typeof(RamRank)).Cast<RamRank>()));
        AddRow(form, r++, "DIMM count", _dimms = Combo(Config.DimmCount, new[] { 2, 4 }));
        var mtsOptions = (ctx != null && ctx.MemoryType == MemoryType.DDR5)
            ? new[] { 4000, 4200, 4400, 4600, 4800, 5000, 5200, 5400, 5600, 5800, 6000, 6200, 6400, 6600, 6800, 7000, 7200, 7400, 7600, 7800, 8000, 8200, 8400 }
            : new[] { 3200, 3333, 3466, 3600, 3733, 3800, 3866, 4000 };
        AddRow(form, r++, "Target speed (MT/s)", _mts = Combo(Config.TargetMts, mtsOptions));
                _mts.IsEditable = true;
        _mts.IsTextSearchEnabled = false;
AddRow(form, r++, "Goal", _goal = Combo(Config.Goal, Enum.GetValues(typeof(RamGoal)).Cast<RamGoal>()));

        _vmax = new TextBox { Text = Config.MaxDramVoltage.ToString("0.00"), Margin = new Thickness(0, 4, 0, 10) };
        AddRow(form, r++, "Max DRAM V (cap)", _vmax);

        Grid.SetRow(form, 1);
        root.Children.Add(form);

        // Confidence meter (honest, not pretending certainty)
        var confRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        _confidencePill = new Border
        {
            CornerRadius = new CornerRadius(999),
            Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1B2238"),
            Padding = new Thickness(10, 4, 10, 4),
            BorderThickness = new Thickness(1),
            BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2A2A35")
        };
        _confidenceText = new TextBlock { Text = "Confidence: —", VerticalAlignment = VerticalAlignment.Center };
        _confidencePill.Child = _confidenceText;
        confRow.Children.Add(_confidencePill);

        _autoHint = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.85,
            Margin = new Thickness(10, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        confRow.Children.Add(_autoHint);

        Grid.SetRow(confRow, 2);
        root.Children.Add(confRow);

        var note = new TextBlock
        {
            Text = "Pick the correct IC/rank. Wrong die selection = wrong subtimings/voltage. Output is a starting point — stability test required.",
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.85,
            Margin = new Thickness(0, 10, 0, 0)
        };
        Grid.SetRow(note, 3);
        root.Children.Add(note);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        var btnCancel = new Button { Content = "Cancel", Width = 120, Margin = new Thickness(0, 0, 8, 0) };
        var btnOk = new Button { Content = "Save", Width = 120 };

        // Premium button styling (primary/secondary)
        try
        {
            btnCancel.Style = (Style)Application.Current.FindResource("SecondaryButton");
            btnOk.Style = (Style)Application.Current.FindResource("PrimaryButton");
        }
        catch { /* ignore if styles missing */ }
        btnCancel.Click += (_, __) => { DialogResult = false; Close(); };
        btnOk.Click += (_, __) => { if (TryApply()) { DialogResult = true; Close(); } };
        buttons.Children.Add(btnCancel);
        buttons.Children.Add(btnOk);

        Grid.SetRow(buttons, 4);
        root.Children.Add(buttons);

        Content = root;

        // init values
        _die.SelectedItem = Config.Die;

        // Wire live updates for confidence + auto-hints
        _die.SelectionChanged += (_, __) => UpdateConfidenceAndHint(detected);
        _rank.SelectionChanged += (_, __) => UpdateConfidenceAndHint(detected);
        _dimms.SelectionChanged += (_, __) => UpdateConfidenceAndHint(detected);
        _mts.SelectionChanged += (_, __) => UpdateConfidenceAndHint(detected);
        _goal.SelectionChanged += (_, __) => UpdateConfidenceAndHint(detected);
        _vmax.TextChanged += (_, __) => UpdateConfidenceAndHint(detected);

        UpdateConfidenceAndHint(detected);
    }

    private void UpdateConfidenceAndHint(SystemContext? detected)
    {
        // Read current selections (best-effort).
        var die = (RamDieType)(_die.SelectedItem ?? Config.Die);
        var rank = (RamRank)(_rank.SelectedItem ?? Config.Rank);
        var dimms = (int)(_dimms.SelectedItem ?? Config.DimmCount);
        var mts = Config.TargetMts;
        if (_mts.SelectedItem is int mi) mts = mi;
        else if (int.TryParse((_mts.Text ?? "").Trim(), out var mtParsed)) mts = mtParsed;
        var goal = (RamGoal)(_goal.SelectedItem ?? Config.Goal);
        double vmax = Config.MaxDramVoltage;
        if (double.TryParse((_vmax.Text ?? "").Trim(), out var tmp)) vmax = tmp;

        // Simple, honest scoring.
        int score = 100;
        if (die == RamDieType.Unknown) score -= 40;
        if (goal == RamGoal.Fast) score -= 10;
        if (goal == RamGoal.Extreme) score -= 25;
        if (dimms == 4) score -= 20;
        if (rank == RamRank.DualRank) score -= 10;
        if (mts >= 4000) score -= 20;
        else if (mts >= 3866) score -= 12;
        else if (mts >= 3800) score -= 8;
        if (vmax > 1.50) score -= 15;
        else if (vmax > 1.45) score -= 8;
        if (score < 0) score = 0;

        string label;
        string pill;
        if (score >= 75) { label = "High"; pill = "#162A22"; }
        else if (score >= 45) { label = "Medium"; pill = "#2A2416"; }
        else { label = "Low"; pill = "#2A1717"; }

        _confidencePill.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(pill);
        _confidenceText.Text = $"Confidence: {label}";

        // Auto-detect hint line
        if (detected != null)
        {
            var parts = new List<string>();
            if (detected.DetectedDimmCount.HasValue) parts.Add($"{detected.DetectedDimmCount.Value} DIMMs");
            if (!string.IsNullOrWhiteSpace(detected.DetectedRankHint)) parts.Add(detected.DetectedRankHint);
            if (detected.EffectiveMemMts.HasValue) parts.Add($"{detected.EffectiveMemMts.Value} MT/s");
            _autoHint.Text = parts.Count > 0 ? $"Auto-detected: {string.Join(" | ", parts)} (verify)" : "";
        
            // Live preview of computed primaries so users can see Goal/Speed changes immediately.
            try
            {
                var preview = RamTimingCalculator.BuildPreview(detected, die, rank, dimms, mts, goal, vmax);
                if (!string.IsNullOrWhiteSpace(preview))
                    _autoHint.Text = (string.IsNullOrWhiteSpace(_autoHint.Text) ? "" : _autoHint.Text + "\n") + preview;
            }
            catch { /* preview is best-effort */ }
}
        else
        {
            _autoHint.Text = "";
        }
    }

    private static void AddRow(Grid g, int row, string label, UIElement input)
    {
        var tb = new TextBlock { Text = label, Margin = new Thickness(0, 6, 10, 6), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(tb, row);
        Grid.SetColumn(tb, 0);
        g.Children.Add(tb);

        Grid.SetRow(input, row);
        Grid.SetColumn(input, 1);
        g.Children.Add(input);

        if (input is Control c)
            c.Margin = new Thickness(0, 4, 0, 4);
    }

    private static ComboBox Combo<T>(T selected, IEnumerable<T> items)
    {
        var cb = new ComboBox { ItemsSource = items.ToList(), SelectedItem = selected };
        cb.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#EDEDED");
        cb.Foreground = System.Windows.Media.Brushes.Black;
        return cb;
    }

    private bool TryApply()
    {
        Config.Enabled = _enabled.IsChecked == true;
        Config.Die = (RamDieType)(_die.SelectedItem ?? RamDieType.Unknown);
        Config.Rank = (RamRank)(_rank.SelectedItem ?? RamRank.SingleRank);
        Config.DimmCount = (int)(_dimms.SelectedItem ?? 2);
        Config.TargetMts = (int)(_mts.SelectedItem ?? 3600);
        Config.Goal = (RamGoal)(_goal.SelectedItem ?? RamGoal.Safe);

        if (!double.TryParse(_vmax.Text.Trim(), out var vmax)) vmax = 1.40;
        vmax = Math.Max(1.30, Math.Min(1.55, vmax));
        Config.MaxDramVoltage = vmax;

        // basic sanity
        if (Config.Die == RamDieType.Unknown)
        {
            MessageBox.Show(this, "Select a memory IC / die type (or the calculator will be guessy).", "RAM Calc", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        return true;
    }
}

internal static class RamTimingCalculator
{
    // A pragmatic, DRAM-Calculator-style approach: IC-based baselines + goal modifiers.
    // Values are starting points. Users must validate.

    private sealed record TimingSet(
        int tCL, int tRCDRD, int tRCDWR, int tRP, int tRAS,
        int tRC, int tCWL,
        int tRRDS, int tRRDL, int tFAW,
        int tWTRS, int tWTRL, int tWR,
        int tRFC,
        int tREFI,
        double vDram
    );

    internal static string BuildPreview(SystemContext? detected, RamDieType die, RamRank rank, int dimms, int mts, RamGoal goal, double vmax)
    {
        var memType = detected?.MemoryType ?? InferMemoryType(die, mts);

        // Build a minimal context for the calculator
        var ctx = detected ?? new SystemContext
        {
            CpuVendor = CpuVendor.Unknown,
            BoardVendor = BoardVendor.Unknown,
            MemoryType = memType
        };

        var cfg = new RamCalcConfig
        {
            Enabled = true,
            Die = die,
            Rank = rank,
            DimmCount = dimms,
            TargetMts = mts,
            Goal = goal,
            MaxDramVoltage = vmax
        };

        TimingSet set;
        if (memType == MemoryType.DDR5)
        {
            set = GetBaselineDdr5(cfg.Die, cfg.TargetMts);
            set = ApplyTopologyModifiers(set, cfg);
            set = ApplyGoalDdr5(set, cfg.Goal);
            set = ApplyVoltageCap(set, cfg.MaxDramVoltage, cfg.Goal, cfg.Die);
        }
        else
        {
            set = GetBaseline(cfg.Die, cfg.TargetMts);
            set = ApplyTopologyModifiers(set, cfg);
            set = ApplyGoal(set, cfg.Goal);
            set = ApplyVoltageCap(set, cfg.MaxDramVoltage, cfg.Goal, cfg.Die);
        }

        return $"Preview: {memType} {mts} MT/s | CL{set.tCL}-RCD{set.tRCDRD}-RP{set.tRP}-RAS{set.tRAS} | tRFC {set.tRFC} | V {set.vDram:0.000}V";
    }

    internal static MemoryType InferMemoryType(RamDieType die, int mts)
    {
        // Heuristic: DDR5 starts at 4800 JEDEC, but allow lower if user inputs it.
        // If the kit is explicitly DDR5 in detected context, that path is used above.
        // Here we bias toward DDR5 when the speed is high enough.
        if (mts >= 4400) return MemoryType.DDR5;
        return MemoryType.DDR4;
    }

    public static List<Recommendation> BuildRecommendations(SystemContext ctx, RamCalcConfig cfg)
    {
        var recs = new List<Recommendation>();

        // Heading / inputs row
        recs.Add(new Recommendation
        {
            Section = "MEMORY",
            Setting = "RAM Calc: Inputs",
            Value = $"{cfg.Die} | {cfg.Rank} | {cfg.DimmCount} DIMMs | {cfg.TargetMts} MT/s | {cfg.Goal} | DRAM cap {cfg.MaxDramVoltage:0.00}V",
            Path = "",
            Reason = "These are generated timing starting points based on your selected IC and target speed. Verify stability.",
            WarningLevel = "Info"
        });

        // If we can't calculate for this platform, still return the header.
        if (ctx.MemoryType != MemoryType.DDR4 && ctx.MemoryType != MemoryType.DDR5)
        {
            recs.Add(new Recommendation
            {
                Section = "MEMORY",
                Setting = "RAM Calc: Unsupported",
                Value = "RAM timing calculator currently supports DDR4 and DDR5",
                Reason = "Your detected memory type is not DDR4/DDR5.",
                WarningLevel = "Warning"
            });
            return recs;
        }

        TimingSet set;
        if (ctx.MemoryType == MemoryType.DDR5)
        {
            set = GetBaselineDdr5(cfg.Die, cfg.TargetMts);
            set = ApplyGoalDdr5(set, cfg.Goal);
            set = ApplyVoltageCap(set, cfg.MaxDramVoltage, cfg.Goal, cfg.Die);

            recs.Add(new Recommendation
            {
                Section = "MEMORY",
                Setting = "RAM Calc: DDR5 note",
                Value = "Start with XMP/EXPO; treat these as stability-first starting points.",
                Reason = "DDR5 training/IMC limits vary a lot by CPU and board; validate stability.",
                WarningLevel = "Info"
            });
        }
        else
        {
            set = GetBaseline(cfg.Die, cfg.TargetMts);
            set = ApplyTopologyModifiers(set, cfg);
            set = ApplyGoal(set, cfg.Goal);
            set = ApplyVoltageCap(set, cfg.MaxDramVoltage, cfg.Goal, cfg.Die);
        }

        // Fabric note
        var idealFclk = ctx.IdealFclkMHz;
        var recFclk = ctx.TargetFclkMHz;
        if (idealFclk.HasValue && recFclk.HasValue && idealFclk.Value > recFclk.Value)
        {
            recs.Add(new Recommendation
            {
                Section = "MEMORY",
                Setting = "RAM Calc: Fabric note",
                Value = $"Ideal FCLK {idealFclk} MHz (1:1), but stability-first cap suggests {recFclk} MHz",
                Path = VendorizeTimingPath(ctx, "Infinity Fabric"),
                Reason = "Zen 3 often walls around 1900–2000 FCLK depending on silicon. If you see WHEA 19, reduce FCLK or adjust VDDG/VDDP.",
                WarningLevel = (idealFclk.Value >= 2000) ? "Danger" : "Warning"
            });
        }

        var timingPath = VendorizeTimingPath(ctx, "DRAM Timing Control");

        // Primaries
        AddTiming(recs, "tCL", set.tCL.ToString(), timingPath, "Primary timing. Lower reduces latency; too low causes no-POST/errors.", LevelFor(cfg));
        AddTiming(recs, "tRCDRD", set.tRCDRD.ToString(), timingPath, "Read activate to read delay.", LevelFor(cfg));
        AddTiming(recs, "tRCDWR", set.tRCDWR.ToString(), timingPath, "Write activate to write delay.", LevelFor(cfg));
        AddTiming(recs, "tRP", set.tRP.ToString(), timingPath, "Row precharge.", LevelFor(cfg));
        AddTiming(recs, "tRAS", set.tRAS.ToString(), timingPath, "Row active time.", LevelFor(cfg));
        AddTiming(recs, "tRC", set.tRC.ToString(), timingPath, "tRAS + tRP baseline.", LevelFor(cfg));
        AddTiming(recs, "tCWL", set.tCWL.ToString(), timingPath, "CAS write latency.", LevelFor(cfg));

        // Key secondaries
        AddTiming(recs, "tRRDS", set.tRRDS.ToString(), timingPath, "Short row-to-row.", LevelFor(cfg));
        AddTiming(recs, "tRRDL", set.tRRDL.ToString(), timingPath, "Long row-to-row.", LevelFor(cfg));
        AddTiming(recs, "tFAW", set.tFAW.ToString(), timingPath, "Four activate window.", LevelFor(cfg));
        AddTiming(recs, "tWTRS", set.tWTRS.ToString(), timingPath, "Write-to-read short.", LevelFor(cfg));
        AddTiming(recs, "tWTRL", set.tWTRL.ToString(), timingPath, "Write-to-read long.", LevelFor(cfg));
        AddTiming(recs, "tWR", set.tWR.ToString(), timingPath, "Write recovery.", LevelFor(cfg));
        AddTiming(recs, "tRFC", set.tRFC.ToString(), timingPath, "Refresh cycle time. Too tight causes memory errors under load.", (cfg.Goal == RamGoal.Extreme) ? "Danger" : "Warning");
        AddTiming(recs, "tREFI", set.tREFI.ToString(), timingPath, "Refresh interval. Higher improves perf but can reduce stability/temps.", (cfg.Goal == RamGoal.Extreme) ? "Warning" : "Info");

        // Voltages
        recs.Add(new Recommendation
        {
            Section = "MEMORY",
            Setting = "RAM Calc: DRAM Voltage",
            Value = $"{set.vDram:0.000} V",
            Path = VendorizeVoltagePath(ctx, "DRAM Voltage"),
            Reason = "Starting point for the selected die + goal. Stay within your cap; higher increases heat and risk.",
            WarningLevel = (set.vDram > 1.45) ? "Danger" : (set.vDram > 1.40 ? "Warning" : "Info")
        });

        return recs;
    }

    private static string VendorizeTimingPath(SystemContext ctx, string leaf)
    {
        return ctx.BoardVendor switch
        {
            BoardVendor.ASUS => $"AI Tweaker → {leaf}",
            BoardVendor.MSI => $"OC → Advanced DRAM Configuration → {leaf}",
            BoardVendor.Gigabyte => $"Tweaker → Advanced Memory Settings → {leaf}",
            BoardVendor.ASRock => $"OC Tweaker → DRAM Configuration → {leaf}",
            _ => $"BIOS → {leaf}"
        };
    }

    private static string VendorizeVoltagePath(SystemContext ctx, string leaf)
    {
        return ctx.BoardVendor switch
        {
            BoardVendor.ASUS => $"AI Tweaker → {leaf}",
            BoardVendor.MSI => $"OC → Voltage Setting → {leaf}",
            BoardVendor.Gigabyte => $"Tweaker → Advanced Voltage Settings → {leaf}",
            BoardVendor.ASRock => $"OC Tweaker → Voltage Configuration → {leaf}",
            _ => $"BIOS → {leaf}"
        };
    }

    private static void AddTiming(List<Recommendation> recs, string name, string value, string path, string reason, string level)
    {
        recs.Add(new Recommendation
        {
            Section = "MEMORY",
            Setting = $"RAM Calc: {name}",
            Value = value,
            Path = path,
            Reason = reason,
            WarningLevel = level
        });
    }

    private static string LevelFor(RamCalcConfig cfg)
    {
        return cfg.Goal switch
        {
            RamGoal.Safe => "Info",
            RamGoal.Fast => "Warning",
            RamGoal.Extreme => "Danger",
            _ => "Info"
        };
    }

    private static TimingSet GetBaseline(RamDieType die, int mts)
    {
        // Snap speed to nearest supported step.
        int s = new[] { 3200, 3600, 3800, 4000 }.OrderBy(x => Math.Abs(x - mts)).First();

        // Baselines are intentionally conservative for first-boot on typical AM4.
        // Subtimings beyond this are extremely kit/board dependent.
        return (die, s) switch
        {
            (RamDieType.Samsung_BDie, 3200) => new TimingSet(14, 14, 14, 14, 28, 42, 14, 4, 6, 16, 4, 12, 12, 288, 65535, 1.35),
            (RamDieType.Samsung_BDie, 3600) => new TimingSet(16, 16, 16, 16, 32, 48, 16, 4, 6, 16, 4, 12, 12, 312, 65535, 1.40),
            (RamDieType.Samsung_BDie, 3800) => new TimingSet(16, 16, 16, 16, 34, 50, 16, 4, 6, 16, 4, 12, 12, 320, 65535, 1.42),
            (RamDieType.Samsung_BDie, 4000) => new TimingSet(18, 18, 18, 18, 36, 54, 18, 4, 6, 16, 4, 12, 12, 340, 65535, 1.45),

            (RamDieType.Hynix_DJR, 3200) => new TimingSet(16, 18, 18, 18, 36, 54, 16, 4, 7, 20, 4, 12, 14, 360, 65535, 1.35),
            (RamDieType.Hynix_DJR, 3600) => new TimingSet(18, 20, 20, 20, 38, 58, 18, 4, 7, 20, 4, 12, 14, 380, 65535, 1.40),
            (RamDieType.Hynix_DJR, 3800) => new TimingSet(18, 21, 21, 21, 40, 61, 18, 4, 7, 20, 4, 12, 14, 400, 65535, 1.42),
            (RamDieType.Hynix_DJR, 4000) => new TimingSet(20, 22, 22, 22, 42, 64, 20, 4, 8, 24, 4, 12, 16, 420, 65535, 1.45),

            (RamDieType.Hynix_CJR, 3200) => new TimingSet(16, 18, 18, 18, 36, 54, 16, 4, 7, 20, 4, 12, 14, 380, 65535, 1.35),
            (RamDieType.Hynix_CJR, 3600) => new TimingSet(18, 20, 20, 20, 38, 58, 18, 4, 7, 20, 4, 12, 14, 420, 65535, 1.38),
            (RamDieType.Hynix_CJR, 3800) => new TimingSet(18, 21, 21, 21, 40, 61, 18, 4, 7, 20, 4, 12, 14, 440, 65535, 1.40),
            (RamDieType.Hynix_CJR, 4000) => new TimingSet(20, 22, 22, 22, 42, 64, 20, 4, 8, 24, 4, 12, 16, 460, 65535, 1.42),

            (RamDieType.Micron_EDie, 3200) => new TimingSet(16, 18, 18, 18, 36, 54, 16, 4, 7, 20, 4, 12, 14, 350, 65535, 1.35),
            (RamDieType.Micron_EDie, 3600) => new TimingSet(16, 19, 19, 19, 38, 57, 16, 4, 7, 20, 4, 12, 14, 360, 65535, 1.38),
            (RamDieType.Micron_EDie, 3800) => new TimingSet(18, 20, 20, 20, 40, 60, 18, 4, 7, 20, 4, 12, 14, 380, 65535, 1.40),
            (RamDieType.Micron_EDie, 4000) => new TimingSet(18, 22, 22, 22, 42, 64, 18, 4, 8, 24, 4, 12, 16, 400, 65535, 1.42),

            (RamDieType.Micron_RevB, 3200) => new TimingSet(16, 18, 18, 18, 36, 54, 16, 4, 7, 20, 4, 12, 14, 360, 65535, 1.35),
            (RamDieType.Micron_RevB, 3600) => new TimingSet(18, 20, 20, 20, 38, 58, 18, 4, 7, 20, 4, 12, 14, 390, 65535, 1.38),
            (RamDieType.Micron_RevB, 3800) => new TimingSet(18, 21, 21, 21, 40, 61, 18, 4, 7, 20, 4, 12, 14, 410, 65535, 1.40),
            (RamDieType.Micron_RevB, 4000) => new TimingSet(20, 22, 22, 22, 42, 64, 20, 4, 8, 24, 4, 12, 16, 430, 65535, 1.42),

            // If unknown die, produce a conservative "JEDEC-ish" baseline near XMP.
            _ => s switch
            {
                3200 => new TimingSet(16, 18, 18, 18, 36, 54, 16, 4, 7, 20, 4, 12, 14, 380, 65535, 1.35),
                3600 => new TimingSet(18, 20, 20, 20, 38, 58, 18, 4, 8, 24, 4, 12, 16, 420, 65535, 1.35),
                3800 => new TimingSet(18, 22, 22, 22, 40, 62, 18, 4, 8, 24, 4, 12, 16, 440, 65535, 1.38),
                _ => new TimingSet(20, 22, 22, 22, 42, 64, 20, 4, 8, 24, 4, 12, 16, 460, 65535, 1.40)
            }
        };
    }

    
    private static TimingSet GetBaselineDdr5(RamDieType die, int mts)
    {
        // Conservative DDR5 baselines. These are stability-first starting points, not guarantees.
        var nearest = new[] { 4800, 5200, 5600, 6000, 6200, 6400, 6600, 6800, 7000, 7200, 7400, 7600, 7800, 8000, 8200, 8400 }
            .OrderBy(x => Math.Abs(x - mts))
            .First();

        // Base CL by speed for Hynix-first (most common for higher DDR5 bins).
        int cl = nearest switch
        {
            <= 5200 => 40,
            <= 5600 => 36,
            <= 6000 => 34,
            <= 6400 => 32,
            <= 7200 => 34,
            <= 8000 => 36,
            _ => 38
        };

        // Vendor/die adjustments (keep it conservative).
        if (die is RamDieType.Samsung_BDie or RamDieType.Samsung_CDie or RamDieType.Samsung_DDie)
            cl += 4;
        else if (die is RamDieType.Micron_EDie or RamDieType.Micron_RevB or RamDieType.Micron_DDie)
            cl += 2;
        else if (die is RamDieType.Hynix_CJR or RamDieType.Hynix_DJR)
            cl += 2;

        int rcd = cl + 2;
        int rp = cl + 2;

        // Samsung often needs looser RCD/RP at the same CL.
        if (die is RamDieType.Samsung_BDie or RamDieType.Samsung_CDie or RamDieType.Samsung_DDie)
        {
            rcd += 2;
            rp += 2;
        }

        var tras = cl * 2 + 28;
        var trc = tras + rp;
        var cwl = Math.Max(24, cl - 2);

        // Secondaries: generic DDR5-safe defaults.
        var rrds = 8;
        var rrdl = 12;
        var faw = 32;
        var wtrs = 4;
        var wtrl = 12;
        var wr = 48;

        // tRFC depends on density; keep conservative.
        var trfc = 560;
        var trefi = 65535;

        // DRAM voltage: conservative bins.
        double vdram = nearest >= 7200 ? 1.40 : 1.35;

        return new TimingSet(
            tCL: cl, tRCDRD: rcd, tRCDWR: rcd, tRP: rp, tRAS: tras,
            tRC: trc, tCWL: cwl,
            tRRDS: rrds, tRRDL: rrdl, tFAW: faw,
            tWTRS: wtrs, tWTRL: wtrl, tWR: wr,
            tRFC: trfc,
            tREFI: trefi,
            vDram: vdram
        );
    }

    private static TimingSet ApplyGoalDdr5(TimingSet set, RamGoal goal)
    {
        // Goal tiers for DDR5 (starting points):
        // Safe      = stability-first / training-friendly
        // Balanced  = daily OC baseline
        // Fast      = tighter gaming baseline (requires validation)
        // Extreme   = aggressive (high risk; expect training limits)

        return goal switch
        {
            RamGoal.Safe => set,

            RamGoal.Balanced => set with
            {
                tCL = Math.Max(28, set.tCL - 2),
                tRCDRD = Math.Max(30, set.tRCDRD - 2),
                tRCDWR = Math.Max(30, set.tRCDWR - 2),
                tRP = Math.Max(30, set.tRP - 2),
                tCWL = Math.Max(24, set.tCWL - 1),
                tRRDL = Math.Max(8, set.tRRDL - 1),
                tFAW = Math.Max(24, set.tFAW - 4),
                tRFC = Math.Max(360, set.tRFC - 40),
                tREFI = Math.Min(65535, Math.Max(set.tREFI, 49152)),
                vDram = Math.Min(set.vDram + 0.03, 1.45)
            },

            RamGoal.Fast => set with
            {
                tCL = Math.Max(26, set.tCL - 4),
                tRCDRD = Math.Max(28, set.tRCDRD - 4),
                tRCDWR = Math.Max(28, set.tRCDWR - 4),
                tRP = Math.Max(28, set.tRP - 4),
                tCWL = Math.Max(22, set.tCWL - 2),
                tRRDS = Math.Max(4, set.tRRDS - 1),
                tRRDL = Math.Max(6, set.tRRDL - 2),
                tFAW = Math.Max(20, set.tFAW - 8),
                tWTRS = Math.Max(3, set.tWTRS - 1),
                tWTRL = Math.Max(8, set.tWTRL - 2),
                tRFC = Math.Max(320, set.tRFC - 80),
                tREFI = Math.Min(65535, Math.Max(set.tREFI, 60000)),
                vDram = Math.Min(set.vDram + 0.07, 1.50)
            },

            RamGoal.Extreme => set with
            {
                tCL = Math.Max(24, set.tCL - 6),
                tRCDRD = Math.Max(26, set.tRCDRD - 6),
                tRCDWR = Math.Max(26, set.tRCDWR - 6),
                tRP = Math.Max(26, set.tRP - 6),
                tCWL = Math.Max(20, set.tCWL - 3),
                tRRDS = Math.Max(4, set.tRRDS - 1),
                tRRDL = Math.Max(6, set.tRRDL - 2),
                tFAW = Math.Max(16, set.tFAW - 10),
                tWTRS = Math.Max(3, set.tWTRS - 1),
                tWTRL = Math.Max(8, set.tWTRL - 2),
                tWR = Math.Max(24, set.tWR - 4),
                tRFC = Math.Max(280, set.tRFC - 120),
                tREFI = 65535,
                vDram = Math.Min(set.vDram + 0.10, 1.55)
            },

            _ => set
        };
    }



private static TimingSet ApplyTopologyModifiers(TimingSet s, RamCalcConfig cfg)
    {
        // More ranks / DIMMs increase IMC load -> loosen slightly.
        int loosen = 0;
        if (cfg.DimmCount >= 4) loosen++;
        if (cfg.Rank == RamRank.DualRank) loosen++;

        if (loosen == 0) return s;

        return s with
        {
            tRCDRD = s.tRCDRD + loosen,
            tRCDWR = s.tRCDWR + loosen,
            tRP = s.tRP + loosen,
            tRAS = s.tRAS + loosen * 2,
            tRC = s.tRC + loosen * 2,
            tRFC = s.tRFC + loosen * 10,
            vDram = Math.Min(s.vDram + 0.02 * loosen, 1.50)
        };
    }

    private static TimingSet ApplyGoal(TimingSet s, RamGoal goal)
    {
        return goal switch
        {
            RamGoal.Safe => s,
            RamGoal.Fast => s with
            {
                tCL = Math.Max(12, s.tCL - 1),
                tRCDRD = Math.Max(12, s.tRCDRD - 1),
                tRCDWR = Math.Max(12, s.tRCDWR - 1),
                tRP = Math.Max(12, s.tRP - 1),
                tRFC = Math.Max(260, s.tRFC - 20),
                vDram = s.vDram + 0.03
            },
            RamGoal.Extreme => s with
            {
                tCL = Math.Max(12, s.tCL - 2),
                tRCDRD = Math.Max(12, s.tRCDRD - 2),
                tRCDWR = Math.Max(12, s.tRCDWR - 2),
                tRP = Math.Max(12, s.tRP - 2),
                tRRDS = Math.Max(4, s.tRRDS),
                tRRDL = Math.Max(6, s.tRRDL - 1),
                tFAW = Math.Max(16, s.tFAW - 4),
                tRFC = Math.Max(240, s.tRFC - 40),
                vDram = s.vDram + 0.06
            },
            _ => s
        };
    }

    private static TimingSet ApplyVoltageCap(TimingSet s, double cap, RamGoal goal, RamDieType die)
    {
        // If cap is lower than recommendation, loosen primaries and tRFC slightly.
        if (s.vDram <= cap) return s;

        double delta = s.vDram - cap;
        int bump = delta >= 0.08 ? 2 : 1;

        return s with
        {
            vDram = cap,
            tCL = s.tCL + bump,
            tRCDRD = s.tRCDRD + bump,
            tRCDWR = s.tRCDWR + bump,
            tRP = s.tRP + bump,
            tRFC = s.tRFC + bump * 20
        };
    }
}
