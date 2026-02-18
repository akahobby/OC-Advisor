// OC ADVISOR — OUTPUT ONLY (NO AUTO-APPLY)
// Educated guesses (tighter heuristics) based on detected components.
// - Competitive default
// - Auto-grouping + ordered sections
// - No NIC/network
// - Ryzen DDR4 rule: FCLK = RAM MT/s ÷ 2 (cap 2000)
//
// Notes:
// - These are *educated starting points*, not guaranteed “your exact silicon” settings.
// - Always step/test; if instability: back off GPU memory first → GPU core → CO → RAM/FCLK.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OcAdvisor;

internal static class OcAdvisorBackend
{
    internal static (SystemContext ctx, List<Recommendation> recs, string text, string json) Generate(TuningProfile profile)
    {
        var ctx = HardwareDetector.Detect(profile);
        var recs = RuleEngine.Build(ctx);
        var text = ReportBuilder.BuildText(ctx, recs);
        var json = ReportBuilder.BuildJson(ctx, recs);
        return (ctx, recs, text, json);
    }
}

#if CLI
static class Program
{
    static int Main(string[] args)
    {
        try
        {
            var opts = CliOptions.Parse(args);
            var ctx = HardwareDetector.Detect(opts.Profile);
            var recs = RuleEngine.Build(ctx);

            Directory.CreateDirectory(opts.OutDir);
            var baseName = $"OC_Recommendations_{Environment.MachineName}";
            var txtPath = Path.Combine(opts.OutDir, baseName + ".txt");
            var jsonPath = Path.Combine(opts.OutDir, baseName + ".json");

            if (opts.Format is OutputFormat.Txt or OutputFormat.Both)
                File.WriteAllText(txtPath, ReportBuilder.BuildText(ctx, recs), new UTF8Encoding(false));

            if (opts.Format is OutputFormat.Json or OutputFormat.Both)
                File.WriteAllText(jsonPath, ReportBuilder.BuildJson(ctx, recs), new UTF8Encoding(false));

            Console.WriteLine("Generated:");
            if (opts.Format is OutputFormat.Txt or OutputFormat.Both) Console.WriteLine(txtPath);
            if (opts.Format is OutputFormat.Json or OutputFormat.Both) Console.WriteLine(jsonPath);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error:");
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }
}
#endif


#region CLI

enum OutputFormat { Txt, Json, Both }
enum TuningProfile { DailyGaming, Competitive }

sealed class CliOptions
{
    public TuningProfile Profile { get; init; } = TuningProfile.Competitive;
    public OutputFormat Format { get; init; } = OutputFormat.Both;
    public string OutDir { get; init; } = AppContext.BaseDirectory;

    public static CliOptions Parse(string[] args)
    {
        var profile = TuningProfile.Competitive;
        var format = OutputFormat.Both;
        var outDir = AppContext.BaseDirectory;

        for (int i = 0; i < args.Length; i++)
        {
            var a = args[i].Trim();

            if (a.Equals("--help", StringComparison.OrdinalIgnoreCase) || a.Equals("-h", StringComparison.OrdinalIgnoreCase))
            {
                PrintHelp();
                Environment.Exit(0);
            }

            if (a.Equals("--profile", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                var v = args[++i].Trim().ToLowerInvariant();
                profile = v switch
                {
                    "daily" or "dailygaming" => TuningProfile.DailyGaming,
                    "competitive" or "comp" => TuningProfile.Competitive,
                    _ => throw new ArgumentException("Invalid --profile. Use: daily | competitive")
                };
                continue;
            }

            if (a.Equals("--format", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                var v = args[++i].Trim().ToLowerInvariant();
                format = v switch
                {
                    "txt" => OutputFormat.Txt,
                    "json" => OutputFormat.Json,
                    "both" => OutputFormat.Both,
                    _ => throw new ArgumentException("Invalid --format. Use: txt | json | both")
                };
                continue;
            }

            if (a.Equals("--outdir", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                outDir = args[++i].Trim().Trim('"');
                continue;
            }
        }

        return new CliOptions { Profile = profile, Format = format, OutDir = outDir };
    }

    static void PrintHelp()
    {
        Console.WriteLine(@"
OC Advisor (no auto-apply)

Usage:
  OcAdvisor.exe [--profile daily|competitive] [--format txt|json|both] [--outdir PATH]

Defaults:
  --profile competitive
  --format  both
  --outdir  (exe folder)
");
    }
}

#endregion

#region Models

enum CpuVendor { Unknown, AMD, Intel }
enum GpuVendor { Unknown, NVIDIA, AMD }
enum MemoryType { Unknown, DDR4, DDR5 }
enum DiskKind { Unknown, NVMe, SATA_SSD, SATA_HDD, USB, Other }
enum BoardVendor { Unknown, ASUS, MSI, Gigabyte, ASRock, Other }

enum AmdSocketEra { Unknown, AM4, AM5 }
enum AmdTdpBucket { Unknown, W65, W105, W120, W170 } // heuristic buckets

sealed class SystemContext
{
    public TuningProfile Profile;

    public string OsCaption = "";
    public string OsVersion = "";
    public string ActivePowerPlan = "";

    public CpuVendor CpuVendor = CpuVendor.Unknown;
    public string CpuName = "";
    public CpuClass CpuClass = new();

    public string BoardManufacturer = "";
    public string BoardProduct = "";
    public BoardVendor BoardVendor = BoardVendor.Unknown;
    public bool IsAsus => BoardManufacturer.Contains("ASUS", StringComparison.OrdinalIgnoreCase);

    public string BiosVendor = "";
    public string BiosVersion = "";
    public string BiosDate = "";

    public MemoryType MemoryType = MemoryType.Unknown;
    public double TotalRamGb;

    public int? EffectiveMemMts;
    public int? TargetFclkMHz; // stability-first recommendation
    public int? IdealFclkMHz;  // mathematical 1:1 target (MCLK)
    public string FabricMode = ""; // 1:1, 1:2, etc (best-effort)

    // Best-effort RAM module hints for the RAM timing calculator UI.
    public int? DetectedDimmCount;
    // "SingleRank" | "DualRank" (best-effort; verify in SPD tools)
    public string DetectedRankHint = "";
    public readonly List<string> MemoryPartNumbers = new();

    public DisplayInfo Display = new();

    public readonly List<GpuInfo> Gpus = new();
    public readonly List<DiskInfo> Disks = new();

    public GpuVendor PrimaryGpuVendor => Gpus.Count > 0 ? Gpus[0].Vendor : GpuVendor.Unknown;
    public string PrimaryGpuName => Gpus.Count > 0 ? Gpus[0].Name : "";
    public GpuClass PrimaryGpuClass => GpuClassifier.Classify(PrimaryGpuName);
}

sealed class CpuClass
{
    public bool IsX3D;
    public int? AmdModel;         // e.g. 5800, 7800
    public int? RyzenSeries;      // 3000, 5000, 7000
    public AmdSocketEra AmdEra;   // AM4/AM5 (best effort)
    public AmdTdpBucket AmdTdp;   // 65/105/120/170 (best effort)

    public int? IntelGen;         // 12/13/14 best effort
}

sealed class GpuInfo
{
    public string Name = "";
    public string DriverVersion = "";
    public GpuVendor Vendor = GpuVendor.Unknown;
}

sealed class GpuClass
{
    public GpuVendor Vendor;
    public int? NvidiaRtxGen;     // 20/30/40/50
    public int? NvidiaModel;      // e.g. 5070/4070/3080
    public int? NvidiaTier;       // 60/70/80/90 (or 50)
    public bool IsLaptop;
    public bool IsFactoryOC;      // "OC", "Gaming OC", etc (string heuristic)
}

sealed class DiskInfo
{
    public string Model = "";
    public string InterfaceType = "";
    public string MediaType = "";
    public DiskKind Kind = DiskKind.Unknown;
    public ulong SizeBytes;
}

sealed class DisplayInfo
{
    public int Width;
    public int Height;
    public int RefreshHz;
    public bool HasDisplay;
}

#endregion

#region Recommendations

sealed class Recommendation
{
    // WPF data-binding expects public properties (not fields)
    public string Section { get; set; } = "";
    public string Title { get; set; } = "";
    public string Path { get; set; } = "";
    public string Setting { get; set; } = "";
    public string Value { get; set; } = "";
    public string Reason { get; set; } = "";
    public string WarningLevel { get; set; } = "None"; // None | Info | Warning | Danger
}

static class RuleEngine
{
    // Vendor-aware BIOS path helper (kept here so all rule builders can call it)
    private static string VendorizePath(
        BoardVendor vendor,
        string asus,
        string msi,
        string gigabyte,
        string asrock,
        string fallback)
    {
        return vendor switch
        {
            BoardVendor.ASUS => asus,
            BoardVendor.MSI => msi,
            BoardVendor.Gigabyte => gigabyte,
            BoardVendor.ASRock => asrock,
            _ => fallback
        };
    }

    public static readonly string[] SectionOrder =
    {
        "CPU",
        "MEMORY",
        "GPU",
        "BIOS",
        "DISPLAY",
        "STORAGE",
        "TESTING"
    };

    public static List<Recommendation> Build(SystemContext ctx)
    {
        var recs = new List<Recommendation>();

        // CPU
        if (ctx.CpuVendor == CpuVendor.AMD) recs.AddRange(BuildAmdCpu(ctx));
        else if (ctx.CpuVendor == CpuVendor.Intel) recs.AddRange(BuildIntelCpu(ctx));

        // MEMORY
        recs.AddRange(BuildMemory(ctx));

        // GPU
        if (ctx.PrimaryGpuVendor != GpuVendor.Unknown) recs.AddRange(BuildGpu(ctx));

        // BIOS
        recs.AddRange(BuildBios(ctx));

        // DISPLAY
        recs.AddRange(BuildDisplay(ctx));

        // STORAGE
        recs.AddRange(BuildStorage(ctx));

        // TESTING
        recs.AddRange(BuildTesting(ctx));

        // Clean + de-dupe
        recs = recs
            .Where(r => !string.IsNullOrWhiteSpace(r.Section) && !string.IsNullOrWhiteSpace(r.Setting) && !string.IsNullOrWhiteSpace(r.Value))
            .GroupBy(r => $"{r.Section}|{r.Setting}|{r.Value}", StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        // Order by section order
        var order = SectionOrder.Select((s, i) => (s, i)).ToDictionary(x => x.s, x => x.i, StringComparer.OrdinalIgnoreCase);
        recs.Sort((a, b) =>
        {
            var oa = order.TryGetValue(a.Section, out var ia) ? ia : 999;
            var ob = order.TryGetValue(b.Section, out var ib) ? ib : 999;
            var c = oa.CompareTo(ob);
            if (c != 0) return c;
            return string.Compare(a.Setting, b.Setting, StringComparison.OrdinalIgnoreCase);
        });

        return recs;
    }

    static IEnumerable<Recommendation> BuildAmdCpu(SystemContext ctx)
    {
        var list = new List<Recommendation>();
        var cc = ctx.CpuClass;

        var pbo = AmdHeuristics.GuessPbo(cc, ctx.Profile);
        var co = AmdHeuristics.GuessCurveOptimizer(cc, ctx.Profile);
        var boost = AmdHeuristics.GuessBoostOverrideMHz(cc, ctx.Profile);
        var scalar = AmdHeuristics.GuessScalar(cc, ctx.Profile);
        var thermal = AmdHeuristics.GuessThermalLimitC(cc, ctx.Profile);

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "PBO",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "Precision Boost Overdrive",
            Value = "Advanced",
            Reason = "Ryzen performance tuning is best done through PBO + CO (not fixed all-core OC)."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "PBO Limits (educated guess)",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "PPT / TDC / EDC",
            Value = $"{pbo.PptW}W / {pbo.TdcA}A / {pbo.EdcA}A",
            Reason = $"Bucket: {cc.AmdEra}/{cc.AmdTdp} {(cc.IsX3D ? "X3D" : "")}. If temps spike or instability: lower EDC first, then PPT."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Curve Optimizer (educated guess)",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → PBO → Curve Optimizer" : "BIOS → AMD Overclocking → PBO → Curve Optimizer",
            Setting = "CO",
            Value = $"All-core Negative -{co.Start} start (aim -{co.Typical} if stable)",
            Reason = cc.IsX3D
                ? "X3D: keep CO more conservative to avoid WHEA/game instability."
                : "CO often gives the best gaming gains per watt."
        });


        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "CO refinement",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → PBO → Curve Optimizer" : "BIOS → AMD Overclocking → PBO → Curve Optimizer",
            Setting = "Per-core hint",
            Value = "If all-core CO crashes: keep a weaker all-core value OR tune per-core (best cores more negative, weakest cores less negative)",
            Reason = "Ryzen boosts per-core; the weakest core often sets stability. Per-core CO can keep boost without random game/WHEA errors.",
            WarningLevel = "Info"
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Transient stability",
            Path = "Testing",
            Setting = "If you crash during loads/alt-tab",
            Value = "First reduce PBO limits/boost override or LLC aggressiveness before abandoning CO",
            Reason = "Many “CO unstable” reports are actually transient voltage/boost overshoot issues.",
            WarningLevel = "Info"
        });

;

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Boost Override",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "Max CPU Boost Clock Override",
            Value = boost == 0 ? "0 (leave default)" : $"+{boost} MHz (if temps allow)",
            Reason = boost == 0 ? "Conservative default for stability (and many X3D cases)." : "If unstable: drop boost override before touching voltages."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Scalar",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "PBO Scalar",
            Value = scalar,
            Reason = "Higher scalar can increase voltage/heat for small gains; competitive favors conservative."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Thermal Limit",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "Thermal Throttle Limit",
            Value = thermal,
            Reason = "Competitive stability: avoid clock oscillation from heat and power swings."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "LLC",
            Path = ctx.IsAsus ? "AI Tweaker / Digi+ VRM" : "BIOS → OC / VRM",
            Setting = "CPU LLC",
            Value = "Middle level (avoid extremes)",
            Reason = "Too high can spike voltage; too low can droop/crash under transient boost."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Scheduling",
            Path = ctx.IsAsus ? "Advanced → AMD CBS" : "BIOS → AMD CBS",
            Setting = "CPPC + Preferred Cores",
            Value = "Enabled",
            Reason = "Helps scheduler pick best boosting cores."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "C-States",
            Path = ctx.IsAsus ? "Advanced → AMD CBS" : "BIOS → AMD CBS",
            Setting = "Global C-State Control",
            Value = ctx.Profile == TuningProfile.Competitive ? "Disabled" : "Auto",
            Reason = ctx.Profile == TuningProfile.Competitive ? "Latency consistency focus (higher idle power)." : "Auto is fine for daily."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Idle Control",
            Path = ctx.IsAsus ? "Advanced → AMD CBS" : "BIOS → AMD CBS",
            Setting = "Power Supply Idle Control",
            Value = "Typical Current Idle",
            Reason = "Avoids rare low-power idle instability on some Ryzen systems."
        });

        // AM5-specific: EXPO + SoC voltage caution without giving exact volts (stability-only note)
        if (cc.AmdEra == AmdSocketEra.AM5)
        {
            list.Add(new Recommendation
            {
                Section = "CPU",
                Title = "AM5 note",
                Path = ctx.IsAsus ? "AI Tweaker / AMD Overclocking" : "BIOS → OC",
                Setting = "If EXPO/FCLK-related instability",
                Value = "Prefer reducing memory speed one step before raising voltages; keep changes minimal for competitive stability.",
                Reason = "Stability-first approach; avoids chasing multiple voltage rails."
            });
        }

        return list;
    }

    static IEnumerable<Recommendation> BuildIntelCpu(SystemContext ctx)
    {
        var list = new List<Recommendation>();
        var gen = ctx.CpuClass.IntelGen ?? 0;

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Power Limits",
            Path = "BIOS → CPU Power Management",
            Setting = "PL1 / PL2 / Tau",
            Value = gen >= 12 ? "Set to what your cooler sustains (avoid unlimited if it throttles)" : "Set to what your cooler sustains",
            Reason = "Sustained clocks matter for consistent frametimes."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Voltage Mode",
            Path = "BIOS → CPU Voltage",
            Setting = "Core Voltage",
            Value = "Adaptive",
            Reason = "Keeps boost behavior intact and avoids unnecessary heat."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "LLC",
            Path = "BIOS → OC / VRM",
            Setting = "LLC",
            Value = "Middle level (avoid extremes)",
            Reason = "Balance droop vs spikes under transient loads."
        });


        // Competitive-focused Intel extras (safe defaults)
        var vendorCpuFeatures = VendorizePath(
            ctx.BoardVendor,
            asus: "AI Tweaker → CPU Power Management",
            msi: "OC → CPU Features",
            gigabyte: "Tweaker → Advanced CPU Settings",
            asrock: "OC Tweaker → CPU Configuration",
            fallback: "BIOS → CPU Features"
        );

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Thread Director",
            Path = vendorCpuFeatures,
            Setting = "Intel Thread Director",
            Value = "Enabled (Windows 11)",
            Reason = "Helps scheduler place threads on P-cores/E-cores appropriately for smoother frametimes."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Speed Shift",
            Path = vendorCpuFeatures,
            Setting = "Intel Speed Shift (HWP)",
            Value = "Enabled",
            Reason = "Improves responsiveness to load changes (helps 1% lows) compared to legacy P-states."
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "E-cores",
            Path = vendorCpuFeatures,
            Setting = "E-cores",
            Value = "Enabled (disable only if a specific game anti-cheat/scheduler issue)",
            Reason = "Most modern titles + Windows 11 benefit from keeping E-cores; disabling is a last resort troubleshooting step.",
            WarningLevel = "Info"
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Boost features",
            Path = vendorCpuFeatures,
            Setting = "TVB / ABT (if present)",
            Value = "Enabled (if temps allow)",
            Reason = "Extra opportunistic boost can help peak FPS; only useful if cooling keeps temps under control."
        });

        var vendorVoltage = VendorizePath(
            ctx.BoardVendor,
            asus: "AI Tweaker → Internal CPU Power Management / Voltage",
            msi: "OC → Voltage Settings",
            gigabyte: "Tweaker → Advanced Voltage Settings",
            asrock: "OC Tweaker → Voltage Configuration",
            fallback: "BIOS → Voltage"
        );

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Undervolt (optional)",
            Path = vendorVoltage,
            Setting = "Adaptive Voltage Offset",
            Value = "Start -0.050 V (if stable), then validate",
            Reason = "A small negative offset can reduce heat and improve sustained boost; some BIOSes lock undervolt on newer microcode.",
            WarningLevel = "Warning"
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "Protection",
            Path = vendorVoltage,
            Setting = "CEP / Current Excursion Protection",
            Value = "Enabled (recommended)",
            Reason = "Prevents unstable transient behavior; disabling can cause clock/voltage spikes and WHEA-like instability.",
            WarningLevel = "Info"
        });

        list.Add(new Recommendation
        {
            Section = "CPU",
            Title = "SVID behavior",
            Path = vendorVoltage,
            Setting = "SVID Behavior",
            Value = "Typical/Auto (avoid aggressive 'Best case' unless tuning)",
            Reason = "Overly optimistic SVID can undervolt unexpectedly under transient load; stability-first for competitive play."
        });

        return list;
    }

    static IEnumerable<Recommendation> BuildMemory(SystemContext ctx)
    {
        var list = new List<Recommendation>();

        list.Add(new Recommendation
        {
            Section = "MEMORY",
            Title = "Baseline",
            Path = "AI Tweaker / OC",
            Setting = "XMP / EXPO / DOCP",
            Value = "Enabled (baseline)",
            Reason = "Start at rated kit profile before experimenting."
        });

        
        if (ctx.CpuVendor == CpuVendor.AMD && ctx.MemoryType == MemoryType.DDR4 && ctx.EffectiveMemMts.HasValue)
        {
            var desired = ctx.EffectiveMemMts.Value / 2;
            var target = ctx.TargetFclkMHz ?? desired;

            var vendorPathFabric = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker → AMD Overclocking",
                msi: "OC → Advanced CPU Configuration",
                gigabyte: "Tweaker → AMD Overclocking",
                asrock: "OC Tweaker → AMD Overclocking",
                fallback: "BIOS → Memory/Fabric"
            );

            var fclkValue = (target != desired)
                ? $"{target} MHz (try {desired} if stable)"
                : $"{target} MHz";

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "Ryzen DDR4 Fabric",
                Path = vendorPathFabric,
                Setting = "FCLK",
                Value = fclkValue,
                Reason = (target != desired)
                    ? "DDR4-4000 implies 2000 MHz fabric (1:1), but Zen3 commonly tops out ~1900 stable. Presenting a safe default while showing the ideal 1:1 target."
                    : "Match fabric clock to memory clock for lowest latency (1:1).",
                WarningLevel = (target != desired) ? "Info" : "None"
            });

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "Fabric Mode",
                Path = vendorPathFabric,
                Setting = "FCLK : MCLK",
                Value = string.IsNullOrWhiteSpace(ctx.FabricMode) ? "1:1 (preferred)" : ctx.FabricMode,
                Reason = "If fabric desyncs (1:2), latency typically increases and competitive frametimes can worsen.",
                WarningLevel = (target != desired) ? "Warning" : "None"
            });

            // WHEA risk heuristic (Zen3)
            if ((ctx.CpuClass.RyzenSeries ?? 0) >= 5000 && (ctx.CpuClass.RyzenSeries ?? 0) < 7000)
            {
                var whea = (desired >= 2000) ? "High (silicon lottery — watch WHEA 19)" :
                           (target >= 1900) ? "Moderate (monitor Event Viewer for WHEA 19)" :
                           "Low";
                list.Add(new Recommendation
                {
                    Section = "MEMORY",
                    Title = "WHEA risk",
                    Path = "Windows → Event Viewer",
                    Setting = "WHEA 19 (fabric/memory)",
                    Value = whea,
                    Reason = "On Zen3, pushing fabric (especially 2000 MHz) can produce silent corrected errors before visible crashes.",
                    WarningLevel = (desired >= 2000) ? "Danger" : (target >= 1900 ? "Warning" : "None")
                });
            }

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "Stability levers",
                Path = VendorizePath(
                    ctx.BoardVendor,
                    asus: "AI Tweaker → DRAM Timing Control",
                    msi: "OC → DRAM Setting",
                    gigabyte: "Tweaker → Advanced Memory Settings",
                    asrock: "OC Tweaker → DRAM Configuration",
                    fallback: "BIOS → DRAM"
                ),
                Setting = "GDM / CR",
                Value = "If unstable: Gear Down Mode = Enabled, Command Rate = Auto/2T first",
                Reason = "Stability-first levers without guessing primary/secondary timings.",
                WarningLevel = "None"
            });

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "UCLK",
                Path = VendorizePath(
                    ctx.BoardVendor,
                    asus: "AI Tweaker → AMD Overclocking",
                    msi: "OC → Advanced DRAM Configuration",
                    gigabyte: "Tweaker → AMD Overclocking",
                    asrock: "OC Tweaker → AMD Overclocking",
                    fallback: "BIOS → Memory/Fabric"
                ),
                Setting = "UCLK mode",
                Value = "UCLK = MCLK (1:1) preferred (avoid half-rate unless required)",
                Reason = "Half-rate UCLK can add latency even when frequency looks higher on paper.",
                WarningLevel = "Info"
            });

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "Voltages",
                Path = VendorizePath(
                    ctx.BoardVendor,
                    asus: "AI Tweaker → Voltage",
                    msi: "OC → Voltage Setting",
                    gigabyte: "Tweaker → Advanced Voltage Settings",
                    asrock: "OC Tweaker → Voltage Configuration",
                    fallback: "BIOS → Voltage"
                ),
                Setting = "SoC (guardrails)",
                Value = "Try to stay ≤ 1.10V; avoid ≥ 1.15V for daily use",
                Reason = "SoC can help fabric/memory stability, but high SoC voltage has diminishing returns and increases long-term risk.",
                WarningLevel = "Warning"
            });

            list.Add(new Recommendation
            {
                Section = "MEMORY",
                Title = "Boot training",
                Path = VendorizePath(
                    ctx.BoardVendor,
                    asus: "Advanced → AMD CBS",
                    msi: "Settings → Advanced → AMD CBS",
                    gigabyte: "Settings → AMD CBS",
                    asrock: "Advanced → AMD CBS",
                    fallback: "BIOS → Advanced"
                ),
                Setting = "Memory Context Restore (MCR)",
                Value = "Enable for faster boots; disable if cold-boot training fails",
                Reason = "MCR reduces memory retraining time but can expose borderline RAM/FCLK stability on cold boot.",
                WarningLevel = "Info"
            });
        }

        return list;
    }

	    static IEnumerable<Recommendation> BuildGpu(SystemContext ctx)
	    {
	        var list = new List<Recommendation>();
	        if (ctx.Gpus.Count == 0) return list;
	        var g = ctx.Gpus[0];

	        if (g.Vendor == GpuVendor.NVIDIA)
	        {
	            var ranges = NvidiaHeuristics.GuessRanges(ctx.PrimaryGpuClass, ctx.Profile);
	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "Thermals",
	                Path = "MSI Afterburner",
	                Setting = "Fan curve",
	                Value = "Use a curve that holds stable temps under load",
	                Reason = "Stable temps = stable clocks and fewer frametime spikes.",
	                WarningLevel = "None"
	            });

	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "Order + safety",
	                Path = "MSI Afterburner",
	                Setting = "OC steps",
	                Value = "OC memory in small steps; test in-game; then core",
	                Reason = "VRAM OC is the most common cause of random driver resets/timeouts.",
	                WarningLevel = "Warning"
	            });

	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "Power / Core / Memory",
	                Path = "MSI Afterburner",
	                Setting = "Power / Core / Memory",
	                Value = $"Power: {ranges.Power} | Core: {ranges.Core} | Memory: {ranges.Mem}",
	                Reason = "Start in the lower half of the range and step up; validate in your real games.",
	                WarningLevel = "Info"
	            });

	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "Undervolt target",
	                Path = "MSI Afterburner (Curve Editor)",
	                Setting = "Undervolt",
	                Value = ranges.Undervolt,
	                Reason = "Undervolting often improves stability by reducing heat and transient power swings while keeping boost steady.",
	                WarningLevel = "Info"
	            });
	        }
	        else if (g.Vendor == GpuVendor.AMD)
	        {
	            var ranges = AmdGpuHeuristics.GuessRanges(ctx.PrimaryGpuClass, ctx.Profile);
	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "AMD Radeon OC (educated guess)",
	                Path = "AMD Adrenalin",
	                Setting = "Power / Core / VRAM",
	                Value = $"Power: {ranges.Power} | Core: {ranges.Core} | VRAM: {ranges.Mem}",
	                Reason = "Small steps reduce driver timeouts; validate in real games.",
	                WarningLevel = "Info"
	            });

	            list.Add(new Recommendation
	            {
	                Section = "GPU",
	                Title = "Thermals",
	                Path = "AMD Adrenalin",
	                Setting = "Fan curve",
	                Value = "Hold stable temps under load; avoid edge-of-stability VRAM OC",
	                Reason = "Reduces crashes and frametime spikes.",
	                WarningLevel = "None"
	            });
	        }

	        return list;
	    }

    
static IEnumerable<Recommendation> BuildBios(SystemContext ctx)
    {
        var list = new List<Recommendation>();

        var vendorPci = VendorizePath(
            ctx.BoardVendor,
            asus: "Advanced → PCI Subsystem Settings",
            msi: "Settings → Advanced → PCI Subsystem Settings",
            gigabyte: "Settings → IO Ports",
            asrock: "Advanced → Chipset Configuration",
            fallback: "Advanced → PCI Subsystem Settings"
        );

        var vendorOc = VendorizePath(
            ctx.BoardVendor,
            asus: "AI Tweaker / Extreme Tweaker",
            msi: "OC",
            gigabyte: "Tweaker",
            asrock: "OC Tweaker",
            fallback: "BIOS → OC"
        );

        // UEFI baseline
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "UEFI baseline",
            Path = "Boot",
            Setting = "CSM",
            Value = "Disabled",
            Reason = "Modern baseline; often required for ReBAR.",
            WarningLevel = "None"
        });

        // ReBAR / 4G
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "GPU BAR",
            Path = vendorPci,
            Setting = "Above 4G Decoding + ReBAR",
            Value = "Enabled (if supported)",
            Reason = "Can improve performance in supported games; if you see worse frametime consistency in older titles, test with ReBAR off.",
            WarningLevel = "Info"
        });

        // PCIe Link speed
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "PCIe Link Speed",
            Path = vendorPci,
            Setting = "PCIe Link Speed (GPU slot)",
            Value = "Auto (force one gen lower only if instability/black screens)",
            Reason = "Auto is safest; forcing can expose signal issues.",
            WarningLevel = "None"
        });

        // Spread Spectrum
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "Clock stability",
            Path = vendorOc,
            Setting = "Spread Spectrum (CPU/PCIe)",
            Value = "Disabled",
            Reason = "Spread Spectrum modulates clocks for EMI; disabling can improve latency consistency.",
            WarningLevel = "Info"
        });

        // PCIe ASPM
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "PCIe latency",
            Path = vendorPci,
            Setting = "ASPM",
            Value = "Disabled (optional for competitive latency)",
            Reason = "Disabling ASPM can reduce latency at the cost of higher idle power.",
            WarningLevel = "Info"
        });

        if (ctx.CpuVendor == CpuVendor.AMD)
        {


        

        if (ctx.CpuVendor == CpuVendor.Intel)
        {
            var vendorIntelCpu = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker → CPU Power Management",
                msi: "OC → CPU Features",
                gigabyte: "Tweaker → Advanced CPU Settings",
                asrock: "OC Tweaker → CPU Configuration",
                fallback: "BIOS → CPU Features"
            );

            list.Add(new Recommendation
            {
                Section = "BIOS",
                Title = "Scheduler support",
                Path = vendorIntelCpu,
                Setting = "Intel Thread Director",
                Value = "Enabled",
                Reason = "Improves core assignment on hybrid CPUs (P/E cores) to reduce stutter and improve 1% lows."
            });

            list.Add(new Recommendation
            {
                Section = "BIOS",
                Title = "Latency",
                Path = vendorIntelCpu,
                Setting = "CPU C-States / Package C-State Limit",
                Value = "Auto (try disabling deeper package C-states if you get micro-stutter)",
                Reason = "Deeper idle states can add wake latency; only change if you notice stutter/latency spikes.",
                WarningLevel = "Info"
            });

            var vendorIntelOc = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker / Extreme Tweaker",
                msi: "OC",
                gigabyte: "Tweaker",
                asrock: "OC Tweaker",
                fallback: "BIOS → OC"
            );

            list.Add(new Recommendation
            {
                Section = "BIOS",
                Title = "Multi-core enhancement",
                Path = vendorIntelOc,
                Setting = "MCE / Enhanced Turbo",
                Value = "Disabled (recommended) or Enabled only if temps are controlled",
                Reason = "MCE can raise voltage/heat for small gains; competitive stability favors predictable thermals.",
                WarningLevel = "Info"
            });

            var vendorIntelVoltage = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker → Internal CPU Power Management / Voltage",
                msi: "OC → Voltage Settings",
                gigabyte: "Tweaker → Advanced Voltage Settings",
                asrock: "OC Tweaker → Voltage Configuration",
                fallback: "BIOS → Voltage"
            );

            list.Add(new Recommendation
            {
                Section = "BIOS",
                Title = "Protection",
                Path = vendorIntelVoltage,
                Setting = "CEP (Current Excursion Protection)",
                Value = "Enabled (recommended)",
                Reason = "Disabling CEP can increase transient overshoot and instability; keep enabled unless you’re doing manual tuning.",
                WarningLevel = "Info"
            });

            list.Add(new Recommendation
            {
                Section = "BIOS",
                Title = "Voltage behavior",
                Path = vendorIntelVoltage,
                Setting = "SVID Behavior",
                Value = "Auto/Typical (avoid aggressive presets unless tuning)",
                Reason = "Aggressive SVID can cause unexpected voltage behavior under transient load."
            });
        }

        // DF C-States (if exposed)
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "Fabric latency",
            Path = VendorizePath(
                ctx.BoardVendor,
                asus: "Advanced → AMD CBS → DF Common Options",
                msi: "Settings → Advanced → AMD CBS → DF Common Options",
                gigabyte: "Settings → AMD CBS → DF Common Options",
                asrock: "Advanced → AMD CBS → DF Common Options",
                fallback: "Advanced → AMD CBS"
            ),
            Setting = "DF C-States",
            Value = "Disabled (if available)",
            Reason = "Can reduce fabric wake latency/micro-stutter in latency-sensitive titles.",
            WarningLevel = "Info"
        });

        
        }
// VRM tuning (LLC + phase)
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "VRM stability",
            Path = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker → Digi+ VRM",
                msi: "OC → DigitALL Power",
                gigabyte: "Tweaker → Advanced Voltage Settings",
                asrock: "OC Tweaker → Voltage Configuration",
                fallback: "BIOS → VRM"
            ),
            Setting = "LLC (CPU)",
            Value = ctx.BoardVendor switch
            {
                BoardVendor.ASUS => "Level 3–4 (avoid extreme/max)",
                BoardVendor.MSI => "Mode 3–4 (avoid extreme/max)",
                BoardVendor.Gigabyte => "Medium / High (avoid Turbo/Extreme)",
                BoardVendor.ASRock => "Level 2–3 (avoid Level 1/Extreme)",
                _ => "Medium (avoid max levels)"
            },
            Reason = "Too low can droop under boost; too high can overshoot and break CO stability (transient/WHEA).",
            WarningLevel = "Warning"
        });

        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "VRM efficiency",
            Path = VendorizePath(
                ctx.BoardVendor,
                asus: "AI Tweaker → Digi+ VRM",
                msi: "OC → DigitALL Power",
                gigabyte: "Tweaker → Advanced Voltage Settings",
                asrock: "OC Tweaker → Voltage Configuration",
                fallback: "BIOS → VRM"
            ),
            Setting = "Power Phase Control / Switching",
            Value = "Optimized / Auto (avoid Extreme unless high-end VRM)",
            Reason = "Extreme phase/switching can raise VRM heat and reduce sustained boost on midrange boards.",
            WarningLevel = "Info"
        });

        if (ctx.CpuVendor == CpuVendor.AMD)
        {


        // PBO warning
        list.Add(new Recommendation
        {
            Section = "BIOS",
            Title = "PBO limits",
            Path = ctx.IsAsus ? "Advanced → AMD Overclocking → Precision Boost Overdrive" : "BIOS → AMD Overclocking → Precision Boost Overdrive",
            Setting = "PBO Limits",
            Value = "Auto/Manual (avoid 'Motherboard' if chasing stability/temps)",
            Reason = "'Motherboard' limits are often overly aggressive and can increase heat without FPS gain.",
            WarningLevel = "Info"
        });

        
        }
return list;
    }

    static IEnumerable<Recommendation> BuildDisplay(SystemContext ctx)
    {
        return new[]
        {
            new Recommendation
            {
                Section = "DISPLAY",
                Title = "Refresh rate",
                Path = "Windows → Display → Advanced display",
                Setting = "Refresh rate",
                Value = ctx.Display.HasDisplay
                    ? $"Detected current: {ctx.Display.Width}x{ctx.Display.Height} @ {ctx.Display.RefreshHz} Hz (verify max)"
                    : "Verify max refresh rate is selected",
                Reason = "Wrong refresh rate ruins competitive smoothness."
            }
        };
    }

    static IEnumerable<Recommendation> BuildStorage(SystemContext ctx)
    {
        if (ctx.Disks.Count == 0) return Array.Empty<Recommendation>();

        bool hasNvme = ctx.Disks.Any(d => d.Kind == DiskKind.NVMe);
        bool hasSsd = ctx.Disks.Any(d => d.Kind == DiskKind.SATA_SSD);

        return new[]
        {
            new Recommendation
            {
                Section = "STORAGE",
                Title = "Game install drive",
                Path = "General",
                Setting = "Preferred drive",
                Value = hasNvme ? "NVMe preferred for games" : hasSsd ? "SSD preferred for games" : "Avoid HDD for modern games if possible",
                Reason = "Reduces loads and asset-stream hitches."
            }
        };
    }

    
static IEnumerable<Recommendation> BuildTesting(SystemContext ctx)
    {
        return new[]
        {
            new Recommendation
            {
                Section = "TESTING",
                Title = "Stability priority order",
                Path = "General",
                Setting = "If you crash / reset",
                Value = "Back off GPU memory OC → GPU core → CPU CO → RAM speed/FCLK",
                Reason = "Most instability is VRAM OC or CO/FCLK related.",
                WarningLevel = "None"
            },
            new Recommendation
            {
                Section = "TESTING",
                Title = "Validation",
                Path = "General",
                Setting = "How to validate changes",
                Value = "Change one thing at a time; test in your real game 30–60 mins; don’t stack unknowns.",
                Reason = "Prevents chasing multiple variables.",
                WarningLevel = "None"
            },
            new Recommendation
            {
                Section = "TESTING",
                Title = "Crash triage",
                Path = "Quick diagnosis",
                Setting = "Common patterns",
                Value = "Boot fail → RAM/FCLK/GDM | Load/alt-tab crash → PBO/LLC transient | Only under GPU load → VRAM/core/power | Idle crash → CO too negative",
                Reason = "These aren’t perfect, but they shorten debugging by pointing at the most common root cause first.",
                WarningLevel = "Info"
            }
        };
    }
}

#endregion

#region AMD Heuristics (tighter tables)

static class AmdHeuristics
{
    public struct PboLimits { public int PptW; public int TdcA; public int EdcA; }
    public struct CoGuess { public int Start; public int Typical; }

    public static PboLimits GuessPbo(CpuClass c, TuningProfile profile)
    {
        // These are “start points” targeted for competitive stability.
        // Buckets:
        // - AM4 65W:  88/60/90
        // - AM4 105W: 142/95/140
        // - AM5 65W:  90/65/100 (slightly higher transient room)
        // - AM5 105W: 150/100/160
        // - AM5 120W (X3D-ish): 135/90/140 (conservative stability)
        // - AM5 170W: 190/130/190 (conservative vs “unlimited”)

        var era = c.AmdEra;
        var tdp = c.AmdTdp;

        if (era == AmdSocketEra.AM4)
        {
            if (tdp == AmdTdpBucket.W65) return new PboLimits { PptW = 88, TdcA = 60, EdcA = 90 };
            // AM4 X3D (5800X3D) is special; keep conservative
            if (c.IsX3D) return new PboLimits { PptW = 120, TdcA = 80, EdcA = 120 };
            return new PboLimits { PptW = 142, TdcA = 95, EdcA = 140 };
        }

        if (era == AmdSocketEra.AM5)
        {
            if (c.IsX3D || tdp == AmdTdpBucket.W120)
                return new PboLimits { PptW = 135, TdcA = 90, EdcA = 140 };

            if (tdp == AmdTdpBucket.W65)
                return new PboLimits { PptW = 90, TdcA = 65, EdcA = 100 };

            if (tdp == AmdTdpBucket.W170)
                return new PboLimits { PptW = 190, TdcA = 130, EdcA = 190 };

            // default AM5 105W
            return new PboLimits { PptW = 150, TdcA = 100, EdcA = 160 };
        }

        // Unknown → stable middle
        return new PboLimits { PptW = 142, TdcA = 95, EdcA = 140 };
    }

    public static CoGuess GuessCurveOptimizer(CpuClass c, TuningProfile profile)
    {
        // Conservative starts; widen only if user tests stable.
        // X3D: more conservative.
        if (c.IsX3D) return new CoGuess { Start = 6, Typical = 12 };

        // AM5 (7000): start -10, typical -15 (many can go further but we keep stable)
        if (c.RyzenSeries >= 7000) return new CoGuess { Start = 10, Typical = 15 };

        // AM4 (5000): start -10, typical -15
        if (c.RyzenSeries >= 5000) return new CoGuess { Start = 10, Typical = 15 };

        // older: conservative
        return new CoGuess { Start = 8, Typical = 12 };
    }

    public static int GuessBoostOverrideMHz(CpuClass c, TuningProfile profile)
    {
        // X3D: keep 0 by default (stability-first). Non-X3D: +200 is common goal.
        if (c.IsX3D) return 0;

        // Daily can be +200; competitive also +200 but only if temps allow.
        return 200;
    }

    public static string GuessScalar(CpuClass c, TuningProfile profile)
    {
        if (profile == TuningProfile.Competitive) return "Auto (or 1X–2X)";
        return "Auto (or 1X–4X if temps are excellent)";
    }

    public static string GuessThermalLimitC(CpuClass c, TuningProfile profile)
    {
        if (profile == TuningProfile.Competitive) return "85–90°C target";
        return "Auto (or 90°C cap)";
    }
}

#endregion

#region NVIDIA / AMD GPU Heuristics (tighter tables)

static class NvidiaHeuristics
{
    public struct Ranges
    {
        public string Power;
        public string Core;
        public string Mem;
        public string Undervolt;
    }

    public static Ranges GuessRanges(GpuClass g, TuningProfile profile)
    {
        // Baselines by gen + tier, adjusted for laptop/factory OC.
        // Competitive: choose the lower/middle of safe typical ranges.

        string power = "MAX";
        string core, mem, uv;

        int gen = g.NvidiaRtxGen ?? 0;
        int tier = g.NvidiaTier ?? 70;

        bool laptop = g.IsLaptop;
        bool ocModel = g.IsFactoryOC;

        // Helper to tighten by tier
        (string core, string mem) GenTier(int gen, int tier)
        {
            // Default buckets (desktop)
            if (gen >= 50)
            {
                // 50-series baseline
                if (tier >= 90) return ("+120 to +240", "+700 to +1300");
                if (tier >= 80) return ("+140 to +280", "+800 to +1500");
                if (tier >= 70) return ("+150 to +300", "+800 to +1600");
                return ("+120 to +260", "+700 to +1400");
            }
            if (gen >= 40)
            {
                if (tier >= 90) return ("+120 to +230", "+700 to +1300");
                if (tier >= 80) return ("+140 to +260", "+800 to +1500");
                if (tier >= 70) return ("+150 to +280", "+800 to +1600");
                return ("+120 to +240", "+700 to +1400");
            }
            if (gen >= 30)
            {
                if (tier >= 90) return ("+80 to +170", "+600 to +1100");
                if (tier >= 80) return ("+90 to +190", "+700 to +1200");
                if (tier >= 70) return ("+100 to +200", "+700 to +1300");
                return ("+80 to +180", "+600 to +1200");
            }
            // 20-series and older
            if (tier >= 80) return ("+60 to +140", "+400 to +800");
            return ("+50 to +130", "+300 to +700");
        }

        (core, mem) = GenTier(gen, tier);

        if (laptop)
        {
            power = "As allowed (many laptops locked)";
            // shrink ranges
            core = "+50 to +150";
            mem = "+250 to +800";
        }
        else if (ocModel)
        {
            // slight tilt to the upper end (still not “max it”)
            core = BumpUpper(core, 20);
            mem = BumpUpper(mem, 100);
        }

        // Undervolt suggestions by gen
        if (laptop)
            uv = "Optional: undervolt only if your laptop BIOS/software supports it; prioritize temps/fan stability.";
        else if (gen >= 50)
            uv = "Try ~0.90–0.95V at your highest stable clock (then add modest memory OC).";
        else if (gen >= 40)
            uv = "Try ~0.90–0.95V at a stable boost clock (then add modest memory OC).";
        else if (gen >= 30)
            uv = "Try ~0.875–0.93V at a stable clock (then add modest memory OC).";
        else
            uv = "Try a mild undervolt if supported; don’t chase aggressive curves on older cards.";

        return new Ranges
        {
            Power = power,
            Core = core,
            Mem = mem,
            Undervolt = uv
        };
    }

    static string BumpUpper(string range, int add)
    {
        // "+150 to +280" -> add to upper bound
        // If parsing fails, return original.
        try
        {
            var parts = range.Replace(" ", "").Split("to", StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2) return range;

            int a = ParseSigned(parts[0]);
            int b = ParseSigned(parts[1]);
            b += add;
            return $"+{a} to +{b}";
        }
        catch { return range; }
    }

    static int ParseSigned(string s)
    {
        s = s.Replace("+", "");
        return int.Parse(s, CultureInfo.InvariantCulture);
    }
}

static class AmdGpuHeuristics
{
    public struct Ranges { public string Power; public string Core; public string Mem; }

    public static Ranges GuessRanges(GpuClass g, TuningProfile profile)
    {
        // AMD ranges are more variable; keep conservative.
        if (g.IsLaptop)
        {
            return new Ranges
            {
                Power = "As allowed (often locked)",
                Core = "+25 to +75",
                Mem = "+25 to +100"
            };
        }

        return new Ranges
        {
            Power = "+5% to +15%",
            Core = "+50 to +150",
            Mem = "+50 to +200"
        };
    }
}

#endregion

#region Hardware Detection + Classifiers

static class HardwareDetector
{
    public static SystemContext Detect(TuningProfile profile)
    {
        var ctx = new SystemContext { Profile = profile };

        ctx.OsCaption = Wmi.Get("Win32_OperatingSystem", "Caption");
        ctx.OsVersion = Wmi.Get("Win32_OperatingSystem", "Version");
        ctx.ActivePowerPlan = GetActivePowerPlanName();

        ctx.CpuName = Wmi.Get("Win32_Processor", "Name");
        ctx.CpuVendor = DetectCpuVendor(ctx.CpuName);
        ctx.CpuClass = CpuClassifier.Classify(ctx.CpuName, ctx.CpuVendor);

        ctx.BoardManufacturer = Wmi.Get("Win32_BaseBoard", "Manufacturer");
        ctx.BoardProduct = Wmi.Get("Win32_BaseBoard", "Product");
        ctx.BoardVendor = DetectBoardVendor(ctx.BoardManufacturer);

        ctx.BiosVendor = Wmi.Get("Win32_BIOS", "Manufacturer");
        ctx.BiosVersion = Wmi.Get("Win32_BIOS", "SMBIOSBIOSVersion");
        ctx.BiosDate = Wmi.FormatDate(Wmi.Get("Win32_BIOS", "ReleaseDate"));

        DetectRam(ctx);
        DetectGpu(ctx);
        DetectDisks(ctx);

        ctx.Display = DisplayDetector.GetPrimaryDisplay();
        return ctx;
    }

    static void DetectRam(SystemContext ctx)
    {
        ulong ramBytes = 0;
        int smbiosType = 0;
        var speeds = new List<int>();
        var cfgs = new List<int>();
        int dimmCount = 0;
        int maxRanksPerDimm = 0;

        foreach (var m in Wmi.GetMany("Win32_PhysicalMemory"))
        {
            dimmCount++;
            ramBytes += Wmi.ToULong(m["Capacity"]);
            int speed = Wmi.ToInt(m["Speed"]);
            int cfg = Wmi.ToInt(m["ConfiguredClockSpeed"]);
            int st = Wmi.ToInt(m["SMBIOSMemoryType"]);

            // Best-effort: ranks per DIMM (not always exposed by all firmware/WMI providers).
            // If present, values are usually 1 (single-rank) or 2 (dual-rank).
            try
            {
                if (m.Properties["Rank"] != null && m["Rank"] != null)
                {
                    int r = Wmi.ToInt(m["Rank"]);
                    if (r > maxRanksPerDimm) maxRanksPerDimm = r;
                }
            }
            catch { /* ignore */ }

            if (speed > 0) speeds.Add(speed);
            if (cfg > 0) cfgs.Add(cfg);
            if (st > 0) smbiosType = st;
        }

        ctx.DetectedDimmCount = dimmCount > 0 ? dimmCount : null;
        if (maxRanksPerDimm >= 2) ctx.DetectedRankHint = "DualRank";
        else if (maxRanksPerDimm == 1) ctx.DetectedRankHint = "SingleRank";

        ctx.TotalRamGb = ramBytes / 1024.0 / 1024.0 / 1024.0;

        ctx.MemoryType = smbiosType == 26 ? MemoryType.DDR4 :
                         smbiosType == 34 ? MemoryType.DDR5 :
                         MemoryType.Unknown;

        int raw = speeds.Count > 0 ? speeds.Max() : (cfgs.Count > 0 ? cfgs.Max() : 0);
        if (raw > 0)
        {
            // Normalize to MT/s:
            // If raw looks like real clock MHz (1000..2200), multiply by 2. Otherwise assume already MT/s.
            ctx.EffectiveMemMts = (raw >= 1000 && raw <= 2200) ? raw * 2 : raw;
        }

        // Save module hints for the RAM calculator UI.
        if (dimmCount > 0) ctx.DetectedDimmCount = dimmCount;
        if (maxRanksPerDimm >= 2) ctx.DetectedRankHint = "DualRank";
        else if (maxRanksPerDimm == 1) ctx.DetectedRankHint = "SingleRank";

        // Ryzen DDR4 baseline: target FCLK = RAM MT/s ÷ 2.
        // Apply conservative caps by Ryzen generation (stability-first defaults).
        if (ctx.CpuVendor == CpuVendor.AMD && ctx.MemoryType == MemoryType.DDR4 && ctx.EffectiveMemMts.HasValue)
        {
            int desired = ctx.EffectiveMemMts.Value / 2;
            ctx.IdealFclkMHz = desired;

            int cap = 2000; // hard ceiling for "auto" suggestions

            // Zen2 (Ryzen 3000) commonly tops out ~1800 stable.
            if ((ctx.CpuClass.RyzenSeries ?? 0) >= 3000 && (ctx.CpuClass.RyzenSeries ?? 0) < 5000)
                cap = 1800;

            // Zen3 (Ryzen 5000) often does 1900; 5800X3D tends to be pickier → 1800.
            if ((ctx.CpuClass.RyzenSeries ?? 0) >= 5000 && (ctx.CpuClass.RyzenSeries ?? 0) < 7000)
                cap = ctx.CpuClass.IsX3D ? 1800 : 1900;

            ctx.TargetFclkMHz = Math.Min(desired, cap);

            if (ctx.TargetFclkMHz == desired) ctx.FabricMode = "1:1";
            else ctx.FabricMode = $"1:1 (capped at {ctx.TargetFclkMHz} — try {desired} if stable)";
        }
    }

    static void DetectGpu(SystemContext ctx)
    {
        foreach (var g in Wmi.GetMany("Win32_VideoController"))
        {
            var name = (g["Name"]?.ToString() ?? "").Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            ctx.Gpus.Add(new GpuInfo
            {
                Name = name,
                DriverVersion = (g["DriverVersion"]?.ToString() ?? "").Trim(),
                Vendor = DetectGpuVendor(name)
            });
        }
    }

    static void DetectDisks(SystemContext ctx)
    {
        foreach (var d in Wmi.GetMany("Win32_DiskDrive"))
        {
            string model = (d["Model"]?.ToString() ?? "").Trim();
            string iface = (d["InterfaceType"]?.ToString() ?? "").Trim();
            string media = (d["MediaType"]?.ToString() ?? "").Trim();
            ulong size = Wmi.ToULong(d["Size"]);
            if (string.IsNullOrWhiteSpace(model)) continue;

            ctx.Disks.Add(new DiskInfo
            {
                Model = model,
                InterfaceType = iface,
                MediaType = media,
                SizeBytes = size,
                Kind = GuessDiskKind(model, iface, media)
            });
        }
    }

    static DiskKind GuessDiskKind(string model, string iface, string media)
    {
        string s = (model + " " + iface + " " + media).ToLowerInvariant();
        if (s.Contains("nvme")) return DiskKind.NVMe;
        if (s.Contains("usb")) return DiskKind.USB;
        if (s.Contains("ssd") || s.Contains("solid state")) return DiskKind.SATA_SSD;
        if (s.Contains("hdd") || s.Contains("hard disk") || s.Contains("5400") || s.Contains("7200")) return DiskKind.SATA_HDD;
        return DiskKind.Other;
    }

    static CpuVendor DetectCpuVendor(string cpuName)
    {
        var s = cpuName.ToLowerInvariant();
        if (s.Contains("amd") || s.Contains("ryzen") || s.Contains("threadripper")) return CpuVendor.AMD;
        if (s.Contains("intel") || s.Contains("core")) return CpuVendor.Intel;
        return CpuVendor.Unknown;
    }

    static BoardVendor DetectBoardVendor(string manufacturer)
    {
        var s = (manufacturer ?? "").Trim().ToLowerInvariant();
        if (s.Contains("asus")) return BoardVendor.ASUS;
        if (s.Contains("micro-star") || s.Contains("msi")) return BoardVendor.MSI;
        if (s.Contains("gigabyte")) return BoardVendor.Gigabyte;
        if (s.Contains("asrock")) return BoardVendor.ASRock;
        if (string.IsNullOrWhiteSpace(s)) return BoardVendor.Unknown;
        return BoardVendor.Other;
    }

    static string VendorizePath(string genericPath, string vendor)
    {
        if (string.IsNullOrWhiteSpace(genericPath)) return genericPath;
        var v = (vendor ?? "").Trim().ToLowerInvariant();
        if (v.Contains("asus")) return genericPath.Replace("AI Tweaker", "AI Tweaker / Extreme Tweaker").Replace("Advanced", "Advanced Mode");
        if (v.Contains("msi")) return genericPath.Replace("AI Tweaker", "OC").Replace("Advanced", "OC");
        if (v.Contains("gigabyte")) return genericPath.Replace("AI Tweaker", "Tweaker").Replace("Advanced", "Advanced Frequency Settings");
        if (v.Contains("asrock")) return genericPath.Replace("AI Tweaker", "OC Tweaker");
        return genericPath;
    }

    static string VendorizePath(BoardVendor v, string asus, string msi, string gigabyte, string asrock, string fallback)
    {
        return v switch
        {
            BoardVendor.ASUS => asus,
            BoardVendor.MSI => msi,
            BoardVendor.Gigabyte => gigabyte,
            BoardVendor.ASRock => asrock,
            _ => fallback
	    };
	}

	static string GetActivePowerPlanName()
    {
        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powercfg",
                Arguments = "/getactivescheme",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var p = System.Diagnostics.Process.Start(psi);
            if (p == null) return "";
            var output = p.StandardOutput.ReadToEnd();
            p.WaitForExit(1500);

            // Typical: "Power Scheme GUID: xxxx-....  (Balanced)"
            var open = output.IndexOf('(');
            var close = output.IndexOf(')', open + 1);
            if (open >= 0 && close > open)
                return output.Substring(open + 1, close - open - 1).Trim();

            return output.Trim();
        }
        catch
        {
            return "";
        }
    }

    static GpuVendor DetectGpuVendor(string gpuName)
    {
        var s = gpuName.ToLowerInvariant();
        if (s.Contains("nvidia") || s.Contains("geforce") || s.Contains("rtx") || s.Contains("gtx")) return GpuVendor.NVIDIA;
        if (s.Contains("amd") || s.Contains("radeon")) return GpuVendor.AMD;
        return GpuVendor.Unknown;
    }
}

static class CpuClassifier
{
    public static CpuClass Classify(string cpuName, CpuVendor vendor)
    {
        var c = new CpuClass();
        var s = (cpuName ?? "").ToUpperInvariant();

        c.IsX3D = s.Contains("X3D");

        if (vendor == CpuVendor.AMD)
        {
            c.AmdModel = FindAmdModelNumber(s);
            if (c.AmdModel.HasValue)
            {
                var m = c.AmdModel.Value;
                if (m >= 7000) c.RyzenSeries = 7000;
                else if (m >= 5000) c.RyzenSeries = 5000;
                else if (m >= 3000) c.RyzenSeries = 3000;
            }

            c.AmdEra = (c.RyzenSeries >= 7000) ? AmdSocketEra.AM5 :
                       (c.RyzenSeries >= 3000) ? AmdSocketEra.AM4 :
                       AmdSocketEra.Unknown;

            // TDP bucket from common naming patterns:
            // - Non-X: commonly 65W (5600, 5700X, 7600, 7700, 7900)
            // - X: commonly 105W (5600X, 5800X, 5900X, 5950X; 7600X/7700X)
            // - 7900X/7950X: commonly 170W class
            // - X3D AM5: treat as 120W bucket
            // - X3D AM4 (5800X3D): keep conservative 105W bucket but special-case in heuristics
            if (c.IsX3D)
            {
                c.AmdTdp = (c.AmdEra == AmdSocketEra.AM5) ? AmdTdpBucket.W120 : AmdTdpBucket.W105;
            }
            else
            {
                if (MatchesAny(s, "7900X", "7950X", "7900 X", "7950 X"))
                    c.AmdTdp = AmdTdpBucket.W170;
                else if (MatchesAny(s, "7600X", "7700X", "7800X", "5600X", "5800X", "5900X", "5950X"))
                    c.AmdTdp = AmdTdpBucket.W105;
                else
                    c.AmdTdp = AmdTdpBucket.W65;
            }
        }
        else if (vendor == CpuVendor.Intel)
        {
            c.IntelGen = FindIntelGen(s);
        }

        return c;
    }

    static bool MatchesAny(string s, params string[] needles)
        => needles.Any(n => s.Contains(n, StringComparison.Ordinal));

    static int FindIntelGen(string s)
    {
        // Look for patterns like i9-13900K, 14900K, 12700K
        var nums = ExtractNumbers(s);
        foreach (var n in nums)
        {
            if (n >= 12000 && n < 13000) return 12;
            if (n >= 13000 && n < 14000) return 13;
            if (n >= 14000 && n < 15000) return 14;
        }
        return 0;
    }

    static int FindAmdModelNumber(string s)
    {
        // Find 4-digit blocks likely to be Ryzen model numbers (e.g., 5800, 7800)
        var nums = ExtractNumbers(s);
        foreach (var n in nums)
            if (n >= 3000 && n < 10000) return n;
        return 0;
    }

    static List<int> ExtractNumbers(string s)
    {
        var nums = new List<int>();
        int cur = -1;
        foreach (var ch in s)
        {
            if (char.IsDigit(ch))
            {
                if (cur < 0) cur = 0;
                cur = cur * 10 + (ch - '0');
            }
            else
            {
                if (cur >= 0) { nums.Add(cur); cur = -1; }
            }
        }
        if (cur >= 0) nums.Add(cur);
        return nums;
    }
}

static class GpuClassifier
{
    public static GpuClass Classify(string gpuName)
    {
        var g = new GpuClass();
        var s = (gpuName ?? "").ToUpperInvariant();

        g.IsLaptop = s.Contains("LAPTOP") || s.Contains("MOBILE") || s.Contains("MAX-Q") || s.Contains("MAXQ");
        g.IsFactoryOC = s.Contains(" OC") || s.Contains("GAMING OC") || s.Contains("OC EDITION") || s.Contains("FACTORY OC");

        if (s.Contains("NVIDIA") || s.Contains("GEFORCE") || s.Contains("RTX") || s.Contains("GTX"))
        {
            g.Vendor = GpuVendor.NVIDIA;

            // Determine RTX model (first 4-digit block like 5070/4070/3080)
            g.NvidiaModel = ExtractFirst4Digits(s);
            if (g.NvidiaModel.HasValue)
            {
                int model = g.NvidiaModel.Value;
                // Gen from thousands digit: 5070->50, 4070->40, 3080->30, 2080->20
                int gen = model / 100;
                if (gen >= 50) g.NvidiaRtxGen = 50;
                else if (gen >= 40) g.NvidiaRtxGen = 40;
                else if (gen >= 30) g.NvidiaRtxGen = 30;
                else if (gen >= 20) g.NvidiaRtxGen = 20;

                // Tier from last two digits group (60/70/80/90/50)
                int tier = (model / 10) % 10; // e.g., 5070 -> 7, 4080 -> 8, 4090 -> 9
                g.NvidiaTier = tier switch
                {
                    9 => 90,
                    8 => 80,
                    7 => 70,
                    6 => 60,
                    5 => 50,
                    _ => (int?)null
                };
            }
            else
            {
                // fallback
                if (s.Contains("RTX 50")) g.NvidiaRtxGen = 50;
                else if (s.Contains("RTX 40")) g.NvidiaRtxGen = 40;
                else if (s.Contains("RTX 30")) g.NvidiaRtxGen = 30;
                else if (s.Contains("RTX 20")) g.NvidiaRtxGen = 20;
            }
        }
        else if (s.Contains("AMD") || s.Contains("RADEON"))
        {
            g.Vendor = GpuVendor.AMD;
        }
        else g.Vendor = GpuVendor.Unknown;

        return g;
    }

    static int? ExtractFirst4Digits(string s)
    {
        int cur = -1, len = 0;
        foreach (var ch in s)
        {
            if (char.IsDigit(ch))
            {
                if (cur < 0) { cur = 0; len = 0; }
                cur = cur * 10 + (ch - '0');
                len++;
                if (len == 4) return cur;
            }
            else
            {
                cur = -1;
                len = 0;
            }
        }
        return null;
    }
}

static class Wmi
{
    public static string Get(string cls, string prop)
    {
        using var s = new ManagementObjectSearcher($"SELECT {prop} FROM {cls}");
        foreach (ManagementObject o in s.Get())
            return o[prop]?.ToString() ?? "";
        return "";
    }

    public static List<ManagementObject> GetMany(string cls)
    {
        var list = new List<ManagementObject>();
        using var s = new ManagementObjectSearcher($"SELECT * FROM {cls}");

        foreach (ManagementObject o in s.Get())
            list.Add(o);
        return list;
    }

    public static int ToInt(object? o) { try { return Convert.ToInt32(o, CultureInfo.InvariantCulture); } catch { return 0; } }
    public static ulong ToULong(object? o) { try { return Convert.ToUInt64(o, CultureInfo.InvariantCulture); } catch { return 0; } }

    public static string FormatDate(string d)
    {
        try { return ManagementDateTimeConverter.ToDateTime(d).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
        catch { return d; }
    }
}

#endregion

#region Display Detection

static class DisplayDetector
{
    public static DisplayInfo GetPrimaryDisplay()
    {
        var devMode = new DEVMODE();
        devMode.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));

        bool ok = EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref devMode);
        if (!ok) return new DisplayInfo { HasDisplay = false };

        return new DisplayInfo
        {
            HasDisplay = true,
            Width = devMode.dmPelsWidth,
            Height = devMode.dmPelsHeight,
            RefreshHz = devMode.dmDisplayFrequency
        };
    }

    private const int ENUM_CURRENT_SETTINGS = -1;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    private struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;

        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;

        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;

        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;

        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [DllImport("user32.dll", CharSet = CharSet.Ansi)]
    private static extern bool EnumDisplaySettings(string? lpszDeviceName, int iModeNum, ref DEVMODE lpDevMode);
}

#endregion

#region Reports

static class ReportBuilder
{
    public static string BuildText(SystemContext ctx, List<Recommendation> recs)
    {
        var sb = new StringBuilder(80_000);

        sb.AppendLine("OC ADVISOR — OUTPUT ONLY (EDUCATED GUESSES, TIGHTER TABLES)");
        sb.AppendLine("==========================================================");
        sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"Profile:   {ctx.Profile}");
        sb.AppendLine();

        sb.AppendLine("SYSTEM");
        sb.AppendLine("------");
        sb.AppendLine($"OS:    {ctx.OsCaption} ({ctx.OsVersion})");
        sb.AppendLine($"CPU:   {ctx.CpuName} [{ctx.CpuVendor}]");
        if (ctx.CpuVendor == CpuVendor.AMD)
            sb.AppendLine($"CPU Class: {ctx.CpuClass.AmdEra} | {ctx.CpuClass.AmdTdp} | {(ctx.CpuClass.IsX3D ? "X3D" : "non-X3D")} | RyzenSeries {ctx.CpuClass.RyzenSeries?.ToString() ?? "?"}");
        if (ctx.CpuVendor == CpuVendor.Intel)
            sb.AppendLine($"CPU Class: Intel Gen {ctx.CpuClass.IntelGen?.ToString() ?? "?"}");

        sb.AppendLine($"Board: {ctx.BoardManufacturer} {ctx.BoardProduct}");
        sb.AppendLine($"BIOS:  {ctx.BiosVendor} {ctx.BiosVersion} ({ctx.BiosDate})");
        sb.AppendLine($"RAM:   {ctx.TotalRamGb:0.#} GB [{ctx.MemoryType}]");
        if (ctx.EffectiveMemMts.HasValue)
            sb.AppendLine($"RAM Rate: {ctx.EffectiveMemMts.Value} MT/s (FCLK goal: {(ctx.TargetFclkMHz.HasValue ? ctx.TargetFclkMHz.Value + " MHz" : "n/a")})");

        sb.AppendLine();
        if (ctx.Display.HasDisplay)
            sb.AppendLine($"Display: {ctx.Display.Width}x{ctx.Display.Height} @ {ctx.Display.RefreshHz} Hz");

        sb.AppendLine();
        sb.AppendLine("GPU(s):");
        foreach (var g in ctx.Gpus)
            sb.AppendLine($"- {g.Name} [{g.Vendor}] | Driver {g.DriverVersion}");

        if (ctx.Disks.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Disks:");
            foreach (var d in ctx.Disks)
                sb.AppendLine($"- {d.Model} | {d.Kind} | {FormatBytes(d.SizeBytes)}");
        }

        sb.AppendLine();

        // Auto-grouping: print sections in order, only if there are items
        foreach (var section in RuleEngine.SectionOrder)
        {
            var items = recs.Where(x => x.Section.Equals(section, StringComparison.OrdinalIgnoreCase)).ToList();
            if (items.Count == 0) continue;

            sb.AppendLine(section);
            sb.AppendLine(new string('-', section.Length));
            foreach (var it in items)
            {
                sb.AppendLine($"[{it.Title}]");
                if (!string.IsNullOrWhiteSpace(it.Path)) sb.AppendLine($"Path: {it.Path}");
                sb.AppendLine($"- {it.Setting}: {it.Value}");
                if (!string.IsNullOrWhiteSpace(it.Reason)) sb.AppendLine($"  Reason: {it.Reason}");
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }

    public static string BuildJson(SystemContext ctx, List<Recommendation> recs)
    {
        var dto = new
        {
            Generated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            Profile = ctx.Profile.ToString(),
            System = ctx,
            Recommendations = recs
        };

        return JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
    }

    static string FormatBytes(ulong bytes)
    {
        if (bytes == 0) return "0 B";
        double b = bytes;
        string[] units = { "B", "KB", "MB", "GB", "TB", "PB" };
        int u = 0;
        while (b >= 1024 && u < units.Length - 1) { b /= 1024; u++; }
        return $"{b:0.##} {units[u]}";
    }
}

#endregion
