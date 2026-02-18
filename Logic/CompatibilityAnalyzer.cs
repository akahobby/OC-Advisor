using System;
using System.Collections.Generic;

namespace OcAdvisor;

internal static class CompatibilityAnalyzer
{
    public static IEnumerable<Recommendation> Build(SystemContext ctx, OverclockAssumptions a)
    {
        var list = new List<Recommendation>();

        // DDR5 + 4 DIMMs at high MT/s can be unstable on many boards/IMCs
        if (ctx.MemoryType == MemoryType.DDR5 && ctx.DetectedDimmCount == 4)
        {
            if (ctx.EffectiveMemMts.HasValue && ctx.EffectiveMemMts.Value >= 6400)
            {
                list.Add(new Recommendation
                {
                    Section = "MEMORY",
                    Setting = "Compatibility check",
                    Value = "4× DIMM + DDR5 ≥ 6400 MT/s",
                    Path = "BIOS → Memory",
                    Reason = "Many IMCs/boards struggle with 4 sticks at high DDR5 speeds. If unstable, drop MT/s or loosen timings; prioritize stability over peak speed.",
                    WarningLevel = "Warning"
                });
            }
        }

        // AM5: very high FCLK is uncommon; warn if "ideal" suggests too high (1:1 math)
        if (ctx.CpuVendor == CpuVendor.AMD && ctx.CpuClass.AmdEra == AmdSocketEra.AM5 && ctx.IdealFclkMHz.HasValue && ctx.IdealFclkMHz.Value > 2200)
        {
            list.Add(new Recommendation
            {
                Section = "CPU",
                Setting = "Fabric sanity check",
                Value = $"Ideal FCLK math: {ctx.IdealFclkMHz} MHz",
                Path = "BIOS → AMD Overclocking → Infinity Fabric",
                Reason = "AM5 typically prefers stability-first fabric targets. If you see WHEA errors, lower FCLK and/or run 1:2 where appropriate.",
                WarningLevel = "Info"
            });
        }

        // Generic: remind that high assumed vcore + unknown VRM is risky
        var assumedVcore = AssumedCpuVcore(a);
        if (assumedVcore > 1.36)
        {
            list.Add(new Recommendation
            {
                Section = "BIOS",
                Setting = "Voltage caution",
                Value = $"Assumed CPU Vcore target: {assumedVcore:0.00} V",
                Path = "BIOS → CPU Core Voltage / LLC",
                Reason = "Higher Vcore increases heat and VRM load. If temps or stability are marginal, back off voltage/boost and retest.",
                WarningLevel = "Warning"
            });
        }

        return list;
    }

    public static double AssumedCpuVcore(OverclockAssumptions a)
    {
        // Baseline "competitive safe-ish" assumptions; used for scoring and warnings only.
        double v = a.Cooling switch
        {
            CoolingType.AirCooler => 1.28,
            CoolingType.Aio240 => 1.32,
            CoolingType.Aio360 => 1.36,
            CoolingType.CustomLoop => 1.38,
            _ => 1.32
        };

        v += a.Silicon switch
        {
            SiliconQuality.GoldenSample => -0.02,
            SiliconQuality.AboveAverage => -0.01,
            SiliconQuality.Average => 0.00,
            SiliconQuality.Conservative => 0.02,
            _ => 0.00
        };

        return Math.Clamp(v, 1.20, 1.45);
    }
}
