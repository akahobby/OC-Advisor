using System;

namespace OcAdvisor;

internal static class PerformanceEstimator
{
    /// <summary>
    /// Heuristic uplift estimate. This is intentionally rough and biased conservative.
    /// </summary>
    public static string Estimate(SystemContext ctx, OverclockAssumptions a)
    {
        // CPU: assume mild PBO/CO / Intel boost tuning
        var cpu = (ctx.CpuVendor == CpuVendor.AMD)
            ? (ctx.CpuClass.IsX3D ? "CPU +3–7%" : "CPU +5–12%")
            : "CPU +4–10%";

        // RAM: depends on memory type and whether user is likely to push
        var ram = ctx.MemoryType switch
        {
            MemoryType.DDR4 => "RAM +2–6%",
            MemoryType.DDR5 => "RAM +1–5%",
            _ => "RAM +1–4%"
        };

        // GPU: depends on tier (roughly)
        var gc = ctx.PrimaryGpuClass;
        string gpu;
        if (gc.Vendor == GpuVendor.NVIDIA)
        {
            gpu = gc.NvidiaTier switch
            {
                90 => "GPU +4–10%",
                80 => "GPU +5–12%",
                70 => "GPU +6–14%",
                60 => "GPU +6–12%",
                _ => "GPU +5–12%"
            };
        }
        else if (gc.Vendor == GpuVendor.AMD) gpu = "GPU +5–12%";
        else gpu = "GPU +4–10%";

        // Cooling/silicon can nudge top-end slightly (keep it subtle)
        bool highHeadroom = a.Cooling is CoolingType.Aio360 or CoolingType.CustomLoop;
        bool goodSilicon = a.Silicon is SiliconQuality.AboveAverage or SiliconQuality.GoldenSample;
        if (highHeadroom && goodSilicon)
        {
            // bump the upper bound by ~1%
            cpu = cpu.Replace("12%", "13%").Replace("10%", "11%").Replace("7%", "8%");
            gpu = gpu.Replace("14%", "15%").Replace("12%", "13%");
        }

        return $"{cpu} • {ram} • {gpu} (rough estimate)";
    }
}
