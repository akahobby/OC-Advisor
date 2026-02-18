using System;

namespace OcAdvisor;

internal static class StabilityCalculator
{
    /// <summary>
    /// Produces a conservative "confidence" score intended for user guidance (not a guarantee).
    /// </summary>
    public static int Calculate(CoolingType cooling, SiliconQuality silicon, double assumedCpuVcore)
    {
        int score = 70;

        // Cooling headroom
        score += cooling switch
        {
            CoolingType.AirCooler => -5,
            CoolingType.Aio240 => 0,
            CoolingType.Aio360 => 6,
            CoolingType.CustomLoop => 10,
            _ => 0
        };

        // Silicon variance
        score += silicon switch
        {
            SiliconQuality.Conservative => -6,
            SiliconQuality.Average => 0,
            SiliconQuality.AboveAverage => 4,
            SiliconQuality.GoldenSample => 8,
            _ => 0
        };

        // Voltage realism: higher vcore reduces confidence quickly
        if (assumedCpuVcore > 1.40) score -= 18;
        else if (assumedCpuVcore > 1.36) score -= 10;
        else if (assumedCpuVcore > 1.32) score -= 4;

        return Math.Clamp(score, 40, 98);
    }

    public static string DetermineRisk(double assumedCpuVcore)
    {
        if (assumedCpuVcore <= 1.32) return "Safe Daily";
        if (assumedCpuVcore <= 1.40) return "Aggressive Daily";
        return "Benchmark Only";
    }
}
