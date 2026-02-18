using System;

namespace OcAdvisor;

internal enum CoolingType
{
    AirCooler,
    Aio240,
    Aio360,
    CustomLoop
}

internal enum SiliconQuality
{
    Conservative,
    Average,
    AboveAverage,
    GoldenSample
}

internal sealed class OverclockAssumptions
{
    public CoolingType Cooling { get; init; } = CoolingType.Aio240;
    public SiliconQuality Silicon { get; init; } = SiliconQuality.Average;

    public override string ToString() => $"{Cooling} / {Silicon}";
}
