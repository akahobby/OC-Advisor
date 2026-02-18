namespace OcAdvisor;

internal sealed class OverclockProfile
{
    public string CpuSummary { get; set; } = "";
    public string RamSummary { get; set; } = "";
    public string GpuSummary { get; set; } = "";
    public string BiosSummary { get; set; } = "";

    public int StabilityConfidence { get; set; }
    public string RiskLevel { get; set; } = "";
    public string EstimatedPerformanceGain { get; set; } = "";
    public double AssumedCpuVcore { get; set; }
}
