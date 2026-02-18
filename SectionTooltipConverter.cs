using System;
using System.Globalization;
using System.Windows.Data;

namespace OcAdvisor;

public sealed class SectionTooltipConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var name = (value?.ToString() ?? string.Empty).Trim().ToUpperInvariant();
        return name switch
        {
            "CPU" => "CPU tuning & boost behavior",
            "MEMORY" => "RAM & Infinity Fabric (stability/latency sensitive)",
            "GPU" => "GPU/driver-side settings",
            "BIOS" => "Firmware (UEFI/BIOS) settings",
            "DISPLAY" => "Display & frame pacing",
            "SYSTEM" => "System baseline info",
            _ => "Section"
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
