using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace OcAdvisor;

public partial class App : Application
{
    private static string CrashLogPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OcAdvisor", "crash.log");

    public App()
    {
        // Catch UI thread exceptions
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        // Catch non-UI thread exceptions
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
    }

    private static void EnsureCrashLogDir()
    {
        var dir = Path.GetDirectoryName(CrashLogPath);
        if (!string.IsNullOrWhiteSpace(dir))
            Directory.CreateDirectory(dir);
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        try
        {
            EnsureCrashLogDir();
            File.AppendAllText(CrashLogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] DispatcherUnhandledException\n{e.Exception}\n\n");
        }
        catch { /* ignore logging failures */ }

        MessageBox.Show(e.Exception.ToString(), "OC Advisor - Crash", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true; // keep app alive
    }

    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        try
        {
            EnsureCrashLogDir();
            File.AppendAllText(CrashLogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] UnhandledException (IsTerminating={e.IsTerminating})\n{e.ExceptionObject}\n\n");
        }
        catch { /* ignore logging failures */ }

        try
        {
            MessageBox.Show(e.ExceptionObject?.ToString() ?? "Unknown crash",
                "OC Advisor - Crash", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch { /* if UI can't show */ }
    }
}
