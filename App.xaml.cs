using System.Windows;
using Microsoft.Win32;
using Wpf.Ui.Appearance;
using Application = System.Windows.Application;

namespace BatchRename;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Apply Windows app mode theme immediately on startup
        ApplySystemTheme();

        // Watch for Windows theme changes (light/dark toggle in Settings)
        SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;

        var window = new MainWindow(e.Args);
        MainWindow = window;
        window.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
        base.OnExit(e);
    }

    private void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
    {
        // UserPreferenceCategory.General fires when the user changes
        // the Windows colour mode (light / dark) in Settings
        if (e.Category == UserPreferenceCategory.General)
            Dispatcher.Invoke(ApplySystemTheme);
    }

    /// <summary>
    /// Reads the Windows "Apps use light theme" registry value and applies
    /// the matching WPF-UI theme. Falls back to Light if the key is absent.
    /// </summary>
    internal static void ApplySystemTheme()
    {
        var theme = IsWindowsDarkMode() ? ApplicationTheme.Dark : ApplicationTheme.Light;
        ApplicationThemeManager.Apply(theme);
    }

    private static bool IsWindowsDarkMode()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            // AppsUseLightTheme == 0 means dark mode is active
            return key?.GetValue("AppsUseLightTheme") is int v && v == 0;
        }
        catch
        {
            return false; // default to light on any failure
        }
    }
}
