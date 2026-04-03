using System.IO;
using System.Text.Json;

namespace BatchRename.Services;

/// <summary>
/// Loads and saves <see cref="AppSettings"/> from a JSON file placed in the
/// same directory as the running executable — portable, no registry, no AppData.
/// </summary>
public static class SettingsService
{
    private static readonly JsonSerializerOptions s_opts = new() { WriteIndented = true };

    /// <summary>Full path to the settings file (next to the .exe).</summary>
    public static string SettingsPath { get; } = GetSettingsPath();

    private static string GetSettingsPath()
    {
        // Environment.ProcessPath is the actual .exe location, even for single-file publish.
        // Fall back to AppContext.BaseDirectory if unavailable.
        var dir = Path.GetDirectoryName(Environment.ProcessPath)
               ?? AppContext.BaseDirectory;
        return Path.Combine(dir, "settings.json");
    }

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { /* corrupt file — use defaults */ }
        return new AppSettings();
    }

    public static void Save(AppSettings settings)
    {
        try
        {
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, s_opts));
        }
        catch { /* non-fatal — writable location not guaranteed in all deployments */ }
    }
}
