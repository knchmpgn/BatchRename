namespace BatchRename.Services;

/// <summary>Settings persisted to settings.json next to the executable.</summary>
public class AppSettings
{
    /// <summary>Last saved window left position.</summary>
    public double WindowLeft { get; set; } = double.NaN;

    /// <summary>Last saved window top position.</summary>
    public double WindowTop { get; set; } = double.NaN;

    /// <summary>Last saved window width (in device-independent pixels).</summary>
    public double WindowWidth { get; set; } = 1600;

    /// <summary>Last saved window height (in device-independent pixels).</summary>
    public double WindowHeight { get; set; } = 1400;

    /// <summary>Whether the window was maximized when last closed.</summary>
    public bool IsMaximized { get; set; } = false;
}
