using System.Windows;
using System.Windows.Controls;
using BatchRename.ViewModels;
using BatchRename.Services;

using ContextMenu = System.Windows.Controls.ContextMenu;
using DragEventArgs = System.Windows.DragEventArgs;
using DragDropEffects = System.Windows.DragDropEffects;
using DataFormats = System.Windows.DataFormats;

namespace BatchRename;

public partial class MainWindow
{
    private readonly MainViewModel _vm;
    private readonly AppSettings _settings;

    public MainWindow()
    {
        InitializeComponent();

        _vm = new MainViewModel();
        DataContext = _vm;

        // Load settings and restore window state
        _settings = SettingsService.Load();

        // Always use 1600x1400 as the primary size
        // Important: WPF Width/Height are in device-independent pixels and are DPI-aware
        // DO NOT apply DPI scaling - WPF handles this automatically
        if (!double.IsNaN(_settings.WindowWidth) && _settings.WindowWidth > 0)
            Width = _settings.WindowWidth;
        else
            Width = 1600;

        if (!double.IsNaN(_settings.WindowHeight) && _settings.WindowHeight > 0)
            Height = _settings.WindowHeight;
        else
            Height = 1400;

        // Restore position if previously saved
        if (!double.IsNaN(_settings.WindowLeft) && !double.IsNaN(_settings.WindowTop))
        {
            Left = _settings.WindowLeft;
            Top = _settings.WindowTop;
        }

        // Restore maximized state
        if (_settings.IsMaximized)
            WindowState = WindowState.Maximized;
    }

    /// <summary>
    /// Save window state when the window is closing, so position and size persist
    /// between application sessions.
    /// </summary>
    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        // Persist current window state before closing
        if (WindowState == WindowState.Normal)
        {
            _settings.WindowLeft = Left;
            _settings.WindowTop = Top;
            _settings.WindowWidth = Width;
            _settings.WindowHeight = Height;
            _settings.IsMaximized = false;
        }
        else if (WindowState == WindowState.Maximized)
        {
            _settings.IsMaximized = true;
            // Store the restore bounds (pre-maximized size) so user can un-maximize sensibly
            _settings.WindowWidth = RestoreBounds.Width;
            _settings.WindowHeight = RestoreBounds.Height;
        }

        SettingsService.Save(_settings);
        base.OnClosing(e);
    }

    // ── Add Rule chevron dropdown ────────────────────────────────────────────

    private void AddRuleChevronBtn_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement btn &&
            Resources["AddOpMenu"] is ContextMenu menu)
        {
            menu.DataContext = _vm;
            menu.PlacementTarget = btn;
            menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            menu.IsOpen = true;
        }
    }

    private void AddRuleBtn_Click(object sender, RoutedEventArgs e)
    {
        var menu = AddRuleBtn.ContextMenu;
        if (menu != null)
        {
            // Ensure the context menu has the view model as its DataContext
            menu.DataContext = _vm;
            menu.PlacementTarget = AddRuleBtn;
            menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            menu.IsOpen = true;
        }
    }

    // ── Drag & Drop ──────────────────────────────────────────────────────────

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
            _vm.HandleDrop(paths);
    }

    private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {

    }
}
