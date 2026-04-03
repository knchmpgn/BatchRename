using CommunityToolkit.Mvvm.ComponentModel;

namespace BatchRename.Models;

/// <summary>
/// Represents a single file queued for renaming.
/// Exposes the original name and a live preview of the new name.
/// </summary>
public partial class FileEntry : ObservableObject
{
    /// <summary>Full path to the original file on disk.</summary>
    public string OriginalPath { get; init; } = string.Empty;

    /// <summary>Directory that contains the file.</summary>
    public string Directory => System.IO.Path.GetDirectoryName(OriginalPath) ?? string.Empty;

    /// <summary>Original file name (name + extension).</summary>
    public string OriginalName { get; init; } = string.Empty;

    /// <summary>Live preview of the name after all rename operations are applied.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasChanges))]
    [NotifyPropertyChangedFor(nameof(StatusGlyph))]
    [NotifyPropertyChangedFor(nameof(StatusBrushKey))]
    private string _newName = string.Empty;

    /// <summary>True when the preview name differs from the original.</summary>
    public bool HasChanges => NewName != OriginalName;

    /// <summary>
    /// Fluent System Icons glyph for the status column.
    /// ✓ = will be renamed, = = no change, ⚠ = conflict.
    /// </summary>
    public string StatusGlyph => HasChanges ? "\uE73E" : "\uE8D9"; // Checkmark : SkypeEquals

    /// <summary>Brush resource key used to colour the status glyph.</summary>
    public string StatusBrushKey => HasChanges ? "SystemFillColorSuccessBrush" : "TextFillColorTertiaryBrush";

    /// <summary>File size string for display.</summary>
    public string FileSize { get; init; } = string.Empty;

    /// <summary>File extension (lower-case, including leading dot).</summary>
    public string Extension => System.IO.Path.GetExtension(OriginalName).ToLowerInvariant();
}
