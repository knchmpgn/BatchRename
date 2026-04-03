using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using BatchRename.Models;
using Microsoft.Win32;

namespace BatchRename.ViewModels;

public partial class MainViewModel : ObservableObject
{
    // ── Collections ─────────────────────────────────────────────────────────

    public ObservableCollection<FileEntry>       Files      { get; } = new();
    public ObservableCollection<RenameOperation> Operations { get; } = new();

    // ── Observable state ────────────────────────────────────────────────────

    [ObservableProperty] private string _statusText     = "No files added. Drag & drop files or click Add Files.";
    [ObservableProperty] private bool   _isBusy         = false;
    [ObservableProperty] private string _successMessage = string.Empty;
    [ObservableProperty] private bool   _showSuccess    = false;

    public int  ChangedCount  => Files.Count(f => f.HasChanges);
    public int  TotalCount    => Files.Count;
    public bool HasFiles      => Files.Count > 0;
    public bool HasOperations => Operations.Count > 0;

    // ── Constructor ─────────────────────────────────────────────────────────

    public MainViewModel()
    {
        Files.CollectionChanged      += OnFilesChanged;
        Operations.CollectionChanged += OnOperationsChanged;
    }

    // ── Collection change handlers ──────────────────────────────────────────

    private void OnFilesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        UpdatePreviews();
        OnPropertyChanged(nameof(TotalCount));
        OnPropertyChanged(nameof(ChangedCount));
        OnPropertyChanged(nameof(HasFiles));
        RefreshStatus();
    }

    private void OnOperationsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems != null)
            foreach (RenameOperation op in e.NewItems)
                op.PropertyChanged += OnOperationPropertyChanged;

        if (e.OldItems != null)
            foreach (RenameOperation op in e.OldItems)
                op.PropertyChanged -= OnOperationPropertyChanged;

        OnPropertyChanged(nameof(HasOperations));
        UpdatePreviews();
    }

    private void OnOperationPropertyChanged(object? sender, PropertyChangedEventArgs e)
        => UpdatePreviews();

    // ── File commands ────────────────────────────────────────────────────────

    [RelayCommand]
    private void AddFiles()
    {
        var dlg = new OpenFileDialog
        {
            Title       = "Select files to rename",
            Multiselect = true,
            Filter      = "All files (*.*)|*.*",
        };

        if (dlg.ShowDialog() == true)
            AddFilePaths(dlg.FileNames);
    }

    [RelayCommand]
    private void AddFolder()
    {
        // OpenFolderDialog is built into WPF / Microsoft.Win32 since .NET 8
        var dlg = new OpenFolderDialog
        {
            Title      = "Select a folder — all files inside will be added",
            Multiselect = false,
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                var paths = Directory.GetFiles(dlg.FolderName, "*.*", SearchOption.TopDirectoryOnly);
                AddFilePaths(paths);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not read folder: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    [RelayCommand]
    private void ClearFiles()
    {
        Files.Clear();
        RefreshStatus();
    }

    [RelayCommand]
    private void RemoveFile(FileEntry? file)
    {
        if (file != null) Files.Remove(file);
    }

    // ── Operation commands ───────────────────────────────────────────────────

    [RelayCommand]
    private void AddOperation(string? typeParam)
    {
        var type = typeParam != null && Enum.TryParse<OperationType>(typeParam, out var t)
            ? t : OperationType.FindReplace;
        Operations.Add(new RenameOperation { Type = type });
    }

    [RelayCommand]
    private void ToggleOperationExpanded(RenameOperation? op)
    {
        if (op != null) op.IsExpanded = !op.IsExpanded;
    }

    [RelayCommand]
    private void RemoveOperation(RenameOperation? op)
    {
        if (op != null) Operations.Remove(op);
    }

    [RelayCommand]
    private void MoveOperationUp(RenameOperation? op)
    {
        if (op == null) return;
        int i = Operations.IndexOf(op);
        if (i > 0) { op.IsExpanded = false; Operations.Move(i, i - 1); }
    }

    [RelayCommand]
    private void MoveOperationDown(RenameOperation? op)
    {
        if (op == null) return;
        int i = Operations.IndexOf(op);
        if (i >= 0 && i < Operations.Count - 1) { op.IsExpanded = false; Operations.Move(i, i + 1); }
    }

    [RelayCommand]
    private void ClearOperations() => Operations.Clear();

    // ── Apply command ────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task ApplyRenames()
    {
        var toRename = Files.Where(f => f.HasChanges).ToList();
        if (!toRename.Any())
        {
            StatusText = "Nothing to rename — no changes detected.";
            return;
        }

        var targetNames = toRename.Select(f => Path.Combine(f.Directory, f.NewName)).ToList();
        var duplicates  = targetNames.GroupBy(p => p).Where(g => g.Count() > 1).ToList();
        if (duplicates.Any())
        {
            MessageBox.Show(
                $"Cannot apply: {duplicates.Count} duplicate target name(s) detected.\nResolve conflicts before renaming.",
                "Name Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        IsBusy = true;
        int success = 0, failed = 0;

        await Task.Run(() =>
        {
            foreach (var file in toRename)
            {
                try
                {
                    var dest = Path.Combine(file.Directory, file.NewName);
                    if (!File.Exists(dest))
                    {
                        File.Move(file.OriginalPath, dest);
                        success++;
                    }
                    else { failed++; }
                }
                catch { failed++; }
            }
        });

        IsBusy = false;

        if (failed == 0)
        {
            SuccessMessage = $"✓  {success} file{(success == 1 ? "" : "s")} renamed successfully.";
            ShowSuccess = true;
            Files.Clear();
            _ = HideSuccessAfterDelay();
        }
        else
        {
            StatusText = $"Completed: {success} renamed, {failed} failed.";
            var done = toRename.Where(f => !File.Exists(f.OriginalPath)).ToList();
            foreach (var f in done) Files.Remove(f);
            UpdatePreviews();
        }

        OnPropertyChanged(nameof(ChangedCount));
        OnPropertyChanged(nameof(TotalCount));
        RefreshStatus();
    }

    [RelayCommand]
    private void ResetPreviews()
    {
        Operations.Clear();
        UpdatePreviews();
    }

    // ── Drag & Drop ──────────────────────────────────────────────────────────

    public void HandleDrop(string[] paths)
    {
        var files   = paths.Where(File.Exists).ToArray();
        var folders = paths.Where(Directory.Exists).ToArray();
        AddFilePaths(files);
        foreach (var folder in folders)
        {
            try { AddFilePaths(Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)); }
            catch { }
        }
    }

    // ── Core preview engine ──────────────────────────────────────────────────

    private void UpdatePreviews()
    {
        for (int i = 0; i < Files.Count; i++)
            Files[i].NewName = ComputeNewName(Files[i].OriginalName, i);
        OnPropertyChanged(nameof(ChangedCount));
    }

    private string ComputeNewName(string originalName, int index)
    {
        if (!Operations.Any(o => o.IsEnabled)) return originalName;

        string name = Path.GetFileNameWithoutExtension(originalName);
        string ext  = Path.GetExtension(originalName);

        foreach (var op in Operations.Where(o => o.IsEnabled))
        {
            try { (name, ext) = ApplyOperation(op, name, ext, index); }
            catch { }
        }

        return name + ext;
    }

    private static (string name, string ext) ApplyOperation(
        RenameOperation op, string name, string ext, int index)
    {
        switch (op.Type)
        {
            case OperationType.FindReplace:
            {
                if (string.IsNullOrEmpty(op.FindText)) break;
                var opts = op.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                var cmp  = op.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                if (op.UseRegex)
                {
                    name = Regex.Replace(name, op.FindText, op.ReplaceText, opts);
                    if (op.FindApplyToExtension)
                        ext = Regex.Replace(ext, op.FindText, op.ReplaceText, opts);
                }
                else
                {
                    name = name.Replace(op.FindText, op.ReplaceText, cmp);
                    if (op.FindApplyToExtension)
                        ext = ext.Replace(op.FindText, op.ReplaceText, cmp);
                }
                break;
            }

            case OperationType.InsertText:
            {
                int    idx    = Math.Max(0, (int)op.InsertIndex);
                string target = op.InsertApplyToExtension ? (name + ext) : name;
                string result = op.InsertPosition switch
                {
                    InsertPosition.Beginning => op.InsertText + target,
                    InsertPosition.End       => target + op.InsertText,
                    InsertPosition.AtIndex   => idx <= target.Length
                                                  ? target[..idx] + op.InsertText + target[idx..]
                                                  : target + op.InsertText,
                    _ => target,
                };
                if (op.InsertApplyToExtension)
                {
                    name = Path.GetFileNameWithoutExtension(result);
                    ext  = Path.GetExtension(result);
                }
                else { name = result; }
                break;
            }

            case OperationType.RemoveRange:
            {
                int count = Math.Max(0, (int)op.RemoveCount);
                if (count == 0 || name.Length == 0) break;
                int start = op.RemoveAnchor == RemoveAnchor.FromEnd
                    ? Math.Max(0, name.Length - (int)op.RemoveFrom - count)
                    : Math.Min((int)op.RemoveFrom, name.Length - 1);
                start = Math.Max(0, start);
                count = Math.Min(count, name.Length - start);
                if (count > 0) name = name.Remove(start, count);
                break;
            }

            case OperationType.AddSequence:
            {
                int    num   = (int)op.SeqStart + index * (int)op.SeqStep;
                int    pad   = Math.Max(1, (int)op.SeqPadding);
                string token = op.SeqPrefix + num.ToString().PadLeft(pad, '0') + op.SeqSuffix;
                name = op.SeqPlace switch
                {
                    SequencePlace.Before  => token + name,
                    SequencePlace.After   => name + token,
                    SequencePlace.Replace => token,
                    _ => name,
                };
                break;
            }

            case OperationType.ChangeCase:
                name = ApplyCase(name, op.CaseStyle);
                if (op.CaseApplyToExtension && ext.Length > 1)
                    ext = "." + ApplyCase(ext[1..], op.CaseStyle);
                break;

            case OperationType.TrimSpaces:
                name = Regex.Replace(name.Trim(), @"\s{2,}", " ");
                break;

            case OperationType.ChangeExtension:
                if (!string.IsNullOrWhiteSpace(op.NewExtension))
                    ext = op.NewExtension.StartsWith('.') ? op.NewExtension : "." + op.NewExtension;
                break;

            case OperationType.RemoveNumbers:
                name = Regex.Replace(name, @"\d+", string.Empty);
                break;

            case OperationType.ReplaceSpaces:
                string rep = op.SpaceReplacement switch
                {
                    SpaceReplacement.Underscore => "_",
                    SpaceReplacement.Hyphen     => "-",
                    SpaceReplacement.Period     => ".",
                    SpaceReplacement.Remove     => string.Empty,
                    _ => "_",
                };
                name = name.Replace(" ", rep);
                break;
        }

        return (name, ext);
    }

    private static string ApplyCase(string s, CaseStyle style) => style switch
    {
        CaseStyle.Uppercase    => s.ToUpperInvariant(),
        CaseStyle.Lowercase    => s.ToLowerInvariant(),
        CaseStyle.TitleCase    => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()),
        CaseStyle.SentenceCase => s.Length > 0
                                    ? char.ToUpperInvariant(s[0]) + s[1..].ToLowerInvariant()
                                    : s,
        _ => s,
    };

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void AddFilePaths(IEnumerable<string> paths)
    {
        var existing = Files.Select(f => f.OriginalPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var path in paths)
        {
            if (existing.Contains(path)) continue;
            existing.Add(path);

            long size = 0;
            try { size = new FileInfo(path).Length; } catch { }
            string sizeStr = size < 1024 ? $"{size} B"
                           : size < 1024 * 1024 ? $"{size / 1024.0:F1} KB"
                           : $"{size / (1024.0 * 1024):F1} MB";

            Files.Add(new FileEntry
            {
                OriginalPath = path,
                OriginalName = Path.GetFileName(path),
                NewName      = Path.GetFileName(path),
                FileSize     = sizeStr,
            });
        }
        RefreshStatus();
    }

    private void RefreshStatus()
    {
        int total   = Files.Count;
        int changed = ChangedCount;
        StatusText = total == 0
            ? "No files added. Drag & drop files or click Add Files."
            : changed == 0
                ? $"{total} file{(total == 1 ? "" : "s")} loaded — no rename operations yet."
                : $"{total} file{(total == 1 ? "" : "s")} · {changed} will be renamed.";
    }

    private async Task HideSuccessAfterDelay()
    {
        await Task.Delay(4000);
        ShowSuccess = false;
    }
}
