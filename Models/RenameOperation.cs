using CommunityToolkit.Mvvm.ComponentModel;

namespace BatchRename.Models;

// ─── Enumerations ─────────────────────────────────────────────────────────────

public enum OperationType
{
    FindReplace    = 0,
    InsertText     = 1,
    RemoveRange    = 2,
    AddSequence    = 3,
    ChangeCase     = 4,
    TrimSpaces     = 5,
    ChangeExtension = 6,
    RemoveNumbers  = 7,
    ReplaceSpaces  = 8,
}

public enum InsertPosition    { Beginning, End, AtIndex }
public enum RemoveAnchor      { FromStart, FromEnd }
public enum SequencePlace     { Before, After, Replace }
public enum CaseStyle         { Uppercase, Lowercase, TitleCase, SentenceCase }
public enum SpaceReplacement  { Underscore, Hyphen, Period, Remove }

// ─── Operation model ──────────────────────────────────────────────────────────

public partial class RenameOperation : ObservableObject
{
    // ── Common ──────────────────────────────────────────────────────────────

    [ObservableProperty] private OperationType _type      = OperationType.FindReplace;
    [ObservableProperty] private bool          _isEnabled = true;
    [ObservableProperty] private bool          _isExpanded = false;

    // ── Find & Replace ──────────────────────────────────────────────────────

    [ObservableProperty] private string _findText    = string.Empty;
    [ObservableProperty] private string _replaceText = string.Empty;
    [ObservableProperty] private bool   _matchCase   = false;
    [ObservableProperty] private bool   _useRegex    = false;
    [ObservableProperty] private bool   _findApplyToExtension = false;

    // ── Insert Text ─────────────────────────────────────────────────────────

    [ObservableProperty] private string         _insertText     = string.Empty;
    [ObservableProperty] private InsertPosition _insertPosition = InsertPosition.Beginning;
    [ObservableProperty] private double         _insertIndex    = 0;
    [ObservableProperty] private bool           _insertApplyToExtension = false;

    // ── Remove Range ────────────────────────────────────────────────────────

    [ObservableProperty] private RemoveAnchor _removeAnchor = RemoveAnchor.FromStart;
    [ObservableProperty] private double       _removeFrom   = 0;
    [ObservableProperty] private double       _removeCount  = 1;

    // ── Add Sequence ────────────────────────────────────────────────────────

    [ObservableProperty] private string        _seqPrefix  = string.Empty;
    [ObservableProperty] private string        _seqSuffix  = string.Empty;
    [ObservableProperty] private double        _seqStart   = 1;
    [ObservableProperty] private double        _seqStep    = 1;
    [ObservableProperty] private double        _seqPadding = 2;
    [ObservableProperty] private SequencePlace _seqPlace   = SequencePlace.After;

    // ── Change Case ─────────────────────────────────────────────────────────

    [ObservableProperty] private CaseStyle _caseStyle            = CaseStyle.Lowercase;
    [ObservableProperty] private bool      _caseApplyToExtension = false;

    // ── Change Extension ────────────────────────────────────────────────────

    [ObservableProperty] private string _newExtension = string.Empty;

    // ── Replace Spaces ──────────────────────────────────────────────────────

    [ObservableProperty] private SpaceReplacement _spaceReplacement = SpaceReplacement.Underscore;

    // ── Property-change hooks → refresh Description ──────────────────────────
    // CommunityToolkit.Mvvm generates On{Prop}Changed partial methods automatically.

    partial void OnTypeChanged(OperationType value)               => OnPropertyChanged(nameof(Description));
    partial void OnFindTextChanged(string value)                  => OnPropertyChanged(nameof(Description));
    partial void OnReplaceTextChanged(string value)               => OnPropertyChanged(nameof(Description));
    partial void OnMatchCaseChanged(bool value)                   => OnPropertyChanged(nameof(Description));
    partial void OnUseRegexChanged(bool value)                    => OnPropertyChanged(nameof(Description));
    partial void OnInsertTextChanged(string value)                => OnPropertyChanged(nameof(Description));
    partial void OnInsertPositionChanged(InsertPosition value)    => OnPropertyChanged(nameof(Description));
    partial void OnInsertIndexChanged(double value)               => OnPropertyChanged(nameof(Description));
    partial void OnRemoveAnchorChanged(RemoveAnchor value)        => OnPropertyChanged(nameof(Description));
    partial void OnRemoveFromChanged(double value)                => OnPropertyChanged(nameof(Description));
    partial void OnRemoveCountChanged(double value)               => OnPropertyChanged(nameof(Description));
    partial void OnSeqPrefixChanged(string value)                 => OnPropertyChanged(nameof(Description));
    partial void OnSeqSuffixChanged(string value)                 => OnPropertyChanged(nameof(Description));
    partial void OnSeqStartChanged(double value)                  => OnPropertyChanged(nameof(Description));
    partial void OnSeqStepChanged(double value)                   => OnPropertyChanged(nameof(Description));
    partial void OnSeqPaddingChanged(double value)                => OnPropertyChanged(nameof(Description));
    partial void OnSeqPlaceChanged(SequencePlace value)           => OnPropertyChanged(nameof(Description));
    partial void OnCaseStyleChanged(CaseStyle value)              => OnPropertyChanged(nameof(Description));
    partial void OnNewExtensionChanged(string value)              => OnPropertyChanged(nameof(Description));
    partial void OnSpaceReplacementChanged(SpaceReplacement value)=> OnPropertyChanged(nameof(Description));

    // ── Display helpers ──────────────────────────────────────────────────────

    public string DisplayTitle => Type switch
    {
        OperationType.FindReplace     => "Find & Replace",
        OperationType.InsertText      => "Insert Text",
        OperationType.RemoveRange     => "Remove Characters",
        OperationType.AddSequence     => "Add Sequence",
        OperationType.ChangeCase      => "Change Case",
        OperationType.TrimSpaces      => "Trim Whitespace",
        OperationType.ChangeExtension => "Change Extension",
        OperationType.RemoveNumbers   => "Remove Numbers",
        OperationType.ReplaceSpaces   => "Replace Spaces",
        _                             => Type.ToString(),
    };

    /// <summary>Human-readable summary of the current configuration.</summary>
    public string Description => Type switch
    {
        OperationType.FindReplace =>
            string.IsNullOrEmpty(FindText)
                ? "Set a search term to get started"
                : $"Find \"{FindText}\"" +
                  (string.IsNullOrEmpty(ReplaceText) ? " → delete" : $" → \"{ReplaceText}\"") +
                  (MatchCase ? "  ·  Match case" : "") +
                  (UseRegex  ? "  ·  Regex"      : ""),

        OperationType.InsertText =>
            string.IsNullOrEmpty(InsertText)
                ? "Set text to insert"
                : InsertPosition switch
                {
                    InsertPosition.Beginning => $"Prepend \"{InsertText}\"",
                    InsertPosition.End       => $"Append \"{InsertText}\"",
                    InsertPosition.AtIndex   => $"Insert \"{InsertText}\" at index {(int)InsertIndex}",
                    _                        => InsertText,
                },

        OperationType.RemoveRange =>
            $"Remove {(int)RemoveCount} char(s) from position {(int)RemoveFrom}" +
            (RemoveAnchor == RemoveAnchor.FromEnd ? " (from end)" : " (from start)"),

        OperationType.AddSequence =>
            $"{SeqPlace.ToString().ToLower()} name · " +
            $"{SeqPrefix}[{(int)SeqStart},{(int)SeqStart + (int)SeqStep},…]{SeqSuffix}" +
            $" · {(int)SeqPadding}-digit pad",

        OperationType.ChangeCase => CaseStyle switch
        {
            CaseStyle.Uppercase    => "Convert to UPPERCASE",
            CaseStyle.Lowercase    => "Convert to lowercase",
            CaseStyle.TitleCase    => "Convert to Title Case",
            CaseStyle.SentenceCase => "Convert to Sentence case",
            _                      => "Change case",
        },

        OperationType.TrimSpaces      => "Strip leading/trailing spaces, collapse internal runs",
        OperationType.RemoveNumbers   => "Remove all numeric digits (0–9)",

        OperationType.ChangeExtension =>
            string.IsNullOrWhiteSpace(NewExtension)
                ? "Set a new extension"
                : $"Change extension to {(NewExtension.StartsWith('.') ? NewExtension : "." + NewExtension)}",

        OperationType.ReplaceSpaces => SpaceReplacement switch
        {
            SpaceReplacement.Underscore => "Replace spaces with underscore  _",
            SpaceReplacement.Hyphen     => "Replace spaces with hyphen  -",
            SpaceReplacement.Period     => "Replace spaces with period  .",
            SpaceReplacement.Remove     => "Remove all spaces",
            _                           => "Replace spaces",
        },

        _ => DisplayTitle,
    };

}
