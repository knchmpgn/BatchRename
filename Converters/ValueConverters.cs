using System.Globalization;
using System.Windows;
using System.Windows.Data;
using BatchRename.Models;

// Pin to WPF's Binding — guards against any future WinForms reference creeping in
using Binding = System.Windows.Data.Binding;

namespace BatchRename.Converters;

// ─── InverseBool (bool → bool) ────────────────────────────────────────────────

[ValueConversion(typeof(bool), typeof(bool))]
public sealed class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is true ? false : (object)true;
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => value is true ? false : (object)true;
}

// ─── BoolToVisibility ─────────────────────────────────────────────────────────

[ValueConversion(typeof(bool), typeof(Visibility))]
public sealed class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is true ? Visibility.Visible : Visibility.Collapsed;
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => value is Visibility.Visible;
}

[ValueConversion(typeof(bool), typeof(Visibility))]
public sealed class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is true ? Visibility.Collapsed : Visibility.Visible;
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => value is not Visibility.Visible;
}

// ─── OperationType == Param → Visibility ──────────────────────────────────────

/// <summary>
/// Returns Visible when the bound OperationType equals the converter parameter.
/// Usage: Visibility="{Binding Type, Converter={StaticResource TypeEqVis}, ConverterParameter=FindReplace}"
/// </summary>
[ValueConversion(typeof(OperationType), typeof(Visibility))]
public sealed class OperationTypeToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
    {
        if (value is OperationType actual && p is string param
            && Enum.TryParse<OperationType>(param, out var expected))
            return actual == expected ? Visibility.Visible : Visibility.Collapsed;
        return Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => Binding.DoNothing;
}

// ─── InsertPosition == Param → bool (for RadioButton) ────────────────────────

public sealed class EnumToBoolConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value?.ToString() == p?.ToString();
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
    {
        if (value is true && p is string s)
        {
            // Try OperationType, InsertPosition, etc.
            foreach (var enumType in new[] {
                typeof(InsertPosition), typeof(RemoveAnchor),
                typeof(SequencePlace),  typeof(CaseStyle),
                typeof(SpaceReplacement)})
            {
                if (Enum.TryParse(enumType, s, out var v)) return v;
            }
        }
        return Binding.DoNothing;
    }
}

// ─── Int → string passthrough for NumberBox compatibility ────────────────────

public sealed class IntToStringConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value?.ToString() ?? "0";
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => int.TryParse(value?.ToString(), out var i) ? i : 0;
}

// ─── HasChanges count → status string ────────────────────────────────────────

public sealed class CountToStringConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is int i ? (i == 1 ? "1 file" : $"{i} files") : "0 files";
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => Binding.DoNothing;
}

// ─── String.IsNullOrEmpty → Visibility ───────────────────────────────────────

public sealed class StringEmptyToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => string.IsNullOrEmpty(value?.ToString()) ? Visibility.Visible : Visibility.Collapsed;
    public object ConvertBack(object value, Type t, object p, CultureInfo c)
        => Binding.DoNothing;
}

// ─── NewName == OldName → styling helper ─────────────────────────────────────

public sealed class StringEqualityConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type t, object p, CultureInfo c)
        => values.Length == 2 && values[0]?.ToString() == values[1]?.ToString();
    public object[] ConvertBack(object value, Type[] t, object p, CultureInfo c)
        => Array.Empty<object>();
}
