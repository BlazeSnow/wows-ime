namespace wows_ime.Views;

public sealed class GamePathOption : Microsoft.UI.Xaml.DependencyObject
{
    public GamePathOption(string displayName, string path, bool isCustom = false)
    {
        DisplayName = displayName;
        Path = path;
        IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public string Path { get; }
    public bool IsCustom { get; }
    public Visibility DeleteButtonVisibility => IsCustom ? Visibility.Visible : Visibility.Collapsed;

    public bool IsSelected
    {
        get => (bool)GetValue(IsSelectedProperty);
        set => SetValue(IsSelectedProperty, value);
    }

    public static readonly Microsoft.UI.Xaml.DependencyProperty IsSelectedProperty =
        Microsoft.UI.Xaml.DependencyProperty.Register(
            nameof(IsSelected),
            typeof(bool),
            typeof(GamePathOption),
            new PropertyMetadata(false));
}
