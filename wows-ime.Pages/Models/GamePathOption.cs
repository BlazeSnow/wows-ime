using wows_ime.Pages.Abstractions;

namespace wows_ime.Pages.Models;

public sealed class GamePathOption : Microsoft.UI.Xaml.DependencyObject
{
    private readonly IPageLocalization localization;

    public GamePathOption(string displayName, string path, IPageLocalization localization, bool isCustom = false)
    {
        DisplayName = displayName;
        Path = path;
        this.localization = localization;
        IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public string Path { get; }
    public bool IsCustom { get; }
    public string DeleteButtonContent => localization.GetString("DeleteCustomGamePathButtonTemplate/Content");
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
