using wows_ime.Services;

namespace wows_ime.Views;

public sealed class InputMethodItem : Microsoft.UI.Xaml.DependencyObject
{
    public InputMethodItem(string displayName, ImeCategory initialCategory = ImeCategory.ChineseSimplified, bool isCustom = false)
    {
        DisplayName = displayName;
        Category = initialCategory;
        CategoryIndex = (int)initialCategory;
        IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public bool IsCustom { get; }
    public int ConfirmationNumber { get; set; }
    public string ConfirmationNumberText => ConfirmationNumber.ToString();
    public string CategoryDisplayName => Category switch
    {
        ImeCategory.ChineseTraditional => AppResources.Get("Category/ChineseTraditional"),
        ImeCategory.Japanese => AppResources.Get("Category/Japanese"),
        _ => AppResources.Get("Category/ChineseSimplified")
    };
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
            typeof(InputMethodItem),
            new PropertyMetadata(false));

    public int CategoryIndex
    {
        get => (int)GetValue(CategoryIndexProperty);
        set
        {
            SetValue(CategoryIndexProperty, value);
            Category = value switch
            {
                1 => ImeCategory.ChineseTraditional,
                2 => ImeCategory.Japanese,
                _ => ImeCategory.ChineseSimplified
            };
        }
    }

    public static readonly Microsoft.UI.Xaml.DependencyProperty CategoryIndexProperty =
        Microsoft.UI.Xaml.DependencyProperty.Register(
            nameof(CategoryIndex),
            typeof(int),
            typeof(InputMethodItem),
            new PropertyMetadata(0));

    public ImeCategory Category { get; private set; }
}
