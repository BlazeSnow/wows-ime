using wows_ime.Core.Models;
using wows_ime.Pages.Abstractions;

namespace wows_ime.Pages.Models;

public sealed class InputMethodItem : Microsoft.UI.Xaml.DependencyObject
{
    private readonly IPageLocalization localization;

    public InputMethodItem(
        string displayName,
        IPageLocalization localization,
        ImeCategory initialCategory = ImeCategory.ChineseSimplified,
        bool isCustom = false)
    {
        DisplayName = displayName;
        this.localization = localization;
        Category = initialCategory;
        CategoryIndex = (int)initialCategory;
        IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public bool IsCustom { get; }
    public int ConfirmationNumber { get; set; }
    public string ConfirmationNumberText => ConfirmationNumber.ToString();
    public string DeleteButtonContent => localization.GetString("DeleteCustomImeButtonTemplate/Content");
    public string ChineseSimplifiedCategoryContent => localization.GetString("ImeCategorySimplifiedItem/Content");
    public string ChineseTraditionalCategoryContent => localization.GetString("ImeCategoryTraditionalItem/Content");
    public string JapaneseCategoryContent => localization.GetString("ImeCategoryJapaneseItem/Content");
    public string CategoryDisplayName => Category switch
    {
        ImeCategory.ChineseTraditional => localization.GetString("Category/ChineseTraditional"),
        ImeCategory.Japanese => localization.GetString("Category/Japanese"),
        _ => localization.GetString("Category/ChineseSimplified")
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
