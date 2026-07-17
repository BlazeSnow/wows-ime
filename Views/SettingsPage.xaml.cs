using Windows.ApplicationModel.Resources.Core;
using wows_ime.Services;

namespace wows_ime.Views;

public sealed partial class SettingsPage : Page
{
    private static readonly Uri ProjectWebsiteUri = new("https://www.blazesnow.com/wows/");
    private static readonly Uri ProjectRepositoryUri = new("https://github.com/BlazeSnow/wows-ime");

    private static readonly List<(string Tag, string DisplayName)> LanguageOptions = new()
    {
        ("", "自动"),
        ("zh-Hans", "简体中文"),
        ("zh-Hant", "繁體中文"),
        ("ja", "日本語")
    };

    private bool suppressSelectionChange;
    private string currentLanguageTag = "";

    public string AppVersionText { get; } = GetPackageVersionText();

    public SettingsPage()
    {
        InitializeComponent();
        LanguageLabel.Text = SR("Settings/LanguageLabel");
        InitializeLanguageCombo();
    }

    private void InitializeLanguageCombo()
    {
        var currentOverride = Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride;
        currentLanguageTag = string.IsNullOrEmpty(currentOverride) ? "" : currentOverride;

        suppressSelectionChange = true;
        foreach (var (tag, displayName) in LanguageOptions)
        {
            var item = new ComboBoxItem { Content = displayName, Tag = tag };
            LanguageComboBox.Items.Add(item);

            if (tag == currentLanguageTag)
            {
                LanguageComboBox.SelectedItem = item;
            }
        }
        suppressSelectionChange = false;
    }

    private async void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (suppressSelectionChange || LanguageComboBox.SelectedItem is not ComboBoxItem { Tag: string tag })
        {
            return;
        }

        if (tag == currentLanguageTag)
        {
            return;
        }

        var restart = await ShowRestartDialogAsync(tag);
        if (!restart)
        {
            // Reset to previous selection
            suppressSelectionChange = true;
            foreach (ComboBoxItem item in LanguageComboBox.Items)
            {
                if (item.Tag is string itemTag && itemTag == currentLanguageTag)
                {
                    LanguageComboBox.SelectedItem = item;
                    break;
                }
            }
            suppressSelectionChange = false;
            return;
        }

        // Apply language override and restart
        Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = tag;
        Microsoft.Windows.AppLifecycle.AppInstance.Restart("");
    }

    private async Task<bool> ShowRestartDialogAsync(string targetLanguageTag)
    {
        var title = GetStringForLanguage("Settings/Restart/Title", targetLanguageTag);
        var content = GetStringForLanguage("Settings/Restart/Content", targetLanguageTag);
        var primary = GetStringForLanguage("Settings/Restart/Primary", targetLanguageTag);
        var close = GetStringForLanguage("Settings/Restart/Close", targetLanguageTag);

        var dialog = new ContentDialog
        {
            Title = title,
            Content = content,
            PrimaryButtonText = primary,
            CloseButtonText = close,
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = XamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    private static string GetStringForLanguage(string key, string languageTag)
    {
        try
        {
            var resourceManager = ResourceManager.Current;
            var resourceMap = resourceManager.MainResourceMap.GetSubtree("Resources");
            var context = new ResourceContext();
            context.QualifierValues["Language"] = languageTag;
            var candidate = resourceMap.GetValue(key, context);
            if (candidate is not null)
            {
                return candidate.ValueAsString;
            }
        }
        catch
        {
            // Fall through
        }

        return key;
    }

    private async void OpenProjectWebsiteButton_Click(object sender, RoutedEventArgs e)
    {
        var launched = await Windows.System.Launcher.LaunchUriAsync(ProjectWebsiteUri);
        if (!launched)
        {
            await ShowErrorAsync(SR("Status/OpenProjectWebsiteFailed"));
        }
    }

    private async void OpenProjectRepositoryButton_Click(object sender, RoutedEventArgs e)
    {
        var launched = await Windows.System.Launcher.LaunchUriAsync(ProjectRepositoryUri);
        if (!launched)
        {
            await ShowErrorAsync(SR("Status/OpenProjectRepositoryFailed"));
        }
    }

    private async Task ShowErrorAsync(string message)
    {
        var dialog = new ContentDialog
        {
            Title = message,
            CloseButtonText = "OK",
            XamlRoot = XamlRoot
        };
        await dialog.ShowAsync();
    }

    private static string SR(string key)
    {
        return AppResources.Get(key);
    }

    private static string GetPackageVersionText()
    {
        try
        {
            var version = Windows.ApplicationModel.Package.Current.Id.Version;
            return $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
        }
        catch
        {
            return string.Empty;
        }
    }
}
