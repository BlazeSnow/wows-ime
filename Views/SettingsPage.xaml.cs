using wows_ime.Services;

namespace wows_ime.Views;

public sealed partial class SettingsPage : Page
{
    private static readonly Uri ProjectWebsiteUri = new("https://www.blazesnow.com/wows/");
    private static readonly Uri ProjectRepositoryUri = new("https://github.com/BlazeSnow/wows-ime");

    private bool suppressSelectionChange;

    public string AppVersionText { get; } = GetPackageVersionText();

    public SettingsPage()
    {
        InitializeComponent();
        ApplyLocalization();
        SelectLanguage(GetSavedLanguage());
    }

    private void ApplyLocalization()
    {
        LanguageCard.Header = new TextBlock { Text = SR("Settings/LanguageLabel") };
        AutomaticLanguageItem.Content = SR("Settings/Automatic");

        ProjectWebsiteCard.Header = new TextBlock { Text = SR("ProjectWebsiteCardHeader/Text") };
        ProjectWebsiteCard.Description = "https://www.blazesnow.com/wows/";
        OpenProjectWebsiteButton.Content = SR("OpenProjectWebsiteButton/Content");

        ProjectRepositoryCard.Header = new TextBlock { Text = SR("ProjectRepositoryCardHeader/Text") };
        ProjectRepositoryCard.Description = "https://github.com/BlazeSnow/wows-ime";
        OpenProjectRepositoryButton.Content = SR("OpenProjectRepositoryButton/Content");

        VersionCard.Header = new TextBlock { Text = SR("AppVersionCardHeader/Text") };
    }

    private void SelectLanguage(string? language)
    {
        suppressSelectionChange = true;
        LanguageComboBox.SelectedIndex = language switch
        {
            "zh-Hans" => 1,
            "zh-Hant" => 2,
            "ja" => 3,
            _ => 0
        };
        suppressSelectionChange = false;
    }

    private static string? GetSavedLanguage()
    {
        return Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride?.Trim() is { Length: > 0 } lang ? lang : null;
    }

    private async void OnLanguageSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (suppressSelectionChange || LanguageComboBox.SelectedItem is not ComboBoxItem { Tag: string tag })
        {
            return;
        }

        var current = GetSavedLanguage() ?? "";
        if (tag == current)
        {
            return;
        }

        Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = tag;

        if (await ShowLanguageRestartDialogAsync())
        {
            RestartApp();
        }
    }

    private async Task<bool> ShowLanguageRestartDialogAsync()
    {
        var dialog = new ContentDialog
        {
            Title = SR("Settings/Restart/Title"),
            Content = SR("Settings/Restart/Content"),
            PrimaryButtonText = SR("Settings/Restart/Primary"),
            CloseButtonText = SR("Settings/Restart/Close"),
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = XamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
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

    private static void RestartApp()
    {
        var exePath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(exePath))
        {
            _ = System.Diagnostics.Process.Start(exePath);
        }

        Application.Current.Exit();
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
