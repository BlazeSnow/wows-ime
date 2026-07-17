using wows_ime.Services;

namespace wows_ime.Views;

public sealed partial class SettingsPage : Page
{
    private static readonly Uri ProjectWebsiteUri = new("https://www.blazesnow.com/wows/");
    private static readonly Uri ProjectRepositoryUri = new("https://github.com/BlazeSnow/wows-ime");

    private static readonly Dictionary<string, (string Title, string Content, string Primary, string Close)> RestartStrings = new()
    {
        ["zh-Hans"] = ("重启应用", "语言更改将在重启后生效，是否立即重启？", "立即重启", "稍后"),
        ["zh-Hant"] = ("重啟應用", "語言更改將在重啟後生效，是否立即重啟？", "立即重啟", "稍後"),
        ["ja"]      = ("アプリの再起動", "言語の変更は再起動後に反映されます。今すぐ再起動しますか？", "今すぐ再起動", "後で"),
    };

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
        LanguageCard.Header = SR("Settings/LanguageLabel");
        AutomaticLanguageItem.Content = "自动";

        ProjectWebsiteCard.Header = SR("ProjectWebsiteCardHeader.Text");
        ProjectWebsiteCard.Description = "https://www.blazesnow.com/wows/";
        OpenProjectWebsiteButton.Content = SR("OpenProjectWebsiteButton.Content");

        ProjectRepositoryCard.Header = SR("ProjectRepositoryCardHeader.Text");
        ProjectRepositoryCard.Description = "https://github.com/BlazeSnow/wows-ime";
        OpenProjectRepositoryButton.Content = SR("OpenProjectRepositoryButton.Content");

        VersionCard.Header = SR("AppVersionCardHeader.Text");
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

        var restart = await ShowLanguageRestartDialogAsync(tag);
        if (!restart)
        {
            SelectLanguage(current.Length > 0 ? current : null);
            return;
        }

        Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = tag;
        RestartApp();
    }

    private async Task<bool> ShowLanguageRestartDialogAsync(string languageTag)
    {
        var title = "重启应用";
        var content = "语言更改将在重启后生效，是否立即重启？";
        var primary = "立即重启";
        var close = "稍后";

        if (RestartStrings.TryGetValue(languageTag, out var strings))
        {
            (title, content, primary, close) = strings;
        }

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
