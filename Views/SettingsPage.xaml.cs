using wows_ime.Services;

namespace wows_ime.Views;

public sealed partial class SettingsPage : Page
{
    private static readonly Uri ProjectWebsiteUri = new("https://www.blazesnow.com/wows/");
    private static readonly Uri ProjectRepositoryUri = new("https://github.com/BlazeSnow/wows-ime");

    public string AppVersionText { get; } = GetPackageVersionText();

    public SettingsPage()
    {
        InitializeComponent();
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
