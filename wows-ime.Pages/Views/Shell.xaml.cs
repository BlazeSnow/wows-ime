using Microsoft.UI.Xaml.Media;
using wows_ime.Pages.Abstractions;

namespace wows_ime.Pages.Views;

public sealed partial class Shell : UserControl
{
    private readonly IPageHost host;

    public Shell(IPageHost host)
    {
        this.host = host ?? throw new ArgumentNullException(nameof(host));
        InitializeComponent();
        AppTitleBar.Title = SR("App/Title");
        NavHomeItem.Content = SR("Nav/Home");
        NavSettingsItem.Content = SR("Nav/Settings");
        _ = ContentFrame.Navigate(typeof(HomePage), host);
        AppNavView.SelectedItem = AppNavView.MenuItems[0];
    }

    private void AppTitleBar_Loaded(object sender, RoutedEventArgs e)
    {
        DispatcherQueue.TryEnqueue(() =>
        {
            var paneButton = FindPaneToggleButton(AppTitleBar);
            if (paneButton is not null)
            {
                paneButton.Click += (_, _) => AppNavView.IsPaneOpen = !AppNavView.IsPaneOpen;
            }
        });
    }

    private void AppNavView_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is not NavigationViewItem { Tag: string tag })
        {
            return;
        }

        var pageType = tag switch { "settings" => typeof(SettingsPage), _ => typeof(HomePage) };
        _ = ContentFrame.Navigate(pageType, host);
    }

    private void ContentFrame_NavigationFailed(object sender, Microsoft.UI.Xaml.Navigation.NavigationFailedEventArgs e)
    {
        throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
    }

    private static Button? FindPaneToggleButton(DependencyObject parent)
    {
        var count = VisualTreeHelper.GetChildrenCount(parent);
        for (var i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is Button button)
            {
                return button;
            }

            var descendant = FindPaneToggleButton(child);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }

    private string SR(string key) => host.Localization.GetString(key);
}
