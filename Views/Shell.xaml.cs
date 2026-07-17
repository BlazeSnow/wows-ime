using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.Resources;

namespace wows_ime.Views;

public sealed partial class Shell : UserControl
{
    private static readonly ResourceLoader ResourceLoader = new();

    public Shell()
    {
        InitializeComponent();
        AppTitleBar.Title = SR("App/Title");
        NavHomeItem.Content = SR("Nav/Home");
        NavSettingsItem.Content = SR("Nav/Settings");
        _ = ContentFrame.Navigate(typeof(HomePage));
        AppNavView.SelectedItem = AppNavView.MenuItems[0];
    }

    private void AppTitleBar_Loaded(object sender, RoutedEventArgs e)
    {
        var paneButton = FindVisualChild<Button>(AppTitleBar, "PaneToggleButton");
        if (paneButton is not null)
        {
            paneButton.Click += (_, _) => AppNavView.IsPaneOpen = !AppNavView.IsPaneOpen;
        }
    }

    private void AppNavView_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is NavigationViewItem { Tag: string tag })
        {
            var pageType = tag switch
            {
                "settings" => typeof(SettingsPage),
                _ => typeof(HomePage)
            };
            _ = ContentFrame.Navigate(pageType);
        }
    }

    private void ContentFrame_NavigationFailed(object sender, Microsoft.UI.Xaml.Navigation.NavigationFailedEventArgs e)
    {
        throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
    }

    private static T? FindVisualChild<T>(DependencyObject parent, string name) where T : FrameworkElement
    {
        var count = VisualTreeHelper.GetChildrenCount(parent);
        for (var i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T element && element.Name == name)
            {
                return element;
            }

            var descendant = FindVisualChild<T>(child, name);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }

    private static string SR(string key)
    {
        var value = ResourceLoader.GetString(key);
        return string.IsNullOrEmpty(value) ? key : value;
    }
}
