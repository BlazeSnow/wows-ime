using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.Win32;
using System.Text;
using Windows.ApplicationModel.Resources;
using Windows.Storage;
using Windows.UI.ViewManagement;
using WinRT.Interop;

namespace wows_ime
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    public partial class App : Application
    {
        private Window window = Window.Current;
        private UISettings? uiSettings;
        private static readonly ResourceLoader ResourceLoader = new();
        public static Window? MainWindow { get; private set; }

        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            this.InitializeComponent();
            this.UnhandledException += App_UnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        /// <summary>
        /// Invoked when the application is launched normally by the end user.  Other entry points
        /// will be used such as when the application is launched to open a specific file.
        /// </summary>
        /// <param name="e">Details about the launch request and process.</param>
        protected override void OnLaunched(LaunchActivatedEventArgs e)
        {
            window ??= new Window();
            MainWindow = window;
            window.Title = SR("App/Title");
            window.SystemBackdrop = new MicaBackdrop();
            SetWindowIcon(window);

            CreateWindowContent(window);
            ApplySystemTitleBarTheme(window);
            EnsureThemeListener();

            window.Activate();
        }

        private void CreateWindowContent(Window targetWindow)
        {
            var contentFrame = new Frame();
            contentFrame.NavigationFailed += OnNavigationFailed;

            var navigationView = new NavigationView
            {
                IsBackButtonVisible = NavigationViewBackButtonVisible.Collapsed,
                IsSettingsVisible = false,
                PaneDisplayMode = NavigationViewPaneDisplayMode.Left,
                IsTitleBarAutoPaddingEnabled = false,
                OpenPaneLength = 200,
                Content = contentFrame
            };

            var titleBar = CreateTitleBar(navigationView);

            navigationView.MenuItems.Add(CreateNavItem("Nav/Home", "\uE80F", "home"));
            navigationView.MenuItems.Add(CreateNavItem("Nav/Settings", "\uE713", "settings"));

            navigationView.SelectionChanged += (sender, args) =>
            {
                if (args.SelectedItem is NavigationViewItem { Tag: string tag })
                {
                    var pageType = tag switch
                    {
                        "settings" => typeof(SettingsPage),
                        _ => typeof(HomePage)
                    };
                    _ = contentFrame.Navigate(pageType);
                }
            };

            var rootGrid = new Grid
            {
                Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0, 0, 0, 0))
            };
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(48) });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            Grid.SetRow(titleBar, 0);
            Grid.SetRow(navigationView, 1);
            rootGrid.Children.Add(titleBar);
            rootGrid.Children.Add(navigationView);

            targetWindow.ExtendsContentIntoTitleBar = true;
            targetWindow.Content = rootGrid;
            targetWindow.SetTitleBar(titleBar);

            // Default to home page
            _ = contentFrame.Navigate(typeof(HomePage));
            navigationView.SelectedItem = navigationView.MenuItems[0];
        }

        private static NavigationViewItem CreateNavItem(string resourceKey, string glyph, string tag)
        {
            return new NavigationViewItem
            {
                Content = SR(resourceKey),
                Icon = new FontIcon { Glyph = glyph },
                Tag = tag
            };
        }

        private static TitleBar CreateTitleBar(NavigationView navigationView)
        {
            var titleBar = new TitleBar
            {
                Title = SR("App/Title"),
                Height = 48,
                Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0, 0, 0, 0)),
                IsBackButtonVisible = false,
                IsPaneToggleButtonVisible = true,
                IconSource = new ImageIconSource
                {
                    ImageSource = new BitmapImage(new Uri("ms-appx:///Assets/AppIcon.ico"))
                }
            };

            titleBar.Loaded += (_, _) =>
            {
                var paneButton = FindVisualChild<Button>(titleBar, "PaneToggleButton");
                if (paneButton is not null)
                {
                    paneButton.Click += (_, _) => navigationView.IsPaneOpen = !navigationView.IsPaneOpen;
                }
            };

            return titleBar;
        }

        private static void SetWindowIcon(Window targetWindow)
        {
            try
            {
                var hwnd = WindowNative.GetWindowHandle(targetWindow);
                var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
                var appWindow = AppWindow.GetFromWindowId(windowId);

                var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico");
                if (!File.Exists(iconPath))
                {
                    return;
                }

                appWindow.SetIcon(iconPath);
            }
            catch
            {
                // Ignore icon setup failures to avoid affecting startup flow.
            }
        }

        private void EnsureThemeListener()
        {
            if (uiSettings is not null)
            {
                return;
            }

            uiSettings = new UISettings();
            uiSettings.ColorValuesChanged += OnSystemColorValuesChanged;
        }

        private void OnSystemColorValuesChanged(UISettings sender, object args)
        {
            var targetWindow = MainWindow;
            if (targetWindow is null)
            {
                return;
            }

            _ = targetWindow.DispatcherQueue.TryEnqueue(() => ApplySystemTitleBarTheme(targetWindow));
        }

        private static void ApplySystemTitleBarTheme(Window targetWindow)
        {
            try
            {
                var hwnd = WindowNative.GetWindowHandle(targetWindow);
                var isDark = IsSystemDarkModeEnabled();
                var useDark = isDark ? 1 : 0;
                _ = DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useDark, sizeof(int));
                ApplyCaptionButtonTheme(targetWindow, isDark);
            }
            catch
            {
                // Ignore title bar theme failures to avoid affecting startup flow.
            }
        }

        private static void ApplyCaptionButtonTheme(Window targetWindow, bool isDark)
        {
            var hwnd = WindowNative.GetWindowHandle(targetWindow);
            var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            if (!AppWindowTitleBar.IsCustomizationSupported())
            {
                return;
            }

            var foreground = isDark
                ? Windows.UI.Color.FromArgb(255, 255, 255, 255)
                : Windows.UI.Color.FromArgb(255, 0, 0, 0);
            var inactiveForeground = isDark
                ? Windows.UI.Color.FromArgb(160, 255, 255, 255)
                : Windows.UI.Color.FromArgb(160, 0, 0, 0);
            var hoverBackground = isDark
                ? Windows.UI.Color.FromArgb(32, 255, 255, 255)
                : Windows.UI.Color.FromArgb(24, 0, 0, 0);
            var pressedBackground = isDark
                ? Windows.UI.Color.FromArgb(48, 255, 255, 255)
                : Windows.UI.Color.FromArgb(36, 0, 0, 0);

            appWindow.TitleBar.ButtonBackgroundColor = Windows.UI.Color.FromArgb(0, 0, 0, 0);
            appWindow.TitleBar.ButtonInactiveBackgroundColor = Windows.UI.Color.FromArgb(0, 0, 0, 0);
            appWindow.TitleBar.ButtonHoverBackgroundColor = hoverBackground;
            appWindow.TitleBar.ButtonPressedBackgroundColor = pressedBackground;
            appWindow.TitleBar.ButtonForegroundColor = foreground;
            appWindow.TitleBar.ButtonInactiveForegroundColor = inactiveForeground;
        }

        private static bool IsSystemDarkModeEnabled()
        {
            try
            {
                using var personalizeKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = personalizeKey?.GetValue("AppsUseLightTheme");
                return value is int lightThemeEnabled && lightThemeEnabled == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Invoked when Navigation to a certain page fails
        /// </summary>
        /// <param name="sender">The Frame which failed navigation</param>
        /// <param name="e">Details about the navigation failure</param>
        void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
        {
            throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
        }

        private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            WriteCrashLog("XamlUnhandledException", e.Exception);
        }

        private void CurrentDomain_UnhandledException(object? sender, System.UnhandledExceptionEventArgs e)
        {
            WriteCrashLog("AppDomainUnhandledException", e.ExceptionObject as Exception);
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            WriteCrashLog("TaskSchedulerUnobservedTaskException", e.Exception);
        }

        private static void WriteCrashLog(string source, Exception? ex)
        {
            try
            {
                var logPath = GetCrashLogPath();
                var directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var text = new StringBuilder()
                    .AppendLine("====")
                    .AppendLine($"Time: {DateTimeOffset.Now:O}")
                    .AppendLine($"Source: {source}")
                    .AppendLine($"ExceptionType: {ex?.GetType().FullName}")
                    .AppendLine($"Message: {ex?.Message}")
                    .AppendLine("StackTrace:")
                    .AppendLine(ex?.ToString())
                    .AppendLine()
                    .ToString();

                File.AppendAllText(logPath, text, new UTF8Encoding(false));
            }
            catch
            {
                // Do not throw from crash logger.
            }
        }

        private static string GetCrashLogPath()
        {
            try
            {
                return Path.Combine(ApplicationData.Current.LocalFolder.Path, "crash.log");
            }
            catch
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "wows-ime",
                    "crash.log");
            }
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr hwnd,
            int dwAttribute,
            ref int pvAttribute,
            int cbAttribute);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

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
}
