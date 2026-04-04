using System.Text;
using Microsoft.Win32;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml.Navigation;
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
            SetWindowIcon(window);
            ApplySystemTitleBarTheme(window);
            EnsureThemeListener();

            if (window.Content is not Frame rootFrame)
            {
                rootFrame = new Frame();
                rootFrame.NavigationFailed += OnNavigationFailed;
                window.Content = rootFrame;
            }

            _ = rootFrame.Navigate(typeof(MainPage), e.Arguments);
            window.Activate();
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
                var useDark = IsSystemDarkModeEnabled() ? 1 : 0;
                _ = DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useDark, sizeof(int));
            }
            catch
            {
                // Ignore title bar theme failures to avoid affecting startup flow.
            }
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

        private static string SR(string key)
        {
            var value = ResourceLoader.GetString(key);
            return string.IsNullOrEmpty(value) ? key : value;
        }
    }
}

