using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml.Media;
using Microsoft.Win32;
using System.Text;
using Windows.ApplicationModel.Resources;
using Windows.Storage;
using Windows.UI.ViewManagement;
using WinRT.Interop;
using wows_ime.Core.Infrastructure;
using wows_ime.Pages.Views;

namespace wows_ime
{
    public partial class App : Application
    {
        private Window window = Window.Current;
        private UISettings? uiSettings;
        private readonly SettingsPersistence settings = new();
        private static readonly ResourceLoader ResourceLoader = new();
        public static Window? MainWindow { get; private set; }

        public App()
        {
            settings.ApplyLanguageMode();
            this.InitializeComponent();
            this.UnhandledException += App_UnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        protected override void OnLaunched(LaunchActivatedEventArgs e)
        {
            window ??= new Window();
            MainWindow = window;
            window.Title = SR("App/Title");
            window.SystemBackdrop = new MicaBackdrop();
            SetWindowIcon(window);

            settings.Initialize();
            var shell = new Shell(new PageHost(window, settings));
            window.ExtendsContentIntoTitleBar = true;
            window.Content = shell;
            window.SetTitleBar(shell.AppTitleBar);

            ApplySystemTitleBarTheme(window);
            EnsureThemeListener();
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
                ? global::Windows.UI.Color.FromArgb(255, 255, 255, 255)
                : global::Windows.UI.Color.FromArgb(255, 0, 0, 0);
            var inactiveForeground = isDark
                ? global::Windows.UI.Color.FromArgb(160, 255, 255, 255)
                : global::Windows.UI.Color.FromArgb(160, 0, 0, 0);
            var hoverBackground = isDark
                ? global::Windows.UI.Color.FromArgb(32, 255, 255, 255)
                : global::Windows.UI.Color.FromArgb(24, 0, 0, 0);
            var pressedBackground = isDark
                ? global::Windows.UI.Color.FromArgb(48, 255, 255, 255)
                : global::Windows.UI.Color.FromArgb(36, 0, 0, 0);

            appWindow.TitleBar.ButtonBackgroundColor = global::Windows.UI.Color.FromArgb(0, 0, 0, 0);
            appWindow.TitleBar.ButtonInactiveBackgroundColor = global::Windows.UI.Color.FromArgb(0, 0, 0, 0);
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
