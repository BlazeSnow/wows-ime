using System.Globalization;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Resources;
using Windows.Storage.Pickers;
using WinRT.Interop;
using wows_ime.Core.Abstractions;
using wows_ime.Core.Infrastructure;
using wows_ime.Core.Models;
using wows_ime.Core.Rules;
using wows_ime.Core.Services;
using wows_ime.Pages.Abstractions;
using ModernResourceManager = Microsoft.Windows.ApplicationModel.Resources.ResourceManager;

namespace wows_ime;

internal sealed class PageHost : IPageHost, IPageConfiguration, IPageLocalization, IPageWindow, IPageApplication
{
    private readonly Window window;
    private readonly ISettingsRepository settings;
    private readonly ISystemLanguagePreferences systemLanguagePreferences = new SystemLanguagePreferences();
    private readonly ResourceLoader resourceLoader = new();

    public PageHost(Window window, ISettingsRepository settings)
    {
        this.window = window ?? throw new ArgumentNullException(nameof(window));
        this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
    }

    public IInputMethodScanner InputMethodScanner { get; } = new InputMethodScanner();
    public ISettingsRepository Settings => settings;
    public IPageConfiguration Configuration => this;
    public IPageLocalization Localization => this;
    public IPageWindow Window => this;
    public IPageApplication Application => this;

    public bool DirectoryExists(string path) => Directory.Exists(path);

    public bool HasGameExecutable(string gameRoot) => GameConfigService.HasGameExecutable(gameRoot);

    public IReadOnlyList<string> ResolveTargetConfigFiles(string gameRoot) => GameConfigService.ResolveTargetConfigFiles(gameRoot);

    public IReadOnlyList<string> GetExistingFiles(IEnumerable<string> paths) => paths.Where(File.Exists).ToList();

    public Task<bool> OpenFolderAsync(string path, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            _ = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            });
            return Task.FromResult(true);
        }
        catch
        {
            return Task.FromResult(false);
        }
    }

    public Task WriteConfigFilesAsync(
        IEnumerable<InputMethodDefinition> selectedInputMethods,
        IEnumerable<string> targetFiles,
        CancellationToken cancellationToken = default) =>
        GameConfigService.WriteConfigFilesAsync(targetFiles, selectedInputMethods, cancellationToken);

    public string GetString(string key)
    {
        var value = resourceLoader.GetString(key);
        return string.IsNullOrEmpty(value) ? key : value;
    }

    public string GetString(string key, string language)
    {
        if (language == LanguageRules.Automatic)
        {
            return GetString(key);
        }

        try
        {
            var resourceManager = new ModernResourceManager();
            var resourceMap = resourceManager.MainResourceMap.GetSubtree("Resources");
            var context = resourceManager.CreateResourceContext();
            context.QualifierValues["Language"] = language;
            var value = resourceMap.GetValue(key, context)?.ValueAsString;
            return string.IsNullOrEmpty(value) ? GetString(key) : value;
        }
        catch
        {
            return GetString(key);
        }
    }

    public string Format(string key, params object[] args) =>
        string.Format(CultureInfo.CurrentUICulture, GetString(key), args);

    public string ResolveLanguage(string languageMode) =>
        LanguageRules.ResolveDisplayLanguage(languageMode, systemLanguagePreferences);

    public IntPtr Handle => WindowNative.GetWindowHandle(window);

    public async Task<string?> PickSingleFolderAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        InitializeWithWindow.Initialize(picker, Handle);
        var folder = await picker.PickSingleFolderAsync();
        cancellationToken.ThrowIfCancellationRequested();
        return folder?.Path;
    }

    public async Task<bool> LaunchUriAsync(Uri uri, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(uri);
        cancellationToken.ThrowIfCancellationRequested();
        var launched = await global::Windows.System.Launcher.LaunchUriAsync(uri);
        cancellationToken.ThrowIfCancellationRequested();
        return launched;
    }

    public string GetAppVersion()
    {
        try
        {
            var version = Package.Current.Id.Version;
            return $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
        }
        catch
        {
            return string.Empty;
        }
    }

    public void SetPrimaryLanguageOverride(string language) =>
        global::Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = language;

    public void Restart()
    {
        var executablePath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(executablePath))
        {
            _ = System.Diagnostics.Process.Start(executablePath);
        }

        global::Microsoft.UI.Xaml.Application.Current.Exit();
    }
}
