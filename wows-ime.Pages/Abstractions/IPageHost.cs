using wows_ime.Core.Abstractions;
using wows_ime.Core.Models;

namespace wows_ime.Pages.Abstractions;

/// <summary>
/// Supplies the platform, persistence, localization, and application-lifecycle behavior required by Pages.
/// </summary>
public interface IPageHost
{
    IInputMethodScanner InputMethodScanner { get; }
    ISettingsRepository Settings { get; }
    IPageConfiguration Configuration { get; }
    IPageLocalization Localization { get; }
    IPageWindow Window { get; }
    IPageApplication Application { get; }
}

public interface IPageConfiguration
{
    bool DirectoryExists(string path);
    bool HasGameExecutable(string gameRoot);
    IReadOnlyList<string> ResolveTargetConfigFiles(string gameRoot);
    IReadOnlyList<string> GetExistingFiles(IEnumerable<string> paths);
    Task<bool> OpenFolderAsync(string path, CancellationToken cancellationToken = default);
    Task WriteConfigFilesAsync(IEnumerable<InputMethodDefinition> selectedInputMethods, IEnumerable<string> targetFiles, CancellationToken cancellationToken = default);
}

public interface IPageLocalization
{
    string GetString(string key);
    string GetString(string key, string language);
    string Format(string key, params object[] args);
    string ResolveLanguage(string languageMode);
}

public interface IPageWindow
{
    IntPtr Handle { get; }
    Task<string?> PickSingleFolderAsync(CancellationToken cancellationToken = default);
    Task<bool> LaunchUriAsync(Uri uri, CancellationToken cancellationToken = default);
    string GetAppVersion();
}

public interface IPageApplication
{
    void SetPrimaryLanguageOverride(string language);
    void Restart();
}
