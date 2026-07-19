using wows_ime.Core.Models;

namespace wows_ime.Core.Abstractions;

public interface ISettingsRepository
{
    string? LoadSelectedGamePath();
    void SaveSelectedGamePath(string selectedGamePath);
    string? LoadLanguageMode();
    void SaveLanguageMode(string languageMode);
    IReadOnlyList<PersistedGamePath> LoadCustomGamePaths();
    void SaveCustomGamePaths(IEnumerable<PersistedGamePath> customGamePaths);
    IReadOnlyList<PersistedInputMethod> LoadCustomInputMethods();
    void SaveCustomInputMethods(IEnumerable<PersistedInputMethod> customInputMethods);
}
