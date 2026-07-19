using wows_ime.Core.Abstractions;

namespace wows_ime.Windows;

public sealed class SystemLanguagePreferences : ISystemLanguagePreferences
{
    public IReadOnlyList<string> Languages
    {
        get
        {
            try
            {
                return global::Windows.System.UserProfile.GlobalizationPreferences.Languages.ToList();
            }
            catch
            {
                return [];
            }
        }
    }
}
