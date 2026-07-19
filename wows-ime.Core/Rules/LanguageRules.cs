using wows_ime.Core.Abstractions;

namespace wows_ime.Core.Rules;

public static class LanguageRules
{
    public const string Automatic = "auto";
    public const string ChineseSimplified = "zh-Hans";
    public const string ChineseTraditional = "zh-Hant";
    public const string Japanese = "ja";

    public static string NormalizeMode(string? languageMode) => languageMode is
        ChineseSimplified or ChineseTraditional or Japanese
        ? languageMode
        : Automatic;

    public static string ResolveDisplayLanguage(string? languageMode, ISystemLanguagePreferences systemLanguagePreferences)
    {
        ArgumentNullException.ThrowIfNull(systemLanguagePreferences);

        var normalizedMode = NormalizeMode(languageMode);
        if (normalizedMode != Automatic)
        {
            return normalizedMode;
        }

        foreach (var preferredLanguage in systemLanguagePreferences.Languages)
        {
            if (IsSimplifiedChinese(preferredLanguage))
            {
                return ChineseSimplified;
            }

            if (IsTraditionalChinese(preferredLanguage))
            {
                return ChineseTraditional;
            }

            if (preferredLanguage.Equals(Japanese, StringComparison.OrdinalIgnoreCase) ||
                preferredLanguage.StartsWith("ja-", StringComparison.OrdinalIgnoreCase))
            {
                return Japanese;
            }
        }

        return ChineseSimplified;
    }

    public static bool IsSupportedMode(string? languageMode) => languageMode is
        Automatic or ChineseSimplified or ChineseTraditional or Japanese;

    private static bool IsSimplifiedChinese(string language) =>
        language.Equals(ChineseSimplified, StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-Hans-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-CN", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-CN-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-SG", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-SG-", StringComparison.OrdinalIgnoreCase);

    private static bool IsTraditionalChinese(string language) =>
        language.Equals(ChineseTraditional, StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-Hant-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-TW", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-TW-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-HK", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-HK-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-MO", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-MO-", StringComparison.OrdinalIgnoreCase);
}
