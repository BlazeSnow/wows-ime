using wows_ime.Core.Abstractions;
using wows_ime.Core.Rules;

namespace wows_ime.Tests.Rules;

public sealed class LanguageRulesTests
{
    [Theory]
    [InlineData("zh-Hans")]
    [InlineData("zh-Hant")]
    [InlineData("ja")]
    public void NormalizeMode_SupportedLanguage_ReturnsUnchanged(string languageMode)
    {
        Assert.Equal(languageMode, LanguageRules.NormalizeMode(languageMode));
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("ZH-Hans")]
    [InlineData("en-US")]
    [InlineData("")]
    [InlineData(null)]
    public void NormalizeMode_NotALanguage_ReturnsAutomatic(string? languageMode)
    {
        Assert.Equal(LanguageRules.Automatic, LanguageRules.NormalizeMode(languageMode));
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("zh-Hans")]
    [InlineData("zh-Hant")]
    [InlineData("ja")]
    public void IsSupportedMode_SupportedModes_ReturnsTrue(string languageMode)
    {
        Assert.True(LanguageRules.IsSupportedMode(languageMode));
    }

    [Theory]
    [InlineData("zh-hans")]
    [InlineData("en-US")]
    [InlineData("")]
    [InlineData(null)]
    public void IsSupportedMode_OtherModes_ReturnsFalse(string? languageMode)
    {
        Assert.False(LanguageRules.IsSupportedMode(languageMode));
    }

    [Theory]
    [InlineData("zh-Hans", "zh-Hans")]
    [InlineData("zh-Hant", "zh-Hant")]
    [InlineData("ja", "ja")]
    public void ResolveDisplayLanguage_ExplicitMode_ReturnsModeAsIs(string languageMode, string expected)
    {
        Assert.Equal(expected, LanguageRules.ResolveDisplayLanguage(languageMode, CreatePreferences("en-US")));
    }

    [Theory]
    [InlineData("zh-Hans")]
    [InlineData("zh-CN")]
    [InlineData("zh-SG")]
    [InlineData("zh-Hans-CN")]
    public void ResolveDisplayLanguage_Auto_SimplifiedChinesePreference_ReturnsSimplified(string preferred)
    {
        var result = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences(preferred));
        Assert.Equal(LanguageRules.ChineseSimplified, result);
    }

    [Theory]
    [InlineData("zh-Hant")]
    [InlineData("zh-TW")]
    [InlineData("zh-HK")]
    [InlineData("zh-MO")]
    [InlineData("zh-Hant-TW")]
    public void ResolveDisplayLanguage_Auto_TraditionalChinesePreference_ReturnsTraditional(string preferred)
    {
        var result = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences(preferred));
        Assert.Equal(LanguageRules.ChineseTraditional, result);
    }

    [Theory]
    [InlineData("ja")]
    [InlineData("ja-JP")]
    public void ResolveDisplayLanguage_Auto_JapanesePreference_ReturnsJapanese(string preferred)
    {
        var result = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences(preferred));
        Assert.Equal(LanguageRules.Japanese, result);
    }

    [Fact]
    public void ResolveDisplayLanguage_Auto_UnmatchedPreference_FallsBackToSimplified()
    {
        var result = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences("en-US", "fr-FR"));
        Assert.Equal(LanguageRules.ChineseSimplified, result);
    }

    [Fact]
    public void ResolveDisplayLanguage_Auto_EmptyPreferences_FallsBackToSimplified()
    {
        var result = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences());
        Assert.Equal(LanguageRules.ChineseSimplified, result);
    }

    [Fact]
    public void ResolveDisplayLanguage_Auto_FirstPreferredLanguageWins()
    {
        var japaneseFirst = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences("ja", "zh-Hans"));
        Assert.Equal(LanguageRules.Japanese, japaneseFirst);

        var simplifiedFirst = LanguageRules.ResolveDisplayLanguage("auto", CreatePreferences("zh-Hans", "ja"));
        Assert.Equal(LanguageRules.ChineseSimplified, simplifiedFirst);
    }

    private static ISystemLanguagePreferences CreatePreferences(params string[] languages) =>
        new FakeLanguagePreferences(languages);

    private sealed class FakeLanguagePreferences(params string[] languages) : ISystemLanguagePreferences
    {
        public IReadOnlyList<string> Languages => languages;
    }
}
