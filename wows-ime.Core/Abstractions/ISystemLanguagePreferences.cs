namespace wows_ime.Core.Abstractions;

public interface ISystemLanguagePreferences
{
    IReadOnlyList<string> Languages { get; }
}
