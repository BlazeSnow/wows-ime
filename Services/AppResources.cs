using System.Globalization;
using Windows.ApplicationModel.Resources;

namespace wows_ime.Services;

internal static class AppResources
{
    private static readonly ResourceLoader ResourceLoader = new();

    internal static string Get(string key)
    {
        var value = ResourceLoader.GetString(key);
        return string.IsNullOrEmpty(value) ? key : value;
    }

    internal static string Format(string key, params object[] args)
    {
        return string.Format(CultureInfo.CurrentUICulture, Get(key), args);
    }
}
