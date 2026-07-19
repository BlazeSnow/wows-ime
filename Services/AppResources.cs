using System.Globalization;
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;

namespace wows_ime.Services;

internal static class AppResources
{
    private static readonly ResourceLoader ResourceLoader = new();

    internal static string Get(string key)
    {
        var value = ResourceLoader.GetString(key);
        return string.IsNullOrEmpty(value) ? key : value;
    }

    internal static string GetForLanguage(string key, string language)
    {
        if (language == "auto")
        {
            return Get(key);
        }

        var context = ResourceContext.GetForCurrentView().Clone();
        context.QualifierValues["Language"] = language;
        var resourceMap = ResourceManager.Current.MainResourceMap.GetSubtree("Resources");
        var value = resourceMap.GetValue(key, context)?.ValueAsString;
        return string.IsNullOrEmpty(value) ? key : value;
    }

    internal static string Format(string key, params object[] args)
    {
        return string.Format(CultureInfo.CurrentUICulture, Get(key), args);
    }
}
