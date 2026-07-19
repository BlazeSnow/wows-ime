using System.Globalization;
using Windows.ApplicationModel.Resources;
using ModernResourceManager = Microsoft.Windows.ApplicationModel.Resources.ResourceManager;

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

        try
        {
            var resourceManager = new ModernResourceManager();
            var resourceMap = resourceManager.MainResourceMap.GetSubtree("Resources");
            var context = resourceManager.CreateResourceContext();
            context.QualifierValues["Language"] = language;
            var value = resourceMap.GetValue(key, context)?.ValueAsString;
            return string.IsNullOrEmpty(value) ? Get(key) : value;
        }
        catch
        {
            return Get(key);
        }
    }

    internal static string Format(string key, params object[] args)
    {
        return string.Format(CultureInfo.CurrentUICulture, Get(key), args);
    }
}
