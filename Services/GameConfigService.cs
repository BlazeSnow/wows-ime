using System.Text;
using System.Xml.Linq;
using wows_ime.Views;

namespace wows_ime.Services;

internal static class GameConfigService
{
    private const string WowsExeName = "WorldOfWarships.exe";
    private const string KorabliExeName = "Korabli.exe";
    private const string TargetConfigRelativePath = "res_mods\\ime_config.xml";
    private const string TagSimplified = "GFxIME_Ch_Simp";
    private const string TagTraditional = "GFxIME_Ch_Trad_Array";
    private const string TagJapanese = "GFxIME_Jp";

    internal static bool HasGameExecutable(string gameRoot)
    {
        var wows = Path.Combine(gameRoot, WowsExeName);
        var korabli = Path.Combine(gameRoot, KorabliExeName);
        return File.Exists(wows) || File.Exists(korabli);
    }

    internal static List<string> ResolveTargetConfigFiles(string gameRoot)
    {
        var binPath = Path.Combine(gameRoot, "bin");
        var results = new List<string>();

        if (Directory.Exists(binPath))
        {
            var numericVersionDirs = Directory
                .GetDirectories(binPath)
                .Where(path => int.TryParse(Path.GetFileName(path), out _))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var versionDir in numericVersionDirs)
            {
                results.Add(Path.Combine(versionDir, TargetConfigRelativePath));
            }
        }

        return results;
    }

    internal static async Task WriteConfigFilesAsync(IEnumerable<string> targetFiles, IEnumerable<InputMethodItem> selectedIme)
    {
        var document = BuildConfigDocument(selectedIme);
        foreach (var targetFile in targetFiles)
        {
            var targetDirectory = Path.GetDirectoryName(targetFile);
            if (!string.IsNullOrEmpty(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            await using var stream = new FileStream(targetFile, FileMode.Create, FileAccess.Write, FileShare.None);
            await using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            await writer.WriteAsync(document.ToString());
        }
    }

    private static XDocument BuildConfigDocument(IEnumerable<InputMethodItem> selectedIme)
    {
        var simplified = new XElement("ChineseSimplified");
        var traditional = new XElement("ChineseTraditional");
        var japanese = new XElement("Japanese");

        foreach (var ime in selectedIme)
        {
            var target = ime.Category switch
            {
                ImeCategory.ChineseSimplified => simplified,
                ImeCategory.ChineseTraditional => traditional,
                _ => japanese
            };

            var tag = ime.Category switch
            {
                ImeCategory.ChineseSimplified => TagSimplified,
                ImeCategory.ChineseTraditional => TagTraditional,
                _ => TagJapanese
            };

            target.Add(new XElement("imeName", ime.DisplayName));
            target.Add(new XElement("displayName", ime.DisplayName));
            target.Add(new XElement("Tag", tag));
        }

        return new XDocument(
            new XElement("data",
                new XElement("language", simplified, japanese, traditional)
            )
        );
    }
}
