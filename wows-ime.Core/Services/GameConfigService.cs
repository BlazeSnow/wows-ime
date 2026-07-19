using System.Text;
using System.Xml.Linq;
using wows_ime.Core.Models;

namespace wows_ime.Core.Services;

public static class GameConfigService
{
    private const string WowsExeName = "WorldOfWarships.exe";
    private const string KorabliExeName = "Korabli.exe";
    private const string TargetConfigRelativePath = "res_mods\\ime_config.xml";
    private const string TagSimplified = "GFxIME_Ch_Simp";
    private const string TagTraditional = "GFxIME_Ch_Trad_Array";
    private const string TagJapanese = "GFxIME_Jp";

    public static bool HasGameExecutable(string gameRoot)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(gameRoot);

        var wows = Path.Combine(gameRoot, WowsExeName);
        var korabli = Path.Combine(gameRoot, KorabliExeName);
        return File.Exists(wows) || File.Exists(korabli);
    }

    public static List<string> ResolveTargetConfigFiles(string gameRoot)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(gameRoot);

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

    public static async Task WriteConfigFilesAsync(
        IEnumerable<string> targetFiles,
        IEnumerable<InputMethodDefinition> selectedInputMethods,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(targetFiles);
        ArgumentNullException.ThrowIfNull(selectedInputMethods);

        var document = BuildConfigDocument(selectedInputMethods);
        foreach (var targetFile in targetFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var targetDirectory = Path.GetDirectoryName(targetFile);
            if (!string.IsNullOrEmpty(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            await using var stream = new FileStream(
                targetFile,
                FileMode.Create,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 4096,
                useAsync: true);
            await using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            await writer.WriteAsync(document.ToString());
        }
    }

    public static XDocument BuildConfigDocument(IEnumerable<InputMethodDefinition> selectedInputMethods)
    {
        ArgumentNullException.ThrowIfNull(selectedInputMethods);

        var simplified = new XElement("ChineseSimplified");
        var traditional = new XElement("ChineseTraditional");
        var japanese = new XElement("Japanese");

        foreach (var inputMethod in selectedInputMethods)
        {
            ArgumentNullException.ThrowIfNull(inputMethod);

            var target = inputMethod.Category switch
            {
                ImeCategory.ChineseSimplified => simplified,
                ImeCategory.ChineseTraditional => traditional,
                _ => japanese
            };

            var tag = inputMethod.Category switch
            {
                ImeCategory.ChineseSimplified => TagSimplified,
                ImeCategory.ChineseTraditional => TagTraditional,
                _ => TagJapanese
            };

            target.Add(new XElement("imeName", inputMethod.DisplayName));
            target.Add(new XElement("displayName", inputMethod.DisplayName));
            target.Add(new XElement("Tag", tag));
        }

        return new XDocument(
            new XElement("data",
                new XElement("language", simplified, japanese, traditional)
            )
        );
    }
}
