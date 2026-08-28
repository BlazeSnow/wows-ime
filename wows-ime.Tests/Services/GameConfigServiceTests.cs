using System.Text;
using System.Xml.Linq;
using wows_ime.Core.Models;
using wows_ime.Core.Services;

namespace wows_ime.Tests.Services;

public sealed class GameConfigServiceTests : IDisposable
{
    private readonly string root;

    public GameConfigServiceTests()
    {
        root = Path.Combine(Path.GetTempPath(), "wows-ime-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(root, recursive: true);
        }
        catch
        {
            // Temp directory cleanup must not fail the test run.
        }
    }

    [Fact]
    public void HasGameExecutable_NullOrWhitespace_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => GameConfigService.HasGameExecutable(null!));
        Assert.Throws<ArgumentException>(() => GameConfigService.HasGameExecutable(" "));
    }

    [Fact]
    public void HasGameExecutable_EmptyDirectory_ReturnsFalse()
    {
        Assert.False(GameConfigService.HasGameExecutable(root));
    }

    [Fact]
    public void HasGameExecutable_WorldOfWarshipsExe_ReturnsTrue()
    {
        File.WriteAllText(Path.Combine(root, "WorldOfWarships.exe"), string.Empty);
        Assert.True(GameConfigService.HasGameExecutable(root));
    }

    [Fact]
    public void HasGameExecutable_KorabliExe_ReturnsTrue()
    {
        File.WriteAllText(Path.Combine(root, "Korabli.exe"), string.Empty);
        Assert.True(GameConfigService.HasGameExecutable(root));
    }

    [Fact]
    public void ResolveTargetConfigFiles_MissingBinDirectory_ReturnsEmpty()
    {
        Assert.Empty(GameConfigService.ResolveTargetConfigFiles(root));
    }

    [Fact]
    public void ResolveTargetConfigFiles_NoNumericVersionDirectory_ReturnsEmpty()
    {
        var bin = Directory.CreateDirectory(Path.Combine(root, "bin"));
        Directory.CreateDirectory(Path.Combine(bin.FullName, "launcher"));
        Directory.CreateDirectory(Path.Combine(bin.FullName, "res_mods"));

        Assert.Empty(GameConfigService.ResolveTargetConfigFiles(root));
    }

    [Fact]
    public void ResolveTargetConfigFiles_NumericVersionDirectories_ReturnsConfigPaths()
    {
        var bin = Directory.CreateDirectory(Path.Combine(root, "bin"));
        Directory.CreateDirectory(Path.Combine(bin.FullName, "8842736"));
        Directory.CreateDirectory(Path.Combine(bin.FullName, "8842737"));
        Directory.CreateDirectory(Path.Combine(bin.FullName, "launcher"));

        var results = GameConfigService.ResolveTargetConfigFiles(root);

        var expected = new[]
        {
            Path.Combine(root, "bin", "8842736", "res_mods", "ime_config.xml"),
            Path.Combine(root, "bin", "8842737", "res_mods", "ime_config.xml")
        };
        Assert.Equal(
            expected.Select(Path.GetFullPath).OrderBy(path => path, StringComparer.OrdinalIgnoreCase),
            results.Select(Path.GetFullPath).OrderBy(path => path, StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void BuildConfigDocument_NullInputMethods_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => GameConfigService.BuildConfigDocument(null!));
    }

    [Fact]
    public void BuildConfigDocument_GroupsByCategory_WithTagsAndNames()
    {
        var inputMethods = new[]
        {
            new InputMethodDefinition("微软拼音", ImeCategory.ChineseSimplified),
            new InputMethodDefinition("微軟速成", ImeCategory.ChineseTraditional),
            new InputMethodDefinition("Microsoft IME", ImeCategory.Japanese)
        };

        var document = GameConfigService.BuildConfigDocument(inputMethods);
        var language = Assert.Single(document.Root!.Elements("language"));

        var simplified = language.Element("ChineseSimplified");
        Assert.NotNull(simplified);
        Assert.Equal("微软拼音", simplified!.Element("imeName")!.Value);
        Assert.Equal("微软拼音", simplified.Element("displayName")!.Value);
        Assert.Equal("GFxIME_Ch_Simp", simplified.Element("Tag")!.Value);

        var traditional = language.Element("ChineseTraditional");
        Assert.NotNull(traditional);
        Assert.Equal("微軟速成", traditional!.Element("imeName")!.Value);
        Assert.Equal("GFxIME_Ch_Trad_Array", traditional.Element("Tag")!.Value);

        var japanese = language.Element("Japanese");
        Assert.NotNull(japanese);
        Assert.Equal("Microsoft IME", japanese!.Element("imeName")!.Value);
        Assert.Equal("GFxIME_Jp", japanese.Element("Tag")!.Value);
    }

    [Fact]
    public void BuildConfigDocument_OrdersCategoryElements_SimplifiedJapaneseTraditional()
    {
        var inputMethods = new[]
        {
            new InputMethodDefinition("微軟速成", ImeCategory.ChineseTraditional),
            new InputMethodDefinition("Microsoft IME", ImeCategory.Japanese),
            new InputMethodDefinition("微软拼音", ImeCategory.ChineseSimplified)
        };

        var document = GameConfigService.BuildConfigDocument(inputMethods);
        var language = Assert.Single(document.Root!.Elements("language"));

        Assert.Equal(
            new[] { "ChineseSimplified", "Japanese", "ChineseTraditional" },
            language.Elements().Select(element => element.Name.LocalName));
    }

    [Fact]
    public async Task WriteConfigFilesAsync_CreatesDirectories_WritesUtf8WithoutBom()
    {
        var inputMethods = new[]
        {
            new InputMethodDefinition("微软拼音", ImeCategory.ChineseSimplified)
        };
        var targetFile = Path.Combine(root, "bin", "9999999", "res_mods", "ime_config.xml");

        await GameConfigService.WriteConfigFilesAsync([targetFile], inputMethods);

        Assert.True(File.Exists(targetFile));
        var bytes = File.ReadAllBytes(targetFile);
        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
        var expected = GameConfigService.BuildConfigDocument(inputMethods).ToString();
        Assert.Equal(expected, File.ReadAllText(targetFile, Encoding.UTF8));
    }

    [Fact]
    public async Task WriteConfigFilesAsync_CanceledToken_ThrowsOperationCanceled()
    {
        var inputMethods = new[]
        {
            new InputMethodDefinition("微软拼音", ImeCategory.ChineseSimplified)
        };
        var targetFile = Path.Combine(root, "bin", "9999999", "res_mods", "ime_config.xml");
        var token = new CancellationToken(canceled: true);

        await Assert.ThrowsAsync<OperationCanceledException>(
            () => GameConfigService.WriteConfigFilesAsync([targetFile], inputMethods, token));
        Assert.False(File.Exists(targetFile));
    }
}
