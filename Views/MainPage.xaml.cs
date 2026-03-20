using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace wows_ime.Views;

public sealed partial class MainPage : Page
{
    private const string SteamDefaultPath = @"C:\Program Files (x86)\Steam\steamapps\common\World of Warships";
    private const string WowsExeName = "WorldOfWarships.exe";
    private const string KorabliExeName = "Korabli.exe";
    private const string TargetConfigRelativePath = "res_mods\\ime_config.xml";
    private const string TagSimplified = "GFxIME_Ch_Simp";
    private const string TagTraditional = "GFxIME_Ch_Trad_Array";
    private const string TagJapanese = "GFxIME_Jp";

    public ObservableCollection<InputMethodItem> InputMethods { get; } = new();

    public MainPage()
    {
        InitializeComponent();
        GameRootPathBox.Text = SteamDefaultPath;
        LoadInputMethods();
    }

    private async void PickFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");

        if (App.MainWindow is null)
        {
            ShowStatus("窗口句柄不可用，无法打开目录选择器。", InfoBarSeverity.Error);
            return;
        }

        var hwnd = WindowNative.GetWindowHandle(App.MainWindow);
        InitializeWithWindow.Initialize(picker, hwnd);

        var folder = await picker.PickSingleFolderAsync();
        if (folder is not null)
        {
            GameRootPathBox.Text = folder.Path;
        }
    }

    private void RefreshImeButton_Click(object sender, RoutedEventArgs e)
    {
        LoadInputMethods();
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var gameRoot = GameRootPathBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(gameRoot) || !Directory.Exists(gameRoot))
        {
            ShowStatus("目录不存在，无法打开。", InfoBarSeverity.Warning);
            return;
        }

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = gameRoot,
                UseShellExecute = true
            };

            _ = Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ShowStatus($"打开目录失败：{ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void AddCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        var nameBox = new TextBox
        {
            PlaceholderText = "请输入输入法名称"
        };

        var categoryCombo = new ComboBox
        {
            SelectedIndex = 0
        };
        categoryCombo.Items.Add(new ComboBoxItem { Content = "中文简体" });
        categoryCombo.Items.Add(new ComboBoxItem { Content = "中文繁体" });
        categoryCombo.Items.Add(new ComboBoxItem { Content = "日文" });

        var panel = new StackPanel { Spacing = 8 };
        panel.Children.Add(new TextBlock { Text = "输入法名称" });
        panel.Children.Add(nameBox);
        panel.Children.Add(new TextBlock { Text = "输入法类型" });
        panel.Children.Add(categoryCombo);

        var dialog = new ContentDialog
        {
            Title = "添加自定义输入法",
            Content = panel,
            PrimaryButtonText = "添加",
            CloseButtonText = "取消",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = XamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        var name = nameBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowStatus("输入法名称不能为空。", InfoBarSeverity.Warning);
            return;
        }

        if (InputMethods.Any(item => string.Equals(item.DisplayName, name, StringComparison.OrdinalIgnoreCase)))
        {
            ShowStatus("该输入法名称已存在。", InfoBarSeverity.Warning);
            return;
        }

        var newItem = new InputMethodItem(name)
        {
            IsSelected = true,
            CategoryIndex = categoryCombo.SelectedIndex < 0 ? 0 : categoryCombo.SelectedIndex
        };

        InputMethods.Add(newItem);
        ShowStatus("已添加自定义输入法。", InfoBarSeverity.Success);
    }

    private async void WriteConfigButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedIme = InputMethods.Where(item => item.IsSelected).ToList();
        if (selectedIme.Count == 0)
        {
            ShowStatus("请至少勾选一个输入法。", InfoBarSeverity.Warning);
            return;
        }

        var gameRoot = GameRootPathBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(gameRoot) || !Directory.Exists(gameRoot))
        {
            ShowStatus("游戏根目录不存在，请先选择有效目录。", InfoBarSeverity.Error);
            return;
        }

        if (!HasGameExecutable(gameRoot))
        {
            ShowStatus("所选目录未找到 WorldOfWarships.exe 或 Korabli.exe。", InfoBarSeverity.Error);
            return;
        }

        var targetFiles = ResolveTargetConfigFiles(gameRoot);
        if (targetFiles.Count == 0)
        {
            await ShowErrorDialogAsync("未在游戏目录的 bin 下找到数字版本目录，无法确定 ime_config.xml 写入位置。");
            return;
        }

        var existing = targetFiles.Where(File.Exists).ToList();
        if (existing.Count > 0)
        {
            var shouldOverwrite = await ConfirmOverwriteAsync(existing.Count);
            if (!shouldOverwrite)
            {
                ShowStatus("已取消写入。", InfoBarSeverity.Informational);
                return;
            }
        }

        var document = BuildConfigDocument(selectedIme);
        try
        {
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

            ShowStatus($"写入成功，共更新 {targetFiles.Count} 个配置文件。", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            ShowStatus($"写入失败：{ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void LoadInputMethods()
    {
        InputMethods.Clear();

        foreach (var imeName in ReadImeNamesFromRegistry())
        {
            InputMethods.Add(new InputMethodItem(imeName));
        }

        ShowStatus($"扫描完成，共发现 {InputMethods.Count} 个输入法。", InfoBarSeverity.Success);
    }

    private static IEnumerable<string> ReadImeNamesFromRegistry()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var preloadCodes = new List<string>();

        using (var preloadKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Keyboard Layout\Preload"))
        {
            if (preloadKey is not null)
            {
                foreach (var valueName in preloadKey.GetValueNames())
                {
                    var value = preloadKey.GetValue(valueName)?.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        preloadCodes.Add(value);
                    }
                }
            }
        }

        var substitutes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using (var substitutesKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Keyboard Layout\Substitutes"))
        {
            if (substitutesKey is not null)
            {
                foreach (var valueName in substitutesKey.GetValueNames())
                {
                    var value = substitutesKey.GetValue(valueName)?.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        substitutes[valueName] = value;
                    }
                }
            }
        }

        foreach (var code in preloadCodes)
        {
            var normalizedCode = substitutes.TryGetValue(code, out var substitute) ? substitute : code;
            var displayName = ResolveLayoutDisplayName(normalizedCode);
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                _ = names.Add(displayName);
            }
        }

        foreach (var tipName in ReadTipProfileNamesFromRegistry())
        {
            _ = names.Add(tipName);
        }

        return names.OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase);
    }

    private static IEnumerable<string> ReadTipProfileNamesFromRegistry()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        CollectTipProfileNamesFromRoot(Microsoft.Win32.Registry.CurrentUser, names);
        CollectTipProfileNamesFromRoot(Microsoft.Win32.Registry.LocalMachine, names);

        return names;
    }

    private static void CollectTipProfileNamesFromRoot(Microsoft.Win32.RegistryKey root, ISet<string> names)
    {
        using var tipRoot = root.OpenSubKey(@"Software\Microsoft\CTF\TIP");
        if (tipRoot is null)
        {
            return;
        }

        foreach (var clsid in tipRoot.GetSubKeyNames())
        {
            using var languageProfileKey = tipRoot.OpenSubKey($@"{clsid}\LanguageProfile");
            if (languageProfileKey is null)
            {
                continue;
            }

            foreach (var langId in languageProfileKey.GetSubKeyNames())
            {
                if (!IsTargetLanguageProfile(langId))
                {
                    continue;
                }

                using var langKey = languageProfileKey.OpenSubKey(langId);
                if (langKey is null)
                {
                    continue;
                }

                foreach (var profileGuid in langKey.GetSubKeyNames())
                {
                    var name = ResolveTipProfileDisplayName(clsid, langId, profileGuid);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        _ = names.Add(NormalizeImeDisplayName(name));
                    }
                }
            }
        }
    }

    private static bool IsTargetLanguageProfile(string langId) =>
        string.Equals(langId, "0x00000804", StringComparison.OrdinalIgnoreCase) || // zh-CN
        string.Equals(langId, "0x00000404", StringComparison.OrdinalIgnoreCase) || // zh-TW
        string.Equals(langId, "0x00000C04", StringComparison.OrdinalIgnoreCase) || // zh-HK
        string.Equals(langId, "0x00001004", StringComparison.OrdinalIgnoreCase) || // zh-SG
        string.Equals(langId, "0x00001404", StringComparison.OrdinalIgnoreCase) || // zh-MO
        string.Equals(langId, "0x00000411", StringComparison.OrdinalIgnoreCase);   // ja-JP

    private static string NormalizeImeDisplayName(string name)
    {
        if (string.Equals(name, "WeType", StringComparison.OrdinalIgnoreCase))
        {
            return "微信输入法";
        }

        return name;
    }

    private static string? ResolveTipProfileDisplayName(string clsid, string langId, string profileGuid)
    {
        var relativePath = $@"Software\Microsoft\CTF\TIP\{clsid}\LanguageProfile\{langId}\{profileGuid}";
        var candidate = ReadTipNameFromPath(Microsoft.Win32.Registry.CurrentUser, relativePath);
        if (!string.IsNullOrWhiteSpace(candidate))
        {
            return candidate;
        }

        return ReadTipNameFromPath(Microsoft.Win32.Registry.LocalMachine, relativePath);
    }

    private static string? ReadTipNameFromPath(Microsoft.Win32.RegistryKey root, string relativePath)
    {
        using var key = root.OpenSubKey(relativePath);
        if (key is null)
        {
            return null;
        }

        var description = key.GetValue("Description")?.ToString();
        var displayDescription = key.GetValue("Display Description")?.ToString();

        var resolvedDisplay = TryResolveIndirectString(displayDescription);
        if (!string.IsNullOrWhiteSpace(resolvedDisplay))
        {
            return resolvedDisplay;
        }

        if (!string.IsNullOrWhiteSpace(description))
        {
            return description;
        }

        return displayDescription;
    }

    private static string? TryResolveIndirectString(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || !value.StartsWith("@", StringComparison.Ordinal))
        {
            return value;
        }

        var buffer = new StringBuilder(512);
        var result = SHLoadIndirectString(value, buffer, buffer.Capacity, IntPtr.Zero);
        if (result != 0)
        {
            return value;
        }

        var resolved = buffer.ToString().Trim();
        return string.IsNullOrWhiteSpace(resolved) ? value : resolved;
    }

    private static string? ResolveLayoutDisplayName(string code)
    {
        using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Control\Keyboard Layouts\{code}");
        if (key is null)
        {
            return null;
        }

        var imeName = key.GetValue("Ime File")?.ToString();
        var layoutText = key.GetValue("Layout Text")?.ToString();
        if (!string.IsNullOrWhiteSpace(layoutText))
        {
            return layoutText;
        }

        return !string.IsNullOrWhiteSpace(imeName) ? imeName : null;
    }

    private static bool HasGameExecutable(string gameRoot)
    {
        var wows = Path.Combine(gameRoot, WowsExeName);
        var korabli = Path.Combine(gameRoot, KorabliExeName);
        return File.Exists(wows) || File.Exists(korabli);
    }

    private static List<string> ResolveTargetConfigFiles(string gameRoot)
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

    private XDocument BuildConfigDocument(IEnumerable<InputMethodItem> selectedIme)
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

    private async Task<bool> ConfirmOverwriteAsync(int existingCount)
    {
        var dialog = new ContentDialog
        {
            Title = "发现已有配置文件",
            Content = $"检测到 {existingCount} 个已有 ime_config.xml，是否覆盖？",
            PrimaryButtonText = "覆盖",
            CloseButtonText = "取消",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = XamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    private async Task ShowErrorDialogAsync(string message)
    {
        var dialog = new ContentDialog
        {
            Title = "无法写入配置",
            Content = message,
            CloseButtonText = "确定",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = XamlRoot
        };

        _ = await dialog.ShowAsync();
    }

    private void ShowStatus(string message, InfoBarSeverity severity)
    {
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;
    }

    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int SHLoadIndirectString(
        string pszSource,
        StringBuilder pszOutBuf,
        int cchOutBuf,
        IntPtr ppvReserved);
}

public enum ImeCategory
{
    ChineseSimplified = 0,
    ChineseTraditional = 1,
    Japanese = 2
}

public sealed class InputMethodItem : Microsoft.UI.Xaml.DependencyObject
{
    public InputMethodItem(string displayName)
    {
        DisplayName = displayName;
        Category = ImeCategory.ChineseSimplified;
    }

    public string DisplayName { get; }

    public bool IsSelected
    {
        get => (bool)GetValue(IsSelectedProperty);
        set => SetValue(IsSelectedProperty, value);
    }

    public static readonly Microsoft.UI.Xaml.DependencyProperty IsSelectedProperty =
        Microsoft.UI.Xaml.DependencyProperty.Register(
            nameof(IsSelected),
            typeof(bool),
            typeof(InputMethodItem),
            new PropertyMetadata(false));

    public int CategoryIndex
    {
        get => (int)GetValue(CategoryIndexProperty);
        set
        {
            SetValue(CategoryIndexProperty, value);
            Category = value switch
            {
                1 => ImeCategory.ChineseTraditional,
                2 => ImeCategory.Japanese,
                _ => ImeCategory.ChineseSimplified
            };
        }
    }

    public static readonly Microsoft.UI.Xaml.DependencyProperty CategoryIndexProperty =
        Microsoft.UI.Xaml.DependencyProperty.Register(
            nameof(CategoryIndex),
            typeof(int),
            typeof(InputMethodItem),
            new PropertyMetadata(0));

    public ImeCategory Category { get; private set; }
}
