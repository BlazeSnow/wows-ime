using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace wows_ime.Views;

public sealed partial class MainPage : Page
{
    private const string SteamDefaultPath = @"C:\Program Files (x86)\Steam\steamapps\common\World of Warships";
    private const string LestaDefaultPath = @"C:\Games\Korabli";
    private const string Cn360DefaultPath = @"C:\Games\World_of_Warships_CN360";
    private const string WowsExeName = "WorldOfWarships.exe";
    private const string KorabliExeName = "Korabli.exe";
    private const string TargetConfigRelativePath = "res_mods\\ime_config.xml";
    private const string TagSimplified = "GFxIME_Ch_Simp";
    private const string TagTraditional = "GFxIME_Ch_Trad_Array";
    private const string TagJapanese = "GFxIME_Jp";
    private string? lastScanWarning;

    public ObservableCollection<InputMethodItem> InputMethods { get; } = new();

    public MainPage()
    {
        InitializeComponent();
        GameRootPathBox.Text = LoadSavedGameDir() ?? SteamDefaultPath;
        LoadInputMethods();
        LoadSavedCustomIme();
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
            SaveSettings();
        }
    }

    private void RefreshImeButton_Click(object sender, RoutedEventArgs e)
    {
        LoadInputMethods();
        LoadSavedCustomIme();
    }

    private void GameRootPathBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        SaveSettings();
    }

    private void GameRootPathBox_LostFocus(object sender, RoutedEventArgs e)
    {
        SaveSettings();
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

        var newItem = new InputMethodItem(name, isCustom: true)
        {
            IsSelected = true,
            CategoryIndex = categoryCombo.SelectedIndex < 0 ? 0 : categoryCombo.SelectedIndex
        };

        InputMethods.Add(newItem);
        ShowStatus("已添加自定义输入法。", InfoBarSeverity.Success);
        SaveSettings();
    }

    private void DeleteCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: InputMethodItem item })
        {
            return;
        }

        if (!item.IsCustom)
        {
            return;
        }

        _ = InputMethods.Remove(item);
        ShowStatus("已删除自定义输入法。", InfoBarSeverity.Success);
        SaveSettings();
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
            ShowStatus("未在游戏目录的 bin 下找到数字版本目录，无法确定写入位置。", InfoBarSeverity.Error);
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
        else
        {
            var shouldAdd = await ConfirmAddAsync(targetFiles.Count);
            if (!shouldAdd)
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
            SaveSettings();
        }
        catch (Exception ex)
        {
            ShowStatus($"写入失败：{ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void LoadInputMethods()
    {
        InputMethods.Clear();
        string? warning;

        foreach (var ime in ReadImeCandidatesFromRegistry(out warning))
        {
            InputMethods.Add(new InputMethodItem(ime.DisplayName, ime.Category));
        }

        lastScanWarning = warning;

        if (InputMethods.Count == 0 && !string.IsNullOrWhiteSpace(lastScanWarning))
        {
            ShowStatus($"扫描完成，未发现输入法。TSF错误：{lastScanWarning}", InfoBarSeverity.Warning);
            return;
        }

        ShowStatus($"扫描完成，共发现 {InputMethods.Count} 个输入法。", InfoBarSeverity.Success);
    }

    private static IEnumerable<ScannedImeCandidate> ReadImeCandidatesFromRegistry(out string? warning)
    {
        warning = null;
        var candidates = new Dictionary<string, ScannedImeCandidate>(StringComparer.OrdinalIgnoreCase);
        foreach (var candidate in ReadImeCandidatesFromTsf(out warning))
        {
            UpsertCandidate(candidates, candidate);
        }

        return candidates.Values.OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase);
    }

    private static IEnumerable<ScannedImeCandidate> ReadImeCandidatesFromTsf(out string? warning)
    {
        warning = null;
        var candidates = new List<ScannedImeCandidate>();
        var coInitHr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
        var shouldUninitialize = coInitHr == 0 || coInitHr == 1;
        if (coInitHr < 0 && coInitHr != unchecked((int)0x80010106))
        {
            warning = $"CoInitializeEx 失败: 0x{coInitHr:X8}";
            return candidates;
        }

        try
        {
            var profilesPtr = CreateInputProcessorProfilesCom();
            if (profilesPtr == IntPtr.Zero)
            {
                warning = "无法创建 TF_InputProcessorProfiles COM 对象。";
                return candidates;
            }

            try
            {
                var hr = GetLanguageList(profilesPtr, out var langPtr, out var langCount);
                if (hr < 0 || langPtr == IntPtr.Zero || langCount == 0)
                {
                    warning = hr < 0
                        ? $"GetLanguageList 失败: 0x{hr:X8}"
                        : "GetLanguageList 返回空语言列表";
                    return candidates;
                }

                try
                {
                    for (var i = 0; i < langCount; i++)
                    {
                        var langId = (ushort)Marshal.ReadInt16(langPtr, (int)i * sizeof(short));
                        if (!IsTargetLanguageProfile(langId))
                        {
                            continue;
                        }

                        hr = EnumLanguageProfiles(profilesPtr, langId, out var enumProfilesPtr);
                        if (hr < 0 || enumProfilesPtr == IntPtr.Zero)
                        {
                            continue;
                        }

                        try
                        {
                            while (true)
                            {
                                var items = new TF_LANGUAGEPROFILE[1];
                                hr = EnumLanguageProfilesNext(enumProfilesPtr, 1, items, out var fetched);
                                if (hr != 0 || fetched == 0)
                                {
                                    break;
                                }

                                var item = items[0];
                                var enabledHr = IsEnabledLanguageProfile(
                                    profilesPtr,
                                    ref item.clsid,
                                    item.langid,
                                    ref item.guidProfile,
                                    out var enabled);
                                if (enabledHr < 0 || enabled == 0)
                                {
                                    continue;
                                }

                                var name = GetTsfProfileDescription(profilesPtr, item);
                                if (string.IsNullOrWhiteSpace(name))
                                {
                                    continue;
                                }

                                name = NormalizeImeDisplayName(name);
                                if (IsNoiseImeName(name))
                                {
                                    continue;
                                }

                                var category = InferCategoryFromLangId(item.langid)
                                    ?? InferCategoryFromName(name)
                                    ?? ImeCategory.ChineseSimplified;

                                candidates.Add(new ScannedImeCandidate(name, category, 10));
                            }
                        }
                        finally
                        {
                            _ = Marshal.Release(enumProfilesPtr);
                        }
                    }
                }
                finally
                {
                    CoTaskMemFree(langPtr);
                }
            }
            finally
            {
                _ = Marshal.Release(profilesPtr);
            }
        }
        catch (COMException ex)
        {
            warning = $"COMException 0x{ex.HResult:X8}: {ex.Message}";
            return candidates;
        }
        catch (Exception ex)
        {
            warning = $"TSF异常: {ex.Message}";
            return candidates;
        }
        finally
        {
            if (shouldUninitialize)
            {
                CoUninitialize();
            }
        }

        return candidates;
    }

    private static IntPtr CreateInputProcessorProfilesCom()
    {
        // CLSID_TF_InputProcessorProfiles
        var clsid = new Guid("33C53A50-F456-4884-B049-85FD643ECFED");
        var iid = new Guid("1F02B6C5-7842-4EE6-8A0B-9A24183A95CA");
        var hr = CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX_INPROC_SERVER, ref iid, out var ptr);
        if (hr < 0)
        {
            return IntPtr.Zero;
        }

        return ptr;
    }

    private static string? GetTsfProfileDescription(IntPtr profilesPtr, TF_LANGUAGEPROFILE item)
    {
        var hr = GetLanguageProfileDescription(profilesPtr, ref item.clsid, item.langid, ref item.guidProfile, out var bstrPtr);
        if (hr >= 0 && bstrPtr != IntPtr.Zero)
        {
            try
            {
                var tsfDescription = Marshal.PtrToStringBSTR(bstrPtr);
                if (!string.IsNullOrWhiteSpace(tsfDescription))
                {
                    return tsfDescription;
                }
            }
            finally
            {
                SysFreeString(bstrPtr);
            }
        }

        return null;
    }

    private static void UpsertCandidate(IDictionary<string, ScannedImeCandidate> candidates, ScannedImeCandidate candidate)
    {
        if (!candidates.TryGetValue(candidate.DisplayName, out var existing) || candidate.Confidence > existing.Confidence)
        {
            candidates[candidate.DisplayName] = candidate;
        }
    }

    private static bool IsTargetLanguageProfile(ushort? langId) => langId is
        0x0804 or // zh-CN
        0x0404 or // zh-TW
        0x0C04 or // zh-HK
        0x1004 or // zh-SG
        0x1404 or // zh-MO
        0x0411;   // ja-JP

    private static string NormalizeImeDisplayName(string name)
    {
        if (name.Contains("wetype", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "微信输入法", StringComparison.OrdinalIgnoreCase))
        {
            return "微信输入法";
        }

        return name.Trim();
    }

    private static bool IsNoiseImeName(string name) =>
        name.Contains("输入体验", StringComparison.OrdinalIgnoreCase) ||
        name.Contains("Input Experience", StringComparison.OrdinalIgnoreCase);

    private static ImeCategory? InferCategoryFromLangId(ushort? langId) => langId switch
    {
        0x0804 or 0x1004 => ImeCategory.ChineseSimplified,
        0x0404 or 0x0C04 or 0x1404 => ImeCategory.ChineseTraditional,
        0x0411 => ImeCategory.Japanese,
        _ => null
    };

    private static ImeCategory? InferCategoryFromName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        if (name.Contains("速成", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("倉頡", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("仓颉", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("注音", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Quick", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Cangjie", StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.ChineseTraditional;
        }

        if (name.Contains("Japanese", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("日文", StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.Japanese;
        }

        if (name.Contains("拼音", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("五笔", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("微信输入法", StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.ChineseSimplified;
        }

        return null;
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

    private async Task<bool> ConfirmAddAsync(int targetCount)
    {
        var dialog = new ContentDialog
        {
            Title = "未发现配置文件",
            Content = $"将新增 {targetCount} 个 ime_config.xml，是否继续？",
            PrimaryButtonText = "新增",
            CloseButtonText = "取消",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = XamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    private void LoadSavedCustomIme()
    {
        var settings = LoadSettings();
        if (settings?.Ime is null || settings.Ime.Count == 0)
        {
            return;
        }

        foreach (var savedIme in settings.Ime)
        {
            if (string.IsNullOrWhiteSpace(savedIme.Name))
            {
                continue;
            }

            if (InputMethods.Any(item => string.Equals(item.DisplayName, savedIme.Name, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var category = savedIme.Category switch
            {
                "ChineseTraditional" => ImeCategory.ChineseTraditional,
                "Japanese" => ImeCategory.Japanese,
                _ => ImeCategory.ChineseSimplified
            };

            InputMethods.Add(new InputMethodItem(savedIme.Name, category, isCustom: true));
        }
    }

    private string? LoadSavedGameDir()
    {
        var settings = LoadSettings();
        return string.IsNullOrWhiteSpace(settings?.GameDir) ? null : settings.GameDir;
    }

    private AppSettings? LoadSettings()
    {
        try
        {
            var path = GetSettingsPath();
            if (!File.Exists(path))
            {
                return null;
            }

            var json = File.ReadAllText(path, Encoding.UTF8);
            return JsonSerializer.Deserialize(json, AppJsonContext.Default.AppSettings);
        }
        catch
        {
            return null;
        }
    }

    private void SaveSettings()
    {
        try
        {
            var settings = new AppSettings
            {
                GameDir = GameRootPathBox.Text?.Trim(),
                Ime = InputMethods
                    .Where(item => item.IsCustom)
                    .Select(item => new SavedIme
                    {
                        Name = item.DisplayName,
                        Category = item.Category.ToString()
                    })
                    .ToList()
            };

            var path = GetSettingsPath();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(settings, AppJsonContext.Default.AppSettings);
            File.WriteAllText(path, json, new UTF8Encoding(false));
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    private static string GetSettingsPath()
    {
        var settingsDirectory = GetSettingsDirectory();
        return Path.Combine(settingsDirectory, "config.json");
    }

    private static string GetSettingsDirectory()
    {
        if (IsPackagedApp())
        {
            return ApplicationData.Current.LocalFolder.Path;
        }

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "wows-ime");
    }

    private static bool IsPackagedApp()
    {
        try
        {
            _ = Windows.ApplicationModel.Package.Current;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void ShowStatus(string message, InfoBarSeverity severity)
    {
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;
    }

    private static int GetLanguageList(IntPtr profilesPtr, out IntPtr langPtr, out uint langCount)
    {
        // ITfInputProcessorProfiles::GetLanguageList is vtable slot 15 (IUnknown + 12 methods before it).
        var fn = GetVtableDelegate<TfGetLanguageListDelegate>(profilesPtr, 15);
        return fn(profilesPtr, out langPtr, out langCount);
    }

    private static int EnumLanguageProfiles(IntPtr profilesPtr, ushort langId, out IntPtr enumProfilesPtr)
    {
        // ITfInputProcessorProfiles::EnumLanguageProfiles is vtable slot 16.
        var fn = GetVtableDelegate<TfEnumLanguageProfilesDelegate>(profilesPtr, 16);
        return fn(profilesPtr, langId, out enumProfilesPtr);
    }

    private static int GetLanguageProfileDescription(IntPtr profilesPtr, ref Guid clsid, ushort langId, ref Guid profileGuid, out IntPtr bstrPtr)
    {
        // ITfInputProcessorProfiles::GetLanguageProfileDescription is vtable slot 12.
        var fn = GetVtableDelegate<TfGetLanguageProfileDescriptionDelegate>(profilesPtr, 12);
        return fn(profilesPtr, ref clsid, langId, ref profileGuid, out bstrPtr);
    }

    private static int EnumLanguageProfilesNext(IntPtr enumProfilesPtr, uint count, TF_LANGUAGEPROFILE[] buffer, out uint fetched)
    {
        // IEnumTfLanguageProfiles::Next is vtable slot 4.
        var fn = GetVtableDelegate<TfEnumLanguageProfilesNextDelegate>(enumProfilesPtr, 4);
        return fn(enumProfilesPtr, count, buffer, out fetched);
    }

    private static int IsEnabledLanguageProfile(IntPtr profilesPtr, ref Guid clsid, ushort langId, ref Guid profileGuid, out int enabled)
    {
        // ITfInputProcessorProfiles::IsEnabledLanguageProfile is vtable slot 18.
        var fn = GetVtableDelegate<TfIsEnabledLanguageProfileDelegate>(profilesPtr, 18);
        return fn(profilesPtr, ref clsid, langId, ref profileGuid, out enabled);
    }

    private static T GetVtableDelegate<T>(IntPtr comPtr, int methodIndex) where T : Delegate
    {
        var vtable = Marshal.ReadIntPtr(comPtr);
        var methodPtr = Marshal.ReadIntPtr(vtable, methodIndex * IntPtr.Size);
        return Marshal.GetDelegateForFunctionPointer<T>(methodPtr);
    }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageListDelegate(IntPtr @this, out IntPtr ppLangId, out uint pulCount);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesDelegate(IntPtr @this, ushort langid, out IntPtr ppEnum);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageProfileDescriptionDelegate(
        IntPtr @this,
        ref Guid rclsid,
        ushort langid,
        ref Guid guidProfile,
        out IntPtr pbstrProfile);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesNextDelegate(
        IntPtr @this,
        uint ulCount,
        [Out] TF_LANGUAGEPROFILE[] pProfile,
        out uint pcFetch);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfIsEnabledLanguageProfileDelegate(
        IntPtr @this,
        ref Guid rclsid,
        ushort langid,
        ref Guid guidProfile,
        out int pfEnable);

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out IntPtr ppv);

    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr pv);

    [DllImport("ole32.dll")]
    private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    [DllImport("oleaut32.dll")]
    private static extern void SysFreeString(IntPtr bstr);

    private const uint COINIT_APARTMENTTHREADED = 0x2;
    private const uint CLSCTX_INPROC_SERVER = 0x1;
}

public enum ImeCategory
{
    ChineseSimplified = 0,
    ChineseTraditional = 1,
    Japanese = 2
}

public sealed class InputMethodItem : Microsoft.UI.Xaml.DependencyObject
{
    public InputMethodItem(string displayName, ImeCategory initialCategory = ImeCategory.ChineseSimplified, bool isCustom = false)
    {
        DisplayName = displayName;
        Category = initialCategory;
        CategoryIndex = (int)initialCategory;
        IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public bool IsCustom { get; }
    public Visibility DeleteButtonVisibility => IsCustom ? Visibility.Visible : Visibility.Collapsed;

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

public sealed record ScannedImeCandidate(string DisplayName, ImeCategory Category, int Confidence);

public sealed class AppSettings
{
    public string? GameDir { get; set; }
    public List<SavedIme> Ime { get; set; } = new();
}

public sealed class SavedIme
{
    public string? Name { get; set; }
    public string Category { get; set; } = nameof(ImeCategory.ChineseSimplified);
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(AppSettings))]
internal partial class AppJsonContext : JsonSerializerContext
{
}

[StructLayout(LayoutKind.Sequential)]
internal struct TF_LANGUAGEPROFILE
{
    public Guid clsid;
    public ushort langid;
    public Guid catid;
    public int fActive;
    public Guid guidProfile;
}
