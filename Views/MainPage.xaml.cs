using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using Microsoft.UI.Xaml.Controls;
using Windows.ApplicationModel.Resources;
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
    private bool suppressSettingsSave;
    private static readonly ResourceLoader ResourceLoader = new();

    public ObservableCollection<InputMethodItem> InputMethods { get; } = new();

    public MainPage()
    {
        suppressSettingsSave = true;
        InitializeComponent();
        GameRootPathBox.Text = LoadSavedGameDir() ?? SteamDefaultPath;
        LoadInputMethods();
        LoadSavedCustomIme();
        suppressSettingsSave = false;
    }

    private async void PickFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");

        if (App.MainWindow is null)
        {
            ShowStatus(SR("Status/WindowHandleUnavailable"), InfoBarSeverity.Error);
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
            ShowStatus(SR("Status/DirectoryNotExistsCannotOpen"), InfoBarSeverity.Warning);
            return;
        }

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{gameRoot}\"",
                UseShellExecute = true
            };

            _ = Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ShowStatus(SRF("Status/OpenDirectoryFailed", ex.Message), InfoBarSeverity.Error);
        }
    }

    private async void AddCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        var nameBox = new TextBox
        {
            PlaceholderText = SR("Dialog/AddCustomIme/Placeholder")
        };

        var categoryCombo = new ComboBox
        {
            SelectedIndex = 0
        };
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/ChineseSimplified") });
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/ChineseTraditional") });
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/Japanese") });

        var panel = new StackPanel { Spacing = 8 };
        panel.Children.Add(new TextBlock { Text = SR("Dialog/AddCustomIme/NameLabel") });
        panel.Children.Add(nameBox);
        panel.Children.Add(new TextBlock { Text = SR("Dialog/AddCustomIme/CategoryLabel") });
        panel.Children.Add(categoryCombo);

        var dialog = new ContentDialog
        {
            Title = SR("Dialog/AddCustomIme/Title"),
            Content = panel,
            PrimaryButtonText = SR("Dialog/AddCustomIme/PrimaryButton"),
            CloseButtonText = SR("Dialog/Common/Cancel"),
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
            ShowStatus(SR("Status/ImeNameEmpty"), InfoBarSeverity.Warning);
            return;
        }

        if (InputMethods.Any(item => string.Equals(item.DisplayName, name, StringComparison.OrdinalIgnoreCase)))
        {
            ShowStatus(SR("Status/ImeNameExists"), InfoBarSeverity.Warning);
            return;
        }

        var newItem = new InputMethodItem(name, isCustom: true)
        {
            IsSelected = true,
            CategoryIndex = categoryCombo.SelectedIndex < 0 ? 0 : categoryCombo.SelectedIndex
        };

        InputMethods.Add(newItem);
        ShowStatus(SR("Status/CustomImeAdded"), InfoBarSeverity.Success);
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
        ShowStatus(SR("Status/CustomImeDeleted"), InfoBarSeverity.Success);
        SaveSettings();
    }

    private async void WriteConfigButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedIme = InputMethods.Where(item => item.IsSelected).ToList();
        if (selectedIme.Count == 0)
        {
            ShowStatus(SR("Status/SelectAtLeastOneIme"), InfoBarSeverity.Warning);
            return;
        }

        var gameRoot = GameRootPathBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(gameRoot) || !Directory.Exists(gameRoot))
        {
            ShowStatus(SR("Status/GameRootInvalid"), InfoBarSeverity.Error);
            return;
        }

        if (!HasGameExecutable(gameRoot))
        {
            ShowStatus(SR("Status/GameExeNotFound"), InfoBarSeverity.Error);
            return;
        }

        var targetFiles = ResolveTargetConfigFiles(gameRoot);
        if (targetFiles.Count == 0)
        {
            ShowStatus(SR("Status/BinVersionDirectoryNotFound"), InfoBarSeverity.Error);
            return;
        }

        var existing = targetFiles.Where(File.Exists).ToList();
        if (existing.Count > 0)
        {
            var shouldOverwrite = await ConfirmOverwriteAsync(existing.Count);
            if (!shouldOverwrite)
            {
                ShowStatus(SR("Status/WriteCanceled"), InfoBarSeverity.Informational);
                return;
            }
        }
        else
        {
            var shouldAdd = await ConfirmAddAsync(targetFiles.Count);
            if (!shouldAdd)
            {
                ShowStatus(SR("Status/WriteCanceled"), InfoBarSeverity.Informational);
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

            ShowStatus(SRF("Status/WriteSucceededWithCount", targetFiles.Count), InfoBarSeverity.Success);
            SaveSettings();
        }
        catch (Exception ex)
        {
            ShowStatus(SRF("Status/WriteFailed", ex.Message), InfoBarSeverity.Error);
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
            ShowStatus(SRF("Status/ScanCompletedNoImeWithWarning", lastScanWarning), InfoBarSeverity.Warning);
            return;
        }

        ShowStatus(SRF("Status/ScanCompletedWithCount", InputMethods.Count), InfoBarSeverity.Success);
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
            warning = SRF("Tsf/CoInitializeFailed", $"0x{coInitHr:X8}");
            return candidates;
        }

        try
        {
            var profilesPtr = CreateInputProcessorProfilesCom();
            if (profilesPtr == IntPtr.Zero)
            {
                warning = SR("Tsf/CreateProfilesObjectFailed");
                return candidates;
            }

            try
            {
                var hr = GetLanguageList(profilesPtr, out var langPtr, out var langCount);
                if (hr < 0 || langPtr == IntPtr.Zero || langCount == 0)
                {
                    warning = hr < 0
                        ? SRF("Tsf/GetLanguageListFailed", $"0x{hr:X8}")
                        : SR("Tsf/GetLanguageListEmpty");
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
            warning = SRF("Tsf/ComException", $"0x{ex.HResult:X8}", ex.Message);
            return candidates;
        }
        catch (Exception ex)
        {
            warning = SRF("Tsf/GenericException", ex.Message);
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
            string.Equals(name, SR("Ime/Weixin"), StringComparison.OrdinalIgnoreCase))
        {
            return SR("Ime/Weixin");
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
            name.Contains(SR("Ime/Weixin"), StringComparison.OrdinalIgnoreCase))
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
            Title = SR("Dialog/Overwrite/Title"),
            Content = SRF("Dialog/Overwrite/Content", existingCount),
            PrimaryButtonText = SR("Dialog/Overwrite/PrimaryButton"),
            CloseButtonText = SR("Dialog/Common/Cancel"),
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
            Title = SR("Dialog/AddConfig/Title"),
            Content = SRF("Dialog/AddConfig/Content", targetCount),
            PrimaryButtonText = SR("Dialog/AddConfig/PrimaryButton"),
            CloseButtonText = SR("Dialog/Common/Cancel"),
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
        if (suppressSettingsSave)
        {
            return;
        }

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

    private static string SR(string key)
    {
        var value = ResourceLoader.GetString(key);
        return string.IsNullOrEmpty(value) ? key : value;
    }

    private static string SRF(string key, params object[] args)
    {
        return string.Format(CultureInfo.CurrentUICulture, SR(key), args);
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


