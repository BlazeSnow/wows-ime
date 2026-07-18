using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Windows.Storage.Pickers;
using WinRT.Interop;
using wows_ime.Services;

namespace wows_ime.Views;

public sealed partial class HomePage : Page, INotifyPropertyChanged
{
    private const string SteamDefaultPath = @"C:\Program Files (x86)\Steam\steamapps\common\World of Warships";
    private const string LestaDefaultPath = @"C:\Games\Korabli";
    private const string Cn360DefaultPath = @"C:\Games\World_of_Warships_CN360";
    private string? lastScanWarning;
    private bool suppressSettingsSave;
    private int currentStep = 1;
    private string currentSelectedGamePathText = string.Empty;
    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<InputMethodItem> InputMethods { get; } = new();
    public ObservableCollection<GamePathOption> GamePathOptions { get; } = new();
    public string CurrentSelectedGamePathText
    {
        get => currentSelectedGamePathText;
        private set
        {
            if (currentSelectedGamePathText == value)
            {
                return;
            }

            currentSelectedGamePathText = value;
            OnPropertyChanged();
        }
    }

    public HomePage()
    {
        suppressSettingsSave = true;
        SettingsPersistence.Initialize();
        InitializeComponent();
        LoadGamePathOptions();
        LoadInputMethods();
        LoadSavedCustomIme();
        suppressSettingsSave = false;
        PrevButton.Content = SR("Step/Previous");
        NextButton.Content = SR("Step/Next");
        UpdateStepVisibility();
    }

    private void UpdateStepVisibility()
    {
        GamePathSection.Visibility = currentStep == 1 ? Visibility.Visible : Visibility.Collapsed;
        ImeSection.Visibility = currentStep == 2 ? Visibility.Visible : Visibility.Collapsed;
        ConfirmSection.Visibility = currentStep == 3 ? Visibility.Visible : Visibility.Collapsed;

        PrevButton.Visibility = currentStep > 1 ? Visibility.Visible : Visibility.Collapsed;
        NextButton.Visibility = currentStep < 3 ? Visibility.Visible : Visibility.Collapsed;

        if (currentStep == 3)
        {
            UpdateConfirmSummary();
        }
    }

    private void UpdateConfirmSummary()
    {
        var path = GamePathOptions.FirstOrDefault(o => o.IsSelected)?.DisplayName ?? SR("Status/GameRootInvalid");
        var imeCount = InputMethods.Count(i => i.IsSelected);
        var imeList = string.Join("\n", InputMethods.Where(i => i.IsSelected).Select(i => $"• {i.DisplayName}"));
        if (string.IsNullOrEmpty(imeList))
        {
            imeList = SR("Status/SelectAtLeastOneIme");
        }

        ConfirmSummary.Text = $"游戏路径: {path}\n\n已选输入法 ({imeCount}):\n{imeList}";
    }

    private void PrevButton_Click(object sender, RoutedEventArgs e)
    {
        if (currentStep > 1)
        {
            currentStep--;
            UpdateStepVisibility();
        }
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
        if (currentStep == 1)
        {
            if (string.IsNullOrEmpty(GetSelectedGameRootPath()))
            {
                ShowStatus(SR("Status/GameRootInvalid"), InfoBarSeverity.Warning);
                return;
            }
        }

        if (currentStep == 2)
        {
            if (!InputMethods.Any(i => i.IsSelected))
            {
                ShowStatus(SR("Status/SelectAtLeastOneIme"), InfoBarSeverity.Warning);
                return;
            }
        }

        if (currentStep < 3)
        {
            currentStep++;
            UpdateStepVisibility();
        }
    }

    private async void AddCustomGamePathButton_Click(object sender, RoutedEventArgs e)
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
            var path = folder.Path.Trim();
            var existing = GamePathOptions.FirstOrDefault(item => string.Equals(item.Path, path, StringComparison.OrdinalIgnoreCase));
            if (existing is not null)
            {
                SelectGamePath(existing);
                ShowStatus(SR("Status/CustomGamePathExists"), InfoBarSeverity.Informational);
                return;
            }

            var displayName = Path.GetFileName(path);
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = path;
            }

            var option = new GamePathOption(displayName, path, isCustom: true) { IsSelected = true };
            GamePathOptions.Add(option);
            SelectGamePath(option);
            ShowStatus(SR("Status/CustomGamePathAdded"), InfoBarSeverity.Success);
            SaveSelectedGamePath();
            SaveCustomGamePaths();
        }
    }

    private void RefreshImeButton_Click(object sender, RoutedEventArgs e)
    {
        suppressSettingsSave = true;
        try
        {
            LoadInputMethods();
            LoadSavedCustomIme();
        }
        finally
        {
            suppressSettingsSave = false;
        }
    }

    private void GamePathRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        if (sender is not RadioButton { Tag: GamePathOption option })
        {
            return;
        }

        SelectGamePath(option);
        SaveSelectedGamePath();
    }

    private async void DeleteCustomGamePathButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: GamePathOption option } || !option.IsCustom)
        {
            return;
        }

        var dialog = CreateConfirmDialog(
            "Dialog/DeleteCustomGamePath/Title",
            SRF("Dialog/DeleteCustomGamePath/Content", option.DisplayName, option.Path),
            "Dialog/DeleteCustomGamePath/PrimaryButton");

        if (await dialog.ShowAsync() != ContentDialogResult.Primary)
        {
            return;
        }

        var wasSelected = option.IsSelected;
        _ = GamePathOptions.Remove(option);
        if (wasSelected)
        {
            SelectGamePath(GamePathOptions.FirstOrDefault());
        }

        ShowStatus(SR("Status/CustomGamePathDeleted"), InfoBarSeverity.Success);
        SaveSelectedGamePath();
        SaveCustomGamePaths();
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var gameRoot = GetSelectedGameRootPath();
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

    private ContentDialog CreateAddCustomImeDialog()
    {
        var nameBox = new TextBox { PlaceholderText = SR("Dialog/AddCustomIme/Placeholder") };
        var categoryCombo = new ComboBox { SelectedIndex = 0 };
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/ChineseSimplified") });
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/ChineseTraditional") });
        categoryCombo.Items.Add(new ComboBoxItem { Content = SR("Category/Japanese") });

        var panel = new StackPanel { Spacing = 8 };
        panel.Children.Add(new TextBlock { Text = SR("Dialog/AddCustomIme/NameLabel") });
        panel.Children.Add(nameBox);
        panel.Children.Add(new TextBlock { Text = SR("Dialog/AddCustomIme/CategoryLabel") });
        panel.Children.Add(categoryCombo);

        return new ContentDialog
        {
            Title = SR("Dialog/AddCustomIme/Title"),
            Content = panel,
            PrimaryButtonText = SR("Dialog/AddCustomIme/PrimaryButton"),
            CloseButtonText = SR("Dialog/Common/Cancel"),
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = XamlRoot
        };
    }

    private ContentDialog CreateConfirmDialog(string titleKey, string content, string primaryKey, bool primaryIsDefault = false)
    {
        return new ContentDialog
        {
            Title = SR(titleKey),
            Content = content,
            PrimaryButtonText = SR(primaryKey),
            CloseButtonText = SR("Dialog/Common/Cancel"),
            DefaultButton = primaryIsDefault ? ContentDialogButton.Primary : ContentDialogButton.Close,
            XamlRoot = XamlRoot
        };
    }

    private async void AddCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = CreateAddCustomImeDialog();
        var result = await dialog.ShowAsync();

        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        var panel = (StackPanel)dialog.Content;
        var nameBox = (TextBox)panel.Children[1];
        var categoryCombo = (ComboBox)panel.Children[3];
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
        SaveCustomInputMethods();
    }

    private async void DeleteCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: InputMethodItem item })
        {
            return;
        }

        if (!item.IsCustom)
        {
            return;
        }

        var dialog = CreateConfirmDialog(
            "Dialog/DeleteCustomIme/Title",
            SRF("Dialog/DeleteCustomIme/Content", item.DisplayName),
            "Dialog/DeleteCustomIme/PrimaryButton");

        if (await dialog.ShowAsync() != ContentDialogResult.Primary)
        {
            return;
        }

        _ = InputMethods.Remove(item);
        ShowStatus(SR("Status/CustomImeDeleted"), InfoBarSeverity.Success);
        SaveCustomInputMethods();
    }

    private void ImeCategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox { Tag: InputMethodItem item, SelectedIndex: >= 0 } comboBox)
        {
            item.CategoryIndex = comboBox.SelectedIndex;
        }

        SaveCustomInputMethods();
    }

    private async void WriteConfigButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedIme = InputMethods.Where(item => item.IsSelected).ToList();
        if (selectedIme.Count == 0)
        {
            ShowStatus(SR("Status/SelectAtLeastOneIme"), InfoBarSeverity.Warning);
            return;
        }

        var gameRoot = GetSelectedGameRootPath();
        if (string.IsNullOrWhiteSpace(gameRoot) || !Directory.Exists(gameRoot))
        {
            ShowStatus(SR("Status/GameRootInvalid"), InfoBarSeverity.Error);
            return;
        }

        if (!GameConfigService.HasGameExecutable(gameRoot))
        {
            ShowStatus(SR("Status/GameExeNotFound"), InfoBarSeverity.Error);
            return;
        }

        var targetFiles = GameConfigService.ResolveTargetConfigFiles(gameRoot);
        if (targetFiles.Count == 0)
        {
            ShowStatus(SR("Status/BinVersionDirectoryNotFound"), InfoBarSeverity.Error);
            return;
        }

        var existing = targetFiles.Where(File.Exists).ToList();
        if (existing.Count > 0)
        {
            var dialog = CreateConfirmDialog(
                "Dialog/Overwrite/Title",
                SRF("Dialog/Overwrite/Content", existing.Count),
                "Dialog/Overwrite/PrimaryButton");

            if (await dialog.ShowAsync() != ContentDialogResult.Primary)
            {
                ShowStatus(SR("Status/WriteCanceled"), InfoBarSeverity.Informational);
                return;
            }
        }
        else
        {
            var dialog = CreateConfirmDialog(
                "Dialog/AddConfig/Title",
                SRF("Dialog/AddConfig/Content", targetFiles.Count),
                "Dialog/AddConfig/PrimaryButton",
                primaryIsDefault: true);

            if (await dialog.ShowAsync() != ContentDialogResult.Primary)
            {
                ShowStatus(SR("Status/WriteCanceled"), InfoBarSeverity.Informational);
                return;
            }
        }

        try
        {
            await GameConfigService.WriteConfigFilesAsync(targetFiles, selectedIme);
            ShowStatus(SRF("Status/WriteSucceededWithCount", targetFiles.Count), InfoBarSeverity.Success);
            SaveSelectedGamePath();
            SaveCustomInputMethods();
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

        foreach (var ime in InputMethodScanner.ReadCandidates(out warning))
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

    private void LoadSavedCustomIme()
    {
        foreach (var savedIme in SettingsPersistence.LoadCustomInputMethods())
        {
            if (string.IsNullOrWhiteSpace(savedIme.DisplayName))
            {
                continue;
            }

            if (InputMethods.Any(item => string.Equals(item.DisplayName, savedIme.DisplayName, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var category = savedIme.Category switch
            {
                "ChineseTraditional" => ImeCategory.ChineseTraditional,
                "Japanese" => ImeCategory.Japanese,
                _ => ImeCategory.ChineseSimplified
            };

            InputMethods.Add(new InputMethodItem(savedIme.DisplayName, category, isCustom: true));
        }
    }

    private void LoadGamePathOptions()
    {
        GamePathOptions.Clear();
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Steam"), SteamDefaultPath));
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Lesta"), LestaDefaultPath));
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Cn360"), Cn360DefaultPath));

        foreach (var customPath in SettingsPersistence.LoadCustomGamePaths())
        {
            var path = customPath.Path?.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }

            if (GamePathOptions.Any(item => string.Equals(item.Path, path, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var displayName = string.IsNullOrWhiteSpace(customPath.DisplayName) ? Path.GetFileName(path) : customPath.DisplayName;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = path;
            }

            GamePathOptions.Add(new GamePathOption(displayName, path, isCustom: true));
        }

        var selectedPath = SettingsPersistence.LoadSelectedGamePath();
        var selected = GamePathOptions.FirstOrDefault(item => string.Equals(item.Path, selectedPath, StringComparison.OrdinalIgnoreCase));
        SelectGamePath(selected ?? GamePathOptions.FirstOrDefault());
    }

    private void SelectGamePath(GamePathOption? option)
    {
        foreach (var item in GamePathOptions)
        {
            item.IsSelected = ReferenceEquals(item, option);
        }

        CurrentSelectedGamePathText = option?.Path ?? string.Empty;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private string GetSelectedGameRootPath()
    {
        return GamePathOptions.FirstOrDefault(item => item.IsSelected)?.Path?.Trim() ?? string.Empty;
    }

    private void SaveSelectedGamePath()
    {
        if (suppressSettingsSave)
        {
            return;
        }

        SettingsPersistence.SaveSelectedGamePath(GetSelectedGameRootPath());
    }

    private void SaveCustomGamePaths()
    {
        if (suppressSettingsSave)
        {
            return;
        }

        var customPaths = GamePathOptions
            .Where(item => item.IsCustom)
            .Select(item => new PersistedGamePath(item.DisplayName, item.Path));
        SettingsPersistence.SaveCustomGamePaths(customPaths);
    }

    private void SaveCustomInputMethods()
    {
        if (suppressSettingsSave)
        {
            return;
        }

        var customInputMethods = InputMethods
            .Where(item => item.IsCustom)
            .Select(item => new PersistedInputMethod(item.DisplayName, item.Category.ToString()));
        SettingsPersistence.SaveCustomInputMethods(customInputMethods);
    }

    private void ShowStatus(string message, InfoBarSeverity severity)
    {
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;

        var timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(5)
        };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            StatusInfoBar.IsOpen = false;
        };
        timer.Start();
    }

    private static string SR(string key)
    {
        return AppResources.Get(key);
    }

    private static string SRF(string key, params object[] args)
    {
        return AppResources.Format(key, args);
    }
}
