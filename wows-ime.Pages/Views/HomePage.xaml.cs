using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using wows_ime.Core.Models;
using wows_ime.Pages.Abstractions;
using wows_ime.Pages.Models;

namespace wows_ime.Pages.Views;

public sealed partial class HomePage : Page, INotifyPropertyChanged
{
    private const string SteamDefaultPath = @"C:\Program Files (x86)\Steam\steamapps\common\World of Warships";
    private const string LestaDefaultPath = @"C:\Games\Korabli";
    private const string Cn360DefaultPath = @"C:\Games\World_of_Warships_CN360";
    private const int TotalSteps = 3;
    private IPageHost host = null!;
    private string? lastScanWarning;
    private bool suppressSettingsSave;
    private int currentStep = 1;
    private string currentSelectedGamePathText = string.Empty;
    private DispatcherTimer? statusTimer;

    public event PropertyChangedEventHandler? PropertyChanged;
    public ObservableCollection<InputMethodItem> InputMethods { get; } = new();
    public ObservableCollection<InputMethodItem> ConfirmSelectedInputMethods { get; } = new();
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
        InitializeComponent();
    }

    protected override void OnNavigatedTo(Microsoft.UI.Xaml.Navigation.NavigationEventArgs e)
    {
        base.OnNavigatedTo(e);
        host = e.Parameter as IPageHost ?? throw new InvalidOperationException("HomePage requires an IPageHost navigation parameter.");
        suppressSettingsSave = true;
        ApplyLocalization();
        LoadGamePathOptions();
        LoadInputMethods();
        LoadSavedCustomIme();
        suppressSettingsSave = false;
        PrevButton.Content = SR("Step/Previous");
        NextButton.Content = SR("Step/Next");
        UpdateStepVisibility();
    }

    private void ApplyLocalization()
    {
        GameRootLabel.Text = SR("GameRootLabel/Text");
        AddCustomGamePathButton.Content = SR("AddCustomGamePathButton/Content");
        OpenFolderButton.Content = SR("OpenFolderButton/Content");
        ImeSectionTitle.Text = SR("ImeSectionTitle/Text");
        AddCustomImeButton.Content = SR("AddCustomImeButton/Content");
        RefreshImeButton.Content = SR("RefreshImeButton/Content");
        ImeTableHeaderName.Text = SR("ImeTableHeaderName/Text");
        ImeTableHeaderCategory.Text = SR("ImeTableHeaderCategory/Text");
        ImeTableHeaderAction.Text = SR("ImeTableHeaderAction/Text");
        ConfirmGameRootLabel.Text = SR("GameRootLabel/Text");
        ConfirmImeSectionHeader.Text = SR("ConfirmImeSectionHeader/Text");
    }

    private void UpdateStepVisibility()
    {
        GamePathSection.Visibility = currentStep == 1 ? Visibility.Visible : Visibility.Collapsed;
        ImeSection.Visibility = currentStep == 2 ? Visibility.Visible : Visibility.Collapsed;
        ConfirmSection.Visibility = currentStep == TotalSteps ? Visibility.Visible : Visibility.Collapsed;
        PrevButton.Visibility = currentStep > 1 ? Visibility.Visible : Visibility.Collapsed;
        NextButton.Content = currentStep == TotalSteps ? SR("WriteConfigButton/Content") : SR("Step/Next");
        StepProgressText.Text = SRF("Step/Progress", currentStep, TotalSteps);

        if (currentStep == TotalSteps)
        {
            UpdateConfirmSummary();
        }
    }

    private void UpdateConfirmSummary()
    {
        ConfirmGameRootPath.Text = GamePathOptions.FirstOrDefault(o => o.IsSelected)?.Path ?? SR("Status/GameRootInvalid");
        ConfirmSelectedInputMethods.Clear();
        var number = 1;
        foreach (var item in InputMethods.Where(item => item.IsSelected))
        {
            item.ConfirmationNumber = number++;
            ConfirmSelectedInputMethods.Add(item);
        }
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
        if (currentStep == TotalSteps)
        {
            WriteConfigButton_Click(sender, e);
            return;
        }

        if (currentStep == 1 && string.IsNullOrEmpty(GetSelectedGameRootPath()))
        {
            ShowStatus(SR("Status/GameRootInvalid"), InfoBarSeverity.Warning);
            return;
        }

        if (currentStep == 2 && !InputMethods.Any(i => i.IsSelected))
        {
            ShowStatus(SR("Status/SelectAtLeastOneIme"), InfoBarSeverity.Warning);
            return;
        }

        currentStep++;
        UpdateStepVisibility();
    }

    private async void AddCustomGamePathButton_Click(object sender, RoutedEventArgs e)
    {
        var path = (await host.Window.PickSingleFolderAsync())?.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

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

        var option = new GamePathOption(displayName, path, host.Localization, isCustom: true) { IsSelected = true };
        GamePathOptions.Add(option);
        SelectGamePath(option);
        ShowStatus(SR("Status/CustomGamePathAdded"), InfoBarSeverity.Success);
        SaveSelectedGamePath();
        SaveCustomGamePaths();
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
        if (sender is RadioButton { Tag: GamePathOption option })
        {
            SelectGamePath(option);
            SaveSelectedGamePath();
        }
    }

    private async void DeleteCustomGamePathButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: GamePathOption option } || !option.IsCustom)
        {
            return;
        }

        var dialog = CreateConfirmDialog("Dialog/DeleteCustomGamePath/Title", SRF("Dialog/DeleteCustomGamePath/Content", option.DisplayName, option.Path), "Dialog/DeleteCustomGamePath/PrimaryButton");
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

    private async void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var gameRoot = GetSelectedGameRootPath();
        if (string.IsNullOrWhiteSpace(gameRoot) || !host.Configuration.DirectoryExists(gameRoot))
        {
            ShowStatus(SR("Status/DirectoryNotExistsCannotOpen"), InfoBarSeverity.Warning);
            return;
        }

        if (!await host.Configuration.OpenFolderAsync(gameRoot))
        {
            ShowStatus(SRF("Status/OpenDirectoryFailed", gameRoot), InfoBarSeverity.Error);
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
        return new ContentDialog { Title = SR("Dialog/AddCustomIme/Title"), Content = panel, PrimaryButtonText = SR("Dialog/AddCustomIme/PrimaryButton"), CloseButtonText = SR("Dialog/Common/Cancel"), DefaultButton = ContentDialogButton.Primary, XamlRoot = XamlRoot };
    }

    private ContentDialog CreateConfirmDialog(string titleKey, string content, string primaryKey, bool primaryIsDefault = false) => new()
    {
        Title = SR(titleKey),
        Content = content,
        PrimaryButtonText = SR(primaryKey),
        CloseButtonText = SR("Dialog/Common/Cancel"),
        DefaultButton = primaryIsDefault ? ContentDialogButton.Primary : ContentDialogButton.Close,
        XamlRoot = XamlRoot
    };

    private async void AddCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = CreateAddCustomImeDialog();
        if (await dialog.ShowAsync() != ContentDialogResult.Primary)
        {
            return;
        }

        var panel = (StackPanel)dialog.Content;
        var name = ((TextBox)panel.Children[1]).Text?.Trim() ?? string.Empty;
        var categoryIndex = ((ComboBox)panel.Children[3]).SelectedIndex;
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

        var newItem = new InputMethodItem(name, host.Localization, isCustom: true) { IsSelected = true, CategoryIndex = categoryIndex < 0 ? 0 : categoryIndex };
        InputMethods.Add(newItem);
        ShowStatus(SR("Status/CustomImeAdded"), InfoBarSeverity.Success);
        SaveCustomInputMethods();
    }

    private async void DeleteCustomImeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: InputMethodItem item } || !item.IsCustom)
        {
            return;
        }

        var dialog = CreateConfirmDialog("Dialog/DeleteCustomIme/Title", SRF("Dialog/DeleteCustomIme/Content", item.DisplayName), "Dialog/DeleteCustomIme/PrimaryButton");
        if (await dialog.ShowAsync() == ContentDialogResult.Primary)
        {
            _ = InputMethods.Remove(item);
            ShowStatus(SR("Status/CustomImeDeleted"), InfoBarSeverity.Success);
            SaveCustomInputMethods();
        }
    }

    private void ImeCategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox { Tag: InputMethodItem item, SelectedIndex: >= 0 } comboBox)
        {
            if (item.CategoryIndex == comboBox.SelectedIndex)
            {
                return;
            }

            item.CategoryIndex = comboBox.SelectedIndex;
            SaveCustomInputMethods();
        }
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
        if (string.IsNullOrWhiteSpace(gameRoot) || !host.Configuration.DirectoryExists(gameRoot))
        {
            ShowStatus(SR("Status/GameRootInvalid"), InfoBarSeverity.Error);
            return;
        }

        if (!host.Configuration.HasGameExecutable(gameRoot))
        {
            ShowStatus(SR("Status/GameExeNotFound"), InfoBarSeverity.Error);
            return;
        }

        var targetFiles = host.Configuration.ResolveTargetConfigFiles(gameRoot);
        if (targetFiles.Count == 0)
        {
            ShowStatus(SR("Status/BinVersionDirectoryNotFound"), InfoBarSeverity.Error);
            return;
        }

        var existing = host.Configuration.GetExistingFiles(targetFiles);
        var dialog = existing.Count > 0
            ? CreateConfirmDialog("Dialog/Overwrite/Title", SRF("Dialog/Overwrite/Content", existing.Count), "Dialog/Overwrite/PrimaryButton")
            : CreateConfirmDialog("Dialog/AddConfig/Title", SRF("Dialog/AddConfig/Content", targetFiles.Count), "Dialog/AddConfig/PrimaryButton", primaryIsDefault: true);
        if (await dialog.ShowAsync() != ContentDialogResult.Primary)
        {
            ShowStatus(SR("Status/WriteCanceled"), InfoBarSeverity.Informational);
            return;
        }

        try
        {
            await host.Configuration.WriteConfigFilesAsync(selectedIme.Select(item => new InputMethodDefinition(item.DisplayName, item.Category)), targetFiles);
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
        var scanResult = host.InputMethodScanner.Scan();
        lastScanWarning = scanResult.WarningCode is null
            ? null
            : scanResult.WarningArguments is { Count: > 0 } arguments
                ? string.Format(SR(scanResult.WarningCode), arguments.Cast<object>().ToArray())
                : SR(scanResult.WarningCode);
        foreach (var ime in scanResult.Candidates)
        {
            InputMethods.Add(new InputMethodItem(ime.DisplayName, host.Localization, ime.Category));
        }

        if (InputMethods.Count == 0 && !string.IsNullOrWhiteSpace(lastScanWarning))
        {
            ShowStatus(SRF("Status/ScanCompletedNoImeWithWarning", lastScanWarning), InfoBarSeverity.Warning);
            return;
        }

        ShowStatus(SRF("Status/ScanCompletedWithCount", InputMethods.Count), InfoBarSeverity.Success);
    }

    private void LoadSavedCustomIme()
    {
        foreach (var savedIme in host.Settings.LoadCustomInputMethods())
        {
            if (string.IsNullOrWhiteSpace(savedIme.DisplayName) || InputMethods.Any(item => string.Equals(item.DisplayName, savedIme.DisplayName, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var category = savedIme.Category switch { "ChineseTraditional" => ImeCategory.ChineseTraditional, "Japanese" => ImeCategory.Japanese, _ => ImeCategory.ChineseSimplified };
            InputMethods.Add(new InputMethodItem(savedIme.DisplayName, host.Localization, category, isCustom: true));
        }
    }

    private void LoadGamePathOptions()
    {
        GamePathOptions.Clear();
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Steam"), SteamDefaultPath, host.Localization));
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Lesta"), LestaDefaultPath, host.Localization));
        GamePathOptions.Add(new GamePathOption(SR("GamePathOption/Cn360"), Cn360DefaultPath, host.Localization));
        foreach (var customPath in host.Settings.LoadCustomGamePaths())
        {
            var path = customPath.Path?.Trim();
            if (string.IsNullOrWhiteSpace(path) || GamePathOptions.Any(item => string.Equals(item.Path, path, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var displayName = string.IsNullOrWhiteSpace(customPath.DisplayName) ? Path.GetFileName(path) : customPath.DisplayName;
            GamePathOptions.Add(new GamePathOption(string.IsNullOrWhiteSpace(displayName) ? path : displayName, path, host.Localization, isCustom: true));
        }

        var selected = GamePathOptions.FirstOrDefault(item => string.Equals(item.Path, host.Settings.LoadSelectedGamePath(), StringComparison.OrdinalIgnoreCase));
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

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    private string GetSelectedGameRootPath() => GamePathOptions.FirstOrDefault(item => item.IsSelected)?.Path?.Trim() ?? string.Empty;
    private void SaveSelectedGamePath() { if (!suppressSettingsSave) host.Settings.SaveSelectedGamePath(GetSelectedGameRootPath()); }
    private void SaveCustomGamePaths() { if (!suppressSettingsSave) host.Settings.SaveCustomGamePaths(GamePathOptions.Where(item => item.IsCustom).Select(item => new PersistedGamePath(item.DisplayName, item.Path))); }
    private void SaveCustomInputMethods() { if (!suppressSettingsSave) host.Settings.SaveCustomInputMethods(InputMethods.Where(item => item.IsCustom).Select(item => new PersistedInputMethod(item.DisplayName, item.Category.ToString()))); }

    private void ShowStatus(string message, InfoBarSeverity severity)
    {
        statusTimer?.Stop();
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;
        statusTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
        statusTimer.Tick += (_, _) => { statusTimer?.Stop(); StatusInfoBar.IsOpen = false; };
        statusTimer.Start();
    }

    private string SR(string key) => host.Localization.GetString(key);
    private string SRF(string key, params object[] args) => host.Localization.Format(key, args);
}
