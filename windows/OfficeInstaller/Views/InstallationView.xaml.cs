using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeInstaller.Core;
using OfficeInstaller.Services;

namespace OfficeInstaller.Views;

public partial class InstallationView : UserControl
{
    private readonly OfficeDeploymentService deployment = new();
    private DeploymentSession? session;
    private string? selectedTool;
    private bool busy;
    private bool initialized;
    private bool installing;
    private bool checking;
    private bool componentsChecked;
    private readonly Dictionary<string, CheckBox> componentChecks = new(StringComparer.Ordinal);
    private readonly Dictionary<string, TextBlock> componentLabels = new(StringComparer.Ordinal);
    private readonly Dictionary<string, bool> componentAvailability = new(StringComparer.Ordinal);
    private HashSet<string> foundApplications = new(StringComparer.OrdinalIgnoreCase);
    public event Action<bool>? BusyChanged;
    public event Action<string>? StatusChanged;

    public InstallationView()
    {
        InitializeComponent();
        LanguageBox.ItemsSource = OfficeCatalog.Languages;
        LanguageBox.SelectedIndex = 0;
        EditionBox.ItemsSource = OfficeCatalog.Editions;
        initialized = true;
        EditionBox.SelectedIndex = 0;
    }

    private void Edition_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!initialized || EditionBox.SelectedItem is not OfficeEdition edition) return;
        componentsChecked = false;
        foundApplications.Clear();
        ApplicationsPanel.Children.Clear();
        OptionalComponentsPanel.Children.Clear();
        componentChecks.Clear();
        componentLabels.Clear();
        componentAvailability.Clear();
        OfficeTitle.Text = edition.Release == "Microsoft 365" ? "Office Installer Microsoft 365" : "Office Installer LTSC " + edition.Release;
        foreach (var id in new[] { "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access" })
            AddComponent(ApplicationsPanel, id, edition.Applications.Contains(id));
        foreach (var id in new[] { "OneDrive", "Teams", "Copilot", "Defender", "Publisher" })
            AddComponent(OptionalComponentsPanel, id, edition.Applications.Contains(id));
        ComponentStatus.Text = "Компоненты ещё не проверены. Доступность определяется выбранной редакцией Office.";
        InvalidateSession();
    }

    private void AddComponent(Panel panel, string id, bool available)
    {
        var row = new Grid { Margin = new Thickness(0, 6, 0, 6) };
        row.ColumnDefinitions.Add(new ColumnDefinition());
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var check = new CheckBox { Content = "Microsoft " + id + (id is "Copilot" or "Defender" ? " (компонент)" : ""),
            Tag = id, IsChecked = false, IsEnabled = available, VerticalAlignment = VerticalAlignment.Center };
        check.Checked += Application_Changed;
        check.Unchecked += Application_Changed;
        var label = new TextBlock { FontSize = 12, Foreground = (Brush)FindResource("SecondaryTextBrush"),
            Text = available ? "Не проверено" : id is "Teams" or "Copilot" or "Defender" ? "Отдельная установка" : "Недоступно",
            Margin = new Thickness(12, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(label, 1);
        row.Children.Add(check);
        row.Children.Add(label);
        panel.Children.Add(row);
        componentChecks.Add(id, check);
        componentLabels.Add(id, label);
        componentAvailability.Add(id, available);
    }

    private void UpdateAvailability()
    {
        componentsChecked = true;
        foreach (var (id, check) in componentChecks)
        {
            if (foundApplications.Contains(id))
            {
                componentLabels[id].Text = "Файл найден";
                componentLabels[id].Foreground = new SolidColorBrush(Color.FromRgb(36, 132, 62));
                continue;
            }
            if (!componentAvailability[id]) continue;
            componentLabels[id].Text = "Доступно";
            componentLabels[id].Foreground = (Brush)FindResource("AccentBrush");
        }
        ComponentStatus.Text = $"Проверка завершена. Доступных компонентов в выбранной редакции: {componentAvailability.Values.Count(available => available)}.";
        InstallButton.IsEnabled = session is not null && componentsChecked;
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        UpdateAvailability();
        foreach (var (id, check) in componentChecks) check.IsChecked = componentAvailability[id];
    }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var check in componentChecks.Values) check.IsChecked = false;
    }

    private async void CheckComponents_Click(object sender, RoutedEventArgs e)
    {
        if (busy) return;
        try
        {
            checking = true;
            SetBusy(true);
            ComponentStatus.Text = "Проверка компонентов…";
            var scan = await Task.Run(new OfficeDetectionService().Scan);
            foundApplications = scan.ApplicationFiles.Select(f => f.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (session is not null)
            {
                session = await deployment.LoadSessionAsync(Path.Combine(session.Directory, "prepared.json"));
                MarkPackageVerified();
            }
            UpdateAvailability();
            if (session is null) ComponentStatus.Text += " Пакеты ещё не скачаны; подпись будет проверена после загрузки.";
            if (scan.Warnings.Count > 0) ComponentStatus.Text += " Сведения об установленных файлах неполные; замечания доступны при проверке Office во вкладке восстановления.";
            StatusChanged?.Invoke(ComponentStatus.Text);
        }
        catch (Exception exception)
        {
            componentsChecked = false;
            InvalidateSession();
            ComponentStatus.Text = "Проверка не завершена: " + exception.Message;
            ShowError(exception.Message);
        }
        finally { SetBusy(false); checking = false; }
    }

    private void Selection_Changed(object sender, SelectionChangedEventArgs e) { if (initialized) InvalidateSession(); }
    private void Application_Changed(object sender, RoutedEventArgs e) => InvalidateSession();
    private void InvalidateSession()
    {
        session = null;
        InstallButton.IsEnabled = false;
        OpenSessionButton.IsEnabled = false;
        DownloadStatus.Text = "Настройки изменены. Подготовьте пакеты для текущего выбора. Ранее скачанные файлы сохраняются.";
        PackageStatus.Text = "Подпись ещё не проверена для текущего выбора.";
        PackageStatus.Foreground = (Brush)FindResource("SecondaryTextBrush");
        InstallationStatus.Text = "Для установки сначала подготовьте пакеты Office.";
    }

    private InstallationRequest Request()
    {
        if (EditionBox.SelectedItem is not OfficeEdition edition || LanguageBox.SelectedItem is not OfficeLanguage language)
            throw new ArgumentException("Выберите редакцию и язык Office.");
        var architecture = (ArchitectureBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() == "32"
            ? OfficeArchitecture.X86 : OfficeArchitecture.X64;
        var apps = componentChecks.Where(pair => componentAvailability[pair.Key] && pair.Value.IsChecked == true)
            .Select(pair => pair.Key).ToHashSet(StringComparer.Ordinal);
        var request = new InstallationRequest(edition.Id, architecture, language.Id, apps);
        OfficeDeployment.Validate(request);
        return request;
    }

    private void ChooseTool_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Title = "Выберите setup.exe из Office Deployment Tool", Filter = "Office Deployment Tool (setup.exe)|setup.exe", CheckFileExists = true };
        if (dialog.ShowDialog(Window.GetWindow(this)) != true) return;
        selectedTool = dialog.FileName;
        ToolPathText.Text = selectedTool;
        InvalidateSession();
    }

    private void AutomaticTool_Click(object sender, RoutedEventArgs e)
    {
        selectedTool = null;
        ToolPathText.Text = "Автоматическая загрузка ODT";
        InvalidateSession();
    }

    private void MicrosoftSite_Click(object sender, RoutedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo(OfficeDeploymentService.DownloadPage) { UseShellExecute = true })?.Dispose(); }
        catch (Exception exception) { ShowError(exception.Message); }
    }

    private void Preview_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var xml = OfficeDeployment.Configuration(Request(), @"C:\OfficeInstaller\Packages", download: false);
            var text = new TextBox { Text = xml, IsReadOnly = true, FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                TextWrapping = TextWrapping.NoWrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(16) };
            new Window { Title = "Предпросмотр конфигурации — пример пути пакетов", Owner = Window.GetWindow(this),
                Width = 720, Height = 600, Content = text, WindowStartupLocation = WindowStartupLocation.CenterOwner }.ShowDialog();
        }
        catch (Exception exception) { ShowError(exception.Message); }
    }

    private async void Download_Click(object sender, RoutedEventArgs e)
    {
        if (busy) return;
        try
        {
            var request = Request();
            var edition = OfficeDeployment.Validate(request);
            if (MessageBox.Show(Window.GetWindow(this), $"Скачать {edition.Name}, {(int)request.Architecture} бит, {request.Language}?\n\nПриложения: {string.Join(", ", request.Applications)}.\n\nБудут подготовлены ODT и пакеты Office в новой папке программы. Установка Office не запускается. Скачанные файлы автоматически не удаляются. Для загрузки через ODT может потребоваться подтверждение прав администратора.",
                "Подготовка пакетов", MessageBoxButton.OKCancel, MessageBoxImage.Information, MessageBoxResult.Cancel) != MessageBoxResult.OK) return;
            InvalidateSession();
            SetBusy(true);
            session = await deployment.PrepareAndDownloadAsync(request, selectedTool, new Progress<string>(Report));
            MarkPackageVerified();
            UpdateAvailability();
            OpenSessionButton.IsEnabled = true;
        }
        catch (OperationCanceledException exception) { Report(exception.Message); }
        catch (Exception exception) { Report("Подготовка не завершена: " + exception.Message); ShowError(exception.Message); }
        finally { if (busy) SetBusy(false); }
    }

    private async void Install_Click(object sender, RoutedEventArgs e)
    {
        if (busy || session is null || !componentsChecked) return;
        var current = session;
        if (MessageBox.Show(Window.GetWindow(this), $"Начать установку {OfficeDeployment.Validate(current.Request).Name}?\n\nПриложения: {string.Join(", ", current.Request.Applications)}.\n\nБудут изменены установленные программы Windows. Если обнаружится существующий Office, операция будет заблокирована. Автоматическое удаление Office и принудительное закрытие приложений не выполняются. Потребуется разрешение администратора и принятие условий Microsoft в установщике.",
            "Подтверждение установки Office", MessageBoxButton.OKCancel, MessageBoxImage.Warning, MessageBoxResult.Cancel) != MessageBoxResult.OK) return;
        try
        {
            installing = true;
            SetBusy(true);
            var code = await deployment.InstallAsync(current, new Progress<string>(Report));
            Report(code switch
            {
                0 => "ODT завершил работу без ошибки. Повторите диагностику и проверьте запуск приложений Office.",
                3010 => "ODT завершил работу; требуется перезагрузка. Выполните её самостоятельно, когда будете готовы.",
                1602 => "Установка отменена пользователем.",
                _ => $"ODT завершился с кодом {code}. Успешная установка не подтверждена. Журнал: {current.LogPath}"
            });
            // A new install attempt must use a newly prepared and rechecked session.
            InstallButton.IsEnabled = false;
        }
        catch (OperationCanceledException exception) { Report(exception.Message); }
        catch (Exception exception) { Report("Установка не завершена: " + exception.Message); ShowError(exception.Message); }
        finally { SetBusy(false); installing = false; }
    }

    private async void RestoreSession_Click(object sender, RoutedEventArgs e)
    {
        if (busy) return;
        var dialog = new OpenFileDialog { Title = "Выберите prepared.json в папке скачанных пакетов",
            Filter = "Описание пакетов (prepared.json)|prepared.json", CheckFileExists = true };
        if (Directory.Exists(OfficeDeploymentService.SessionsRoot)) dialog.InitialDirectory = OfficeDeploymentService.SessionsRoot;
        if (dialog.ShowDialog(Window.GetWindow(this)) != true) return;
        try
        {
            SetBusy(true);
            Report("Повторно проверяем ранее скачанные пакеты и подпись ODT…");
            var restored = await deployment.LoadSessionAsync(dialog.FileName);
            EditionBox.SelectedItem = OfficeCatalog.Editions.Single(e => e.Id == restored.Request.ProductId);
            LanguageBox.SelectedItem = OfficeCatalog.Languages.Single(l => l.Id == restored.Request.Language);
            ArchitectureBox.SelectedIndex = restored.Request.Architecture == OfficeArchitecture.X64 ? 0 : 1;
            foreach (var check in componentChecks.Values)
                check.IsChecked = restored.Request.Applications.Contains((string)check.Tag);
            session = restored;
            MarkPackageVerified();
            UpdateAvailability();
            OpenSessionButton.IsEnabled = true;
            Report("Скачанные пакеты проверены и готовы. Повторная загрузка не требуется. Каталог: " + restored.SourcePath);
        }
        catch (Exception exception) { InvalidateSession(); Report("Не удалось открыть пакеты: " + exception.Message); ShowError(exception.Message); }
        finally { SetBusy(false); }
    }

    private void OpenSession_Click(object sender, RoutedEventArgs e)
    {
        if (session is null) return;
        try
        {
            var start = new ProcessStartInfo(session.Directory) { UseShellExecute = true };
            Process.Start(start)?.Dispose();
        }
        catch (Exception exception) { ShowError(exception.Message); }
    }

    private void MarkPackageVerified()
    {
        PackageStatus.Text = "✓ Подпись Microsoft у ODT подтверждена. Пакеты Office подготовлены.";
        PackageStatus.Foreground = new SolidColorBrush(Color.FromRgb(36, 132, 62));
        InstallationStatus.Text = "Пакеты готовы. Можно установить выбранные компоненты.";
    }
    private void SetBusy(bool value)
    {
        busy = value;
        DownloadProgress.Visibility = value && !installing && !checking ? Visibility.Visible : Visibility.Collapsed;
        InstallProgress.Visibility = value && installing ? Visibility.Visible : Visibility.Collapsed;
        DownloadButton.Content = value && !installing && !checking ? "Загрузка…" : "Скачать Office";
        InstallButton.Content = value && installing ? "Установка…" : "Установить выбранное";
        CheckComponentsButton.Content = value && checking ? "Проверка…" : "Проверить компоненты";
        BusyChanged?.Invoke(value);
    }
    private void Report(string message)
    {
        if (installing) InstallationStatus.Text = message;
        else DownloadStatus.Text = message;
        StatusChanged?.Invoke(message);
    }
    private void ShowError(string message) => MessageBox.Show(Window.GetWindow(this), message, "Office Installer", MessageBoxButton.OK, MessageBoxImage.Error);
}
