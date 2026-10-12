using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Win32;
using OfficeInstaller.Core;
using OfficeInstaller.Services;

namespace OfficeInstaller;

public partial class MainWindow : Window
{
    private readonly OfficeDetectionService detection = new();
    private bool scanning;
    private bool busy;
    private OfficeScanResult? lastScan;
    private DateTimeOffset lastScanTime;

    public MainWindow()
    {
        InitializeComponent();
        InstallView.BusyChanged += SetBusy;
        InstallView.StatusChanged += message => OperationStatus.Text = message;
        Closing += Window_Closing;
    }

    private async void Scan_Click(object sender, RoutedEventArgs e)
    {
        if (scanning || busy) return;
        scanning = true;
        SetBusy(true);
        lastScan = null;
        ExportButton.IsEnabled = false;
        ScanButton.IsEnabled = false;
        ScanButton.Content = "Проверка…";
        ScanProgress.Visibility = Visibility.Visible;
        StatusText.Text = "Читаем сведения об установке Office…";
        EmptyText.Visibility = Visibility.Collapsed;
        ProductsList.ItemsSource = null;
        FilesList.ItemsSource = null;
        WarningsList.ItemsSource = null;
        FilesHeading.Visibility = FilesNote.Visibility = Visibility.Collapsed;
        try
        {
            var result = await Task.Run(detection.Scan);
            lastScan = result;
            lastScanTime = DateTimeOffset.Now;
            ExportButton.IsEnabled = true;
            ProductsList.ItemsSource = result.Products;
            FilesList.ItemsSource = result.ApplicationFiles;
            WarningsList.ItemsSource = result.Warnings;
            if (result.Products.Count == 0)
            {
                EmptyText.Text = result.Warnings.Count > 0
                    ? "Записи Office не распознаны. Проверка могла быть неполной — смотрите сообщения ниже."
                    : "Записи установки Office не найдены в проверенных разделах реестра.";
                EmptyText.Visibility = Visibility.Visible;
            }
            if (result.ApplicationFiles.Count > 0)
                FilesHeading.Visibility = FilesNote.Visibility = Visibility.Visible;
            StatusText.Text = $"Проверка завершена: продуктов — {result.Products.Count}, файлов приложений — {result.ApplicationFiles.Count}, замечаний — {result.Warnings.Count}.";
        }
        catch (Exception exception)
        {
            StatusText.Text = "Не удалось завершить диагностику. Изменения системы не выполнялись.";
            WarningsList.ItemsSource = new[] { exception.Message };
        }
        finally
        {
            scanning = false;
            SetBusy(false);
            ScanButton.IsEnabled = true;
            ScanButton.Content = "Проверить Office";
            ScanProgress.Visibility = Visibility.Collapsed;
        }
    }

    private async void Repair_Click(object sender, RoutedEventArgs e) => await RepairAsync(sender, reinstall: false);
    private async void Reinstall_Click(object sender, RoutedEventArgs e) => await RepairAsync(sender, reinstall: true);

    private async Task RepairAsync(object sender, bool reinstall)
    {
        if (busy || sender is not Button { Tag: OfficeProduct product }) return;
        var direct = RepairPolicy.CanRepairMsiDirectly(product);
        var title = reinstall ? "Переустановка текущего Office" : "Восстановление Office";
        var explanation = direct
            ? (reinstall ? "Windows Installer повторно установит все файлы выбранного MSI-продукта." : "Windows Installer восстановит отсутствующие или устаревшие файлы выбранного продукта.")
                + " Потребуются права администратора; могут понадобиться исходные установочные файлы. Автоматическая перезагрузка отключена."
            : "Будет открыт список установленных программ Windows. Выберите нужный Office → «Изменить». "
                + (reinstall ? "Для Click-to-Run выберите онлайн-восстановление: Microsoft переустановит пакет. Точная сборка может измениться; может потребоваться повторная активация. " : "Для Click-to-Run доступны быстрое и онлайн-восстановление. ")
                + "Для других установок следуйте штатным указаниям Microsoft. Приложение не отслеживает результат в этом окне.";
        if (MessageBox.Show(this, product.Name + "\n\n" + explanation
            + "\n\nПеред восстановлением или переустановкой сохраните документы и закройте приложения Office.", title,
            MessageBoxButton.OKCancel, MessageBoxImage.Warning, MessageBoxResult.Cancel) != MessageBoxResult.OK) return;
        try
        {
            SetBusy(true);
            OperationStatus.Text = "Подготавливаем восстановление выбранного Office…";
            var result = await new OfficeRepairService().RepairAsync(product, reinstall);
            OperationStatus.Text = result;
            StatusText.Text = "После завершения операции повторите проверку Office.";
            lastScan = null;
            ExportButton.IsEnabled = false;
        }
        catch (OperationCanceledException exception) { OperationStatus.Text = exception.Message; }
        catch (Exception exception)
        {
            OperationStatus.Text = "Операция не завершена: " + exception.Message;
            MessageBox.Show(this, exception.Message, title, MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { SetBusy(false); }
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (lastScan is null || busy) return;
        var dialog = new SaveFileDialog { Title = "Сохранить отчёт диагностики", Filter = "Отчёт JSON|*.json",
            FileName = "Office-diagnostics-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".json" };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            options.Converters.Add(new JsonStringEnumConverter());
            File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(new { CheckedAt = lastScanTime, Result = lastScan }, options));
            OperationStatus.Text = "Отчёт сохранён: " + dialog.FileName;
        }
        catch (Exception exception) { MessageBox.Show(this, exception.Message, "Сохранение отчёта", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    private void SetBusy(bool value)
    {
        busy = value;
        MainTabs.IsEnabled = !value;
        OperationProgress.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (!busy) return;
        e.Cancel = true;
        MessageBox.Show(this, "Дождитесь завершения текущей операции. Если открыто окно Microsoft, отменить установку или восстановление можно в нём.",
            "Выполняется операция", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
