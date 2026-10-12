using System.Diagnostics;
using OfficeInstaller.Core;

namespace OfficeInstaller.Services;

public sealed class OfficeRepairService
{
    public async Task<string> RepairAsync(OfficeProduct product, bool reinstall = false)
    {
        if (!RepairPolicy.CanRepairMsiDirectly(product))
        {
            // C2R repair command-line switches are not a stable public API.
            // Use Microsoft's supported Change / Repair UI for every installed release.
            var start = new ProcessStartInfo(WindowsProcess.SystemExecutable("control.exe")) { UseShellExecute = true };
            start.ArgumentList.Add("appwiz.cpl");
            using var process = Process.Start(start) ?? throw new InvalidOperationException("Не удалось открыть список установленных программ.");
            return "Открыт список установленных программ. Выберите «" + product.Name
                + "» → «Изменить» → восстановление. Для полной переустановки Click-to-Run выберите онлайн-восстановление. Результат этой операции приложение не отслеживает; после завершения повторите диагностику.";
        }

        var current = await Task.Run(new OfficeDetectionService().Scan);
        if (current.Warnings.Count > 0 || !current.Products.Any(p => p.InstallType == OfficeInstallType.Msi
            && p.ProductId.Equals(product.ProductId, StringComparison.OrdinalIgnoreCase) && p.Scope == product.Scope))
            throw new InvalidOperationException("Не удалось повторно подтвердить MSI-установку. Повторите диагностику перед восстановлением.");
        var code = await WindowsProcess.RunAsync(WindowsProcess.SystemExecutable("msiexec.exe"),
            RepairPolicy.MsiArguments(product, reinstall), Environment.SystemDirectory, elevated: true);
        return code switch
        {
            0 => "Windows Installer завершил восстановление. Повторите диагностику и проверьте запуск приложений Office.",
            3010 => "Windows Installer завершил восстановление; требуется перезагрузка. Программа не перезагружает компьютер.",
            1602 => "Восстановление отменено пользователем.",
            1618 => "Уже выполняется другая операция Windows Installer. Дождитесь её завершения.",
            1612 => "Нет исходных установочных файлов MSI. Для восстановления укажите их в окне Windows Installer.",
            1605 => "Windows Installer не нашёл этот продукт. Повторите диагностику.",
            _ => $"Windows Installer завершился с кодом {code}. Восстановление не подтверждено."
        };
    }
}
