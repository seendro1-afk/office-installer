using System.Diagnostics;
using System.Text;

public static class OfficeDiagnostics
{
    public static string GenerateReport(
        OfficeDetectionInfo? office,
        string selectedArchitecture,
        string? blockReason,
        Dictionary<string, string> applications)
    {
        var report = new StringBuilder();

        report.AppendLine("==========================================");
        report.AppendLine(" OFFICE INSTALLER — DIAGNOSTICS REPORT");
        report.AppendLine("==========================================");
        report.AppendLine();

        report.AppendLine("ОБЩАЯ ИНФОРМАЦИЯ");
        report.AppendLine("------------------------------------------");

        report.AppendLine(
            "Дата диагностики: " +
            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

        report.AppendLine(
            "Операционная система: " +
            Environment.OSVersion.VersionString);

        report.AppendLine(
            "64-битная Windows: " +
            (Environment.Is64BitOperatingSystem ? "Да" : "Нет"));

        report.AppendLine();

        report.AppendLine("УСТАНОВЛЕННЫЙ MICROSOFT OFFICE");
        report.AppendLine("------------------------------------------");

        if (office == null)
        {
            report.AppendLine("Не удалось получить сведения об Office.");
        }
        else
        {
            report.AppendLine(
                "Click-to-Run обнаружен: " +
                (office.ClickToRunDetected ? "Да" : "Нет"));

            report.AppendLine("Редакция: " + office.Edition);
            report.AppendLine("Продукты: " + office.ProductIds);
            report.AppendLine("Архитектура: " + office.Architecture);
            report.AppendLine("Версия: " + office.Version);
            report.AppendLine("Канал обновлений: " + office.Channel);
        }

        report.AppendLine();

        report.AppendLine("ПРИЛОЖЕНИЯ OFFICE");
        report.AppendLine("------------------------------------------");

        foreach (var app in applications)
        {
            report.AppendLine(
                app.Key + ": " + app.Value);
        }

        report.AppendLine();

        report.AppendLine("ПРОВЕРКА СОВМЕСТИМОСТИ");
        report.AppendLine("------------------------------------------");

        report.AppendLine(
            "Целевая редакция: Office LTSC Professional Plus 2024");

        report.AppendLine(
            "Выбранная архитектура: " + selectedArchitecture);

        if (blockReason == null)
        {
            report.AppendLine(
                "Результат: предварительная проверка пройдена.");

            report.AppendLine(
                "Обнаруженных текущими проверками конфликтов нет.");
        }
        else
        {
            report.AppendLine("Результат: УСТАНОВКА ЗАБЛОКИРОВАНА");
            report.AppendLine("Причина: " + blockReason);
        }

        report.AppendLine();

        report.AppendLine("ПРИМЕЧАНИЕ");
        report.AppendLine("------------------------------------------");

        report.AppendLine(
            "Отчёт сформирован без запуска Office Deployment Tool.");

        report.AppendLine(
            "Диагностика не изменяет установленные продукты Office.");

        report.AppendLine(
            "Предварительная проверка не гарантирует успешную установку.");

        report.AppendLine();

        report.AppendLine("Конец отчёта.");

        return report.ToString();
    }
}
