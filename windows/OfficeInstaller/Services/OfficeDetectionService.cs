using System.Diagnostics;
using System.IO;
using System.Security;
using Microsoft.Win32;
using OfficeInstaller.Core;

namespace OfficeInstaller.Services;

public sealed class OfficeDetectionService
{
    private const string ClickToRun = @"SOFTWARE\Microsoft\Office\ClickToRun";
    private const string Uninstall = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

    public OfficeScanResult Scan()
    {
        var warnings = new List<string>();
        var c2r = new List<ClickToRunSnapshot>();
        var uninstall = new List<RegistrySnapshot>();
        var roots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Registry32 };

        foreach (var view in views)
        {
            ReadSafely(() =>
            {
                using var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                var inventory = Read(machine, ClickToRun + @"\Inventory\Office\16.0", view,
                    ["OfficePackageVersion", "OfficeProductReleaseIds"]);
                var config = Read(machine, ClickToRun + @"\Configuration", view,
                    ["ProductReleaseIds", "Platform", "ClientCulture", "InstallationPath"]);
                c2r.Add(new ClickToRunSnapshot(inventory, config));
                var installPath = config.Get("InstallationPath");
                if (Path.IsPathFullyQualified(installPath) && !installPath.StartsWith(@"\\")) roots.Add(installPath);
                else if (installPath.Length > 0)
                    warnings.Add("Нестандартный или сетевой путь Office не проверялся: " + installPath);
            }, "Click-to-Run / " + view, warnings);

            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                ReadSafely(() =>
                {
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using var parent = baseKey.OpenSubKey(Uninstall, writable: false);
                    if (parent is null) return;
                    foreach (var keyName in parent.GetSubKeyNames())
                    {
                        ReadSafely(() =>
                        {
                            var entry = Read(baseKey, Uninstall + "\\" + keyName, view,
                                ["DisplayName", "DisplayVersion", "Publisher", "WindowsInstaller", "SystemComponent", "UninstallString"]);
                            var values = new Dictionary<string, string>(entry.Values, StringComparer.OrdinalIgnoreCase)
                            {
                                ["RegistryKeyName"] = keyName,
                                ["RegistryScope"] = hive.ToString()
                            };
                            uninstall.Add(entry with { Values = values });
                        }, hive + " / " + view + " / " + keyName, warnings);
                    }
                }, hive + " / " + view, warnings);
            }
        }

        foreach (var folder in new[] { Environment.SpecialFolder.ProgramFiles, Environment.SpecialFolder.ProgramFilesX86 })
        {
            var programFiles = Environment.GetFolderPath(folder);
            if (programFiles.Length > 0) roots.Add(Path.Combine(programFiles, "Microsoft Office"));
        }

        var files = FindApplications(roots, warnings);
        var products = OfficeInventory.Parse(c2r, uninstall);
        if (products.Any(p => p.InstallType == OfficeInstallType.ClickToRun && p.Build == "Не определено"))
            warnings.Add("В Click-to-Run Inventory нет данных о сборке. Запись конфигурации сама по себе не подтверждает завершённую установку.");
        if (products.Count == 0 && files.Count > 0)
            warnings.Add("Найдены файлы приложений без распознанных записей продукта. Возможны остатки установки; редакция и способ восстановления не определены.");
        return new OfficeScanResult(products, files, warnings);
    }

    private static RegistrySnapshot Read(RegistryKey root, string path, RegistryView view, string[] names)
    {
        using var key = root.OpenSubKey(path, writable: false);
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (key is not null)
            foreach (var name in names)
            {
                var value = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (value is not null) values[name] = value is string[] items ? string.Join(",", items) : value.ToString() ?? "";
            }
        return new RegistrySnapshot(root.Name + "\\" + path + " [" + view + "]", values);
    }

    private static List<ApplicationFile> FindApplications(IEnumerable<string> roots, List<string> warnings)
    {
        var files = new List<ApplicationFile>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var applications = new (string Name, string File)[]
        {
            ("Word", "WINWORD.EXE"), ("Excel", "EXCEL.EXE"), ("PowerPoint", "POWERPNT.EXE"),
            ("Outlook", "OUTLOOK.EXE"), ("OneNote", "ONENOTE.EXE"), ("Access", "MSACCESS.EXE"),
            ("Publisher", "MSPUB.EXE"), ("Visio", "VISIO.EXE"), ("Project", "WINPROJ.EXE")
        };
        foreach (var root in roots)
        foreach (var relative in new[] { "", "root/Office16", "Office16", "Office15", "Office14", "Office12", "Office11" })
        foreach (var app in applications)
        {
            ReadSafely(() =>
            {
                // Only fixed candidate locations are read; no disk-wide traversal or executable launch.
                var path = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar), app.File);
                if (!File.Exists(path) || !seen.Add(path)) return;
                var version = FileVersionInfo.GetVersionInfo(path).FileVersion ?? "Не определено";
                files.Add(new ApplicationFile(app.Name, path, version));
            }, root + " / " + app.Name, warnings);
        }
        return files.OrderBy(f => f.Name).ThenBy(f => f.Path).ToList();
    }

    private static void ReadSafely(Action action, string source, List<string> warnings)
    {
        try { action(); }
        catch (Exception exception) when (exception is UnauthorizedAccessException or SecurityException or IOException or ArgumentException)
        {
            warnings.Add("Не удалось прочитать " + source + ": " + exception.Message);
        }
    }
}
