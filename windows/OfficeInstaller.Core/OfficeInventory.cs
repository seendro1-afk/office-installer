using System.Text.RegularExpressions;

namespace OfficeInstaller.Core;

public enum OfficeInstallType { ClickToRun, Msi, Unknown }

public sealed record OfficeProduct(
    string Name, string ProductId, string Release, string Build,
    string Architecture, string Language, OfficeInstallType InstallType, string Source)
{
    public string Scope { get; init; } = "Компьютер";
    public string InstallTypeLabel => InstallType switch
    {
        OfficeInstallType.ClickToRun => "Click-to-Run",
        OfficeInstallType.Msi => "Windows Installer (MSI)",
        _ => "Не определён"
    };

    public string RepairDescription => InstallType switch
    {
        OfficeInstallType.ClickToRun => "В окне установленных программ выберите этот продукт → «Изменить». Microsoft предложит быстрое или онлайн-восстановление всего пакета. Онлайн-восстановление может потребовать повторной активации.",
        OfficeInstallType.Msi => "Штатное восстановление MSI; могут потребоваться исходные установочные файлы.",
        _ => "Сначала необходимо уточнить тип установки. Автоматическое восстановление пока недоступно."
    };
}

public sealed record RegistrySnapshot(string Source, IReadOnlyDictionary<string, string> Values)
{
    public string Get(string name) => Values.TryGetValue(name, out var value) ? value.Trim() : "";
}

public sealed record ClickToRunSnapshot(RegistrySnapshot Inventory, RegistrySnapshot Configuration);
public sealed record ApplicationFile(string Name, string Path, string Version);
public sealed record OfficeScanResult(
    IReadOnlyList<OfficeProduct> Products,
    IReadOnlyList<ApplicationFile> ApplicationFiles,
    IReadOnlyList<string> Warnings);

public static class OfficeInventory
{
    private const string Unknown = "Не определено";
    private static readonly Regex OfficeName = new(
        @"\b(?:Office|Microsoft 365|Word|Excel|PowerPoint|Outlook|OneNote|Access|Publisher|Visio|Project)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex NonProduct = new(
        @"\b(?:Update|Hotfix|Runtime|Redistributable|Proofing|Language Pack|MUI|Click-to-Run Extensibility|Add-in|Addin|Viewer|Driver|Database Engine|Telemetry|Tools)\b|обновлен|средства проверки|языковой пакет",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static IReadOnlyList<OfficeProduct> Parse(
        IEnumerable<ClickToRunSnapshot> clickToRun, IEnumerable<RegistrySnapshot> uninstall)
    {
        var products = new List<OfficeProduct>();
        foreach (var snapshot in clickToRun)
        {
            var inventory = snapshot.Inventory;
            var config = snapshot.Configuration;
            var inventoryIds = inventory.Get("OfficeProductReleaseIds");
            var ids = inventoryIds.Length > 0 ? inventoryIds : config.Get("ProductReleaseIds");
            var build = inventory.Get("OfficePackageVersion");
            // Telemetry values are deliberately not treated as an installed build.
            foreach (var id in ids.Split([',', ';'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            {
                products.Add(new OfficeProduct(
                    ProductName(id), id, Release(id), Present(build),
                    Architecture(config.Get("Platform")), Present(config.Get("ClientCulture")),
                    OfficeInstallType.ClickToRun,
                    inventoryIds.Length > 0 ? inventory.Source : config.Source + " (резервный источник; состояние установки требует проверки)"));
            }
        }

        foreach (var entry in uninstall)
        {
            var name = entry.Get("DisplayName");
            if (!entry.Get("Publisher").StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase)
                || !OfficeName.IsMatch(name) || NonProduct.IsMatch(name)
                || entry.Get("SystemComponent") == "1") continue;

            var isMsi = entry.Get("WindowsInstaller") == "1";
            var isC2R = entry.Get("UninstallString").Contains("ClickToRun", StringComparison.OrdinalIgnoreCase);
            // Only suppress an uninstall record when its key identifies a known product.
            // An unrelated C2R entry may represent an orphaned/different installation.
            if (isC2R && products.Any(p => p.InstallType == OfficeInstallType.ClickToRun
                && entry.Get("RegistryKeyName").Split(['.', '-'], StringSplitOptions.RemoveEmptyEntries)
                    .Contains(p.ProductId, StringComparer.OrdinalIgnoreCase))) continue;
            products.Add(new OfficeProduct(name, entry.Get("RegistryKeyName"), Release(name),
                Present(entry.Get("DisplayVersion")), Unknown, Unknown,
                isMsi ? OfficeInstallType.Msi : isC2R ? OfficeInstallType.ClickToRun : OfficeInstallType.Unknown,
                entry.Source) { Scope = entry.Get("RegistryScope") == "CurrentUser" ? "Текущий пользователь" : "Компьютер" });
        }

        return products.DistinctBy(p => (p.ProductId.ToUpperInvariant(), p.Build, p.InstallType, p.Scope))
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public static string Release(string value)
    {
        foreach (var year in new[] { "2024", "2021", "2019", "2016", "2013", "2010", "2007", "2003" })
            if (value.Contains(year, StringComparison.OrdinalIgnoreCase)) return year;
        if (value.Contains("O365", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Microsoft 365", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Office 365", StringComparison.OrdinalIgnoreCase)) return "Microsoft 365";
        return Unknown;
    }

    private static string ProductName(string id)
    {
        var release = Release(id);
        var suffix = release == Unknown ? "" : " " + release;
        if (id.StartsWith("O365", StringComparison.OrdinalIgnoreCase)) return "Microsoft 365 — " + id;
        var names = new (string Prefix, string Name)[]
        {
            ("ProPlus", "Office Professional Plus"), ("Standard", "Office Standard"),
            ("HomeBusiness", "Office Home & Business"), ("HomeStudent", "Office Home & Student"),
            ("Professional", "Office Professional"), ("Word", "Microsoft Word"),
            ("Excel", "Microsoft Excel"), ("PowerPoint", "Microsoft PowerPoint"),
            ("Outlook", "Microsoft Outlook"), ("OneNote", "Microsoft OneNote"),
            ("Access", "Microsoft Access"), ("Publisher", "Microsoft Publisher"),
            ("Visio", "Microsoft Visio"), ("Project", "Microsoft Project")
        };
        foreach (var item in names)
            if (id.StartsWith(item.Prefix, StringComparison.OrdinalIgnoreCase)) return item.Name + suffix;
        return "Office — " + id;
    }

    private static string Architecture(string value) => value.ToLowerInvariant() switch
    {
        "x86" => "32 бит", "x64" => "64 бита", "arm64" => "ARM64", _ => Unknown
    };
    private static string Present(string value) => value.Length > 0 ? value : Unknown;
}
