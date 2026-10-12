using System.Xml.Linq;
using System.Collections.Frozen;

namespace OfficeInstaller.Core;

public sealed record OfficeEdition(string Id, string Name, string Release, string Channel, IReadOnlyList<string> Applications)
{
    public override string ToString() => Name;
}

public sealed record OfficeLanguage(string Id, string Name)
{
    public override string ToString() => Name;
}

public enum OfficeArchitecture { X86 = 32, X64 = 64 }
public sealed record InstallationRequest(string ProductId, OfficeArchitecture Architecture, string Language,
    IReadOnlySet<string> Applications);

public sealed record PreparedDownload(string ProductId, OfficeArchitecture Architecture, string Language, string[] Applications)
{
    public InstallationRequest ToRequest()
    {
        if (Applications is null || ProductId is null || Language is null)
            throw new ArgumentException("Описание скачанного пакета неполное.");
        var request = new InstallationRequest(ProductId, Architecture, Language, Applications.ToFrozenSet(StringComparer.Ordinal));
        OfficeDeployment.Validate(request);
        return request;
    }
}

public static class OfficeCatalog
{
    public static IReadOnlyList<OfficeEdition> Editions { get; } = new[]
    {
        new OfficeEdition("ProPlus2024Volume", "Office LTSC 2024 Professional Plus", "2024", "PerpetualVL2024", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "OneDrive"]),
        new OfficeEdition("Standard2024Volume", "Office LTSC 2024 Standard", "2024", "PerpetualVL2024", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "OneDrive"]),
        new OfficeEdition("ProPlus2021Volume", "Office LTSC 2021 Professional Plus", "2021", "PerpetualVL2021", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "OneDrive"]),
        new OfficeEdition("Standard2021Volume", "Office LTSC 2021 Standard", "2021", "PerpetualVL2021", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive"]),
        new OfficeEdition("O365ProPlusRetail", "Microsoft 365 Apps для предприятий", "Microsoft 365", "Current", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "OneDrive"]),
        new OfficeEdition("O365BusinessRetail", "Microsoft 365 Apps для бизнеса", "Microsoft 365", "Current", ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "OneDrive"])
    };
    public static IReadOnlyList<OfficeLanguage> Languages { get; } = new[]
    {
        new OfficeLanguage("ru-ru", "Русский"), new OfficeLanguage("en-us", "English"),
        new OfficeLanguage("uk-ua", "Українська"), new OfficeLanguage("de-de", "Deutsch"),
        new OfficeLanguage("fr-fr", "Français")
    };
    public static string ApplicationName(string id) => id == "OneDrive" ? "OneDrive (дополнительно)" : id;
}

public static class OfficeDeployment
{
    // Every optional component is explicitly excluded unless selected. No removal,
    // activation, architecture migration or forced application shutdown is generated.
    private static readonly string[] Excludable =
        ["Access", "Excel", "Groove", "Lync", "OneDrive", "OneNote", "Outlook", "OutlookForWindows", "PowerPoint", "Publisher", "Teams", "Word"];

    public static OfficeEdition Validate(InstallationRequest request)
    {
        var edition = OfficeCatalog.Editions.SingleOrDefault(e => e.Id == request.ProductId)
            ?? throw new ArgumentException("Редакция Office не поддерживается.");
        if (!Enum.IsDefined(request.Architecture)) throw new ArgumentException("Выберите разрядность Office.");
        if (!OfficeCatalog.Languages.Any(l => l.Id == request.Language)) throw new ArgumentException("Язык Office не поддерживается.");
        if (request.Applications.Count == 0 || request.Applications.All(a => a == "OneDrive"))
            throw new ArgumentException("Выберите хотя бы одно приложение Office.");
        if (request.Applications.Any(a => !edition.Applications.Contains(a, StringComparer.Ordinal)))
            throw new ArgumentException("Выбран компонент, отсутствующий в этой редакции Office.");
        return edition;
    }

    public static string Configuration(InstallationRequest request, string sourcePath, bool download)
    {
        var edition = Validate(request);
        if (!System.Text.RegularExpressions.Regex.IsMatch(sourcePath, @"^[A-Za-z]:\\[^\r\n]*$"))
            throw new ArgumentException("Для пакетов нужен абсолютный путь на локальном диске Windows.");
        var product = new XElement("Product", new XAttribute("ID", edition.Id),
            new XElement("Language", new XAttribute("ID", request.Language)));
        foreach (var app in Excludable.Where(a => !request.Applications.Contains(a)))
            product.Add(new XElement("ExcludeApp", new XAttribute("ID", app)));
        var root = new XElement("Configuration",
            new XElement("Add", new XAttribute("SourcePath", sourcePath),
                new XAttribute("OfficeClientEdition", (int)request.Architecture),
                new XAttribute("Channel", edition.Channel),
                new XAttribute("AllowCdnFallback", "FALSE"), product));
        if (!download)
        {
            root.Add(new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "FALSE")));
            root.Add(new XElement("Property", new XAttribute("Name", "FORCEAPPSHUTDOWN"), new XAttribute("Value", "FALSE")));
            root.Add(new XElement("Property", new XAttribute("Name", "AUTOACTIVATE"), new XAttribute("Value", "0")));
        }
        return new XDocument(new XDeclaration("1.0", "utf-8", null), root).ToString();
    }

    public static void CheckNewInstallation(OfficeScanResult scan)
    {
        if (scan.Warnings.Count > 0)
            throw new InvalidOperationException("Диагностика неполная. Установка заблокирована; сначала проверьте замечания во вкладке восстановления.");
        if (scan.Products.Count > 0 || scan.ApplicationFiles.Count > 0)
            throw new InvalidOperationException("Обнаружен установленный Office или его файлы. Для текущей версии используйте восстановление. Новая установка не будет менять состав существующего Office или удалять его.");
    }
}

public static class RepairPolicy
{
    public static bool CanRepairMsiDirectly(OfficeProduct product) =>
        product.InstallType == OfficeInstallType.Msi && product.Scope == "Компьютер"
        && Guid.TryParseExact(product.ProductId, "B", out _);

    public static IReadOnlyList<string> MsiArguments(OfficeProduct product, bool reinstall = false)
    {
        if (!CanRepairMsiDirectly(product)) throw new ArgumentException("Нет проверенного MSI ProductCode для восстановления.");
        return [reinstall ? "/famus" : "/fomus", Guid.Parse(product.ProductId).ToString("B"), "/norestart"];
    }
}

public static class MicrosoftDownloadPolicy
{
    public static Uri? FindDeploymentTool(string html)
    {
        var decoded = System.Net.WebUtility.HtmlDecode(html).Replace("\\/", "/")
            .Replace("\\u002f", "/", StringComparison.OrdinalIgnoreCase);
        return System.Text.RegularExpressions.Regex.Matches(decoded,
                @"https://download\.microsoft\.com/[^\s""'<>\\]+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            .Select(match => Uri.TryCreate(match.Value, UriKind.Absolute, out var uri) ? uri : null)
            .FirstOrDefault(uri => uri is not null && IsDeploymentTool(uri));
    }

    public static bool IsDownloadPage(Uri uri) => IsHttps(uri)
        && (uri.Host.Equals("www.microsoft.com", StringComparison.OrdinalIgnoreCase)
            || uri.Host.Equals("microsoft.com", StringComparison.OrdinalIgnoreCase));

    public static bool IsDeploymentTool(Uri uri) => IsHttps(uri)
        && uri.Host.Equals("download.microsoft.com", StringComparison.OrdinalIgnoreCase)
        && System.Text.RegularExpressions.Regex.IsMatch(uri.AbsolutePath,
            @"/officedeploymenttool_[0-9]+-[0-9]+\.exe$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    private static bool IsHttps(Uri uri) => uri.IsAbsoluteUri && uri.Scheme == Uri.UriSchemeHttps
        && uri.UserInfo.Length == 0 && uri.IsDefaultPort;
}
