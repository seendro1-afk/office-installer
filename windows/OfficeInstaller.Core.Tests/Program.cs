using OfficeInstaller.Core;
using System.Xml.Linq;

RegistrySnapshot Snapshot(params (string Key, string Value)[] values) =>
    new("fixture", values.ToDictionary(v => v.Key, v => v.Value, StringComparer.OrdinalIgnoreCase));
ClickToRunSnapshot C2R(string ids, string build = "16.0.17932.20162") => new(
    Snapshot(("OfficeProductReleaseIds", ids), ("OfficePackageVersion", build)),
    Snapshot(("ProductReleaseIds", "ProPlus2021Volume"), ("Platform", "x64"), ("ClientCulture", "ru-ru")));
RegistrySnapshot Msi(string name, string id, string build) => Snapshot(
    ("DisplayName", name), ("Publisher", "Microsoft Corporation"),
    ("RegistryKeyName", id), ("WindowsInstaller", "1"), ("DisplayVersion", build));

var passed = 0;
void Check(bool condition, string name)
{
    if (!condition) throw new Exception("FAIL: " + name);
    Console.WriteLine("PASS: " + name);
    passed++;
}
var products = OfficeInventory.Parse([C2R("ProPlus2024Volume,VisioPro2024Volume")], []);
Check(products.Count == 2 && products.All(p => p.Release == "2024"), "multiple C2R products; inventory wins over config");
Check(products.All(p => p.Build == "16.0.17932.20162" && p.Architecture == "64 бита" && p.Language == "ru-ru"), "build, architecture and language");
Check(OfficeInventory.Parse([C2R("ProPlus2021Volume")], []).Single().Release == "2021", "Office 2021 shares 16.0 with 2024");
Check(OfficeInventory.Parse([C2R("O365ProPlusRetail")], []).Single().Release == "Microsoft 365", "Microsoft 365 product");
Check(OfficeInventory.Release("ProPlusRetail") == "Не определено", "unknown release is not guessed");
var fallback = new ClickToRunSnapshot(Snapshot(), Snapshot(
    ("ProductReleaseIds", "ProPlus2021Volume"), ("VersionToReport", "16.0.99999.1")));
var fallbackProduct = OfficeInventory.Parse([fallback], []).Single();
Check(fallbackProduct.Build == "Не определено" && fallbackProduct.Source.Contains("резервный"), "telemetry build not used; fallback disclosed");
Check(OfficeInventory.Parse([C2R("ProPlus2024Volume"), C2R("proplus2024volume")], []).Count == 1, "duplicate registry views");
var mixed = OfficeInventory.Parse([C2R("ProPlus2024Volume")], [Msi("Microsoft Office Professional Plus 2010", "{old}", "14.0.1")]);
Check(mixed.Count == 2 && mixed.Any(p => p.InstallType == OfficeInstallType.Msi && p.Release == "2010"), "legacy MSI alongside C2R");
var auxiliary = new[] { "Microsoft Office Update 2021", "Microsoft Access Database Engine 2016", "Microsoft Office Proofing 2016", "Microsoft Office Runtime 2010", "Microsoft Office Language Pack 2013" };
Check(OfficeInventory.Parse([], auxiliary.Select(n => Msi(n, n, "16.0"))).Count == 0, "exclude auxiliary components");
Check(OfficeInventory.Parse([], [Snapshot(("DisplayName", "Office 2024"), ("Publisher", "Other vendor"))]).Count == 0, "exclude other publishers");
var duplicate = Snapshot(("DisplayName", "Microsoft Office Professional Plus 2024"),
    ("Publisher", "Microsoft Corporation"), ("RegistryKeyName", "Office16.ProPlus2024Volume-ru-ru"),
    ("UninstallString", "OfficeClickToRun.exe"));
Check(OfficeInventory.Parse([C2R("ProPlus2024Volume")], [duplicate]).Count == 1, "C2R uninstall entry not counted twice");
Check(OfficeInventory.Parse([], [duplicate]).Single().InstallType == OfficeInstallType.ClickToRun, "C2R uninstall fallback without inventory");
var unknown = OfficeInventory.Parse([], [Snapshot(("DisplayName", "Microsoft Office 2021"), ("Publisher", "Microsoft Corporation"))]).Single();
Check(unknown.InstallType == OfficeInstallType.Unknown && unknown.Architecture == "Не определено", "do not infer MSI or architecture from registry location");
Check(OfficeInventory.Parse([], []).Count == 0, "empty computer");
var orphan = duplicate with { Values = Snapshot(("DisplayName", "Microsoft Office 2021"),
    ("Publisher", "Microsoft Corporation"), ("RegistryKeyName", "Office16.ProPlus2021Volume-ru-ru"),
    ("UninstallString", "OfficeClickToRun.exe")).Values };
Check(OfficeInventory.Parse([C2R("ProPlus2024Volume")], [orphan]).Count == 2, "do not hide an unrelated C2R installation");
var machineMsi = Msi("Microsoft Office Professional Plus 2010", "{00000000-0000-0000-0000-000000000001}", "14.0");
var userMsi = machineMsi with { Values = Snapshot(("DisplayName", "Microsoft Office Professional Plus 2010"),
    ("Publisher", "Microsoft Corporation"), ("RegistryKeyName", "{00000000-0000-0000-0000-000000000001}"),
    ("DisplayVersion", "14.0"), ("WindowsInstaller", "1"), ("RegistryScope", "CurrentUser")).Values };
Check(OfficeInventory.Parse([], [machineMsi, userMsi]).Count == 2, "preserve user and machine MSI scopes");

void Reject(Action action, string name)
{
    var rejected = false;
    try { action(); }
    catch (ArgumentException) { rejected = true; }
    catch (InvalidOperationException) { rejected = true; }
    Check(rejected, name);
}
var request = new InstallationRequest("ProPlus2024Volume", OfficeArchitecture.X64, "ru-ru",
    new HashSet<string> { "Word", "Excel" });
var xml = System.Xml.Linq.XDocument.Parse(OfficeDeployment.Configuration(request, @"C:\Office & Tools\Packages", false));
Check(xml.Root!.Element("Add")!.Attribute("SourcePath")!.Value == @"C:\Office & Tools\Packages", "XML escapes a path with ampersands");
Check(xml.Descendants("Product").Single().Attribute("ID")!.Value == "ProPlus2024Volume"
    && xml.Descendants("Add").Single().Attribute("Channel")!.Value == "PerpetualVL2024", "2024 product and channel match");
var excluded = xml.Descendants("ExcludeApp").Select(e => e.Attribute("ID")!.Value).ToArray();
Check(!excluded.Contains("Word") && !excluded.Contains("Excel") && excluded.Contains("Outlook")
    && excluded.Contains("OutlookForWindows") && excluded.Contains("Teams") && excluded.Contains("OneDrive"), "only requested components enabled");
Check(!xml.Descendants("Remove").Any() && !xml.Descendants("RemoveMSI").Any()
    && !xml.Descendants().Attributes("MigrateArch").Any(), "no automatic uninstall or architecture migration");
Check(xml.Descendants("Display").Single().Attribute("AcceptEULA")!.Value == "FALSE"
    && xml.Descendants("Property").Single(p => p.Attribute("Name")!.Value == "FORCEAPPSHUTDOWN").Attribute("Value")!.Value == "FALSE",
    "visible EULA and no forced app shutdown");
Check(!System.Xml.Linq.XDocument.Parse(OfficeDeployment.Configuration(request, @"C:\Packages", true)).Descendants("Display").Any(), "download config separate from install config");
foreach (var edition in OfficeCatalog.Editions)
{
    var item = request with { ProductId = edition.Id, Architecture = OfficeArchitecture.X86 };
    var config = System.Xml.Linq.XDocument.Parse(OfficeDeployment.Configuration(item, @"C:\Packages", false));
    Check(config.Descendants("Add").Single().Attribute("Channel")!.Value == edition.Channel
        && config.Descendants("Add").Single().Attribute("OfficeClientEdition")!.Value == "32", "catalog configuration: " + edition.Id);
}
Reject(() => OfficeDeployment.Validate(request with { Applications = new HashSet<string>() }), "empty selection rejected");
Reject(() => OfficeDeployment.Validate(request with { Applications = new HashSet<string> { "OneDrive" } }), "OneDrive alone rejected");
Reject(() => OfficeDeployment.Validate(request with { ProductId = "Standard2024Volume", Applications = new HashSet<string> { "Access" } }), "Access unavailable in Standard");
Reject(() => OfficeDeployment.Validate(request with { Applications = new HashSet<string> { "Word", "Injected" } }), "unknown app rejected");
Reject(() => OfficeDeployment.Validate(request with { ProductId = "UnknownEdition" }), "unknown product rejected");
Reject(() => OfficeDeployment.Validate(request with { Language = "ru-ru\" injected=\"true" }), "unknown language rejected");
Reject(() => OfficeDeployment.Validate(request with { Architecture = (OfficeArchitecture)128 }), "invalid architecture rejected");
Reject(() => OfficeDeployment.Configuration(request, @"\\server\share", false), "UNC package path rejected");
Reject(() => OfficeDeployment.Configuration(request, "C:\\Packages\nInjected", false), "newline in path rejected");
var emptyScan = new OfficeScanResult([], [], []);
OfficeDeployment.CheckNewInstallation(emptyScan);
Check(true, "new install allowed on empty inventory");
Reject(() => OfficeDeployment.CheckNewInstallation(emptyScan with { Products = mixed }), "existing Office blocks install");
Reject(() => OfficeDeployment.CheckNewInstallation(emptyScan with { ApplicationFiles = [new("Word", "file", "16.0")] }), "orphan app files block install");
Reject(() => OfficeDeployment.CheckNewInstallation(emptyScan with { Warnings = ["Access denied"] }), "partial scan blocks install");
var msiProduct = OfficeInventory.Parse([], [machineMsi]).Single();
Check(RepairPolicy.CanRepairMsiDirectly(msiProduct) && RepairPolicy.MsiArguments(msiProduct).SequenceEqual(
    new[] { "/fomus", "{00000000-0000-0000-0000-000000000001}", "/norestart" }), "MSI repair uses validated GUID and no restart");
Check(RepairPolicy.MsiArguments(msiProduct, reinstall: true)[0] == "/famus", "MSI reinstall replaces all files without uninstall");
Reject(() => RepairPolicy.MsiArguments(msiProduct with { ProductId = "{bad} /x other" }), "MSI argument injection rejected");
Check(!RepairPolicy.CanRepairMsiDirectly(msiProduct with { Scope = "Текущий пользователь" }), "per-user MSI uses native UI");
Check(!RepairPolicy.CanRepairMsiDirectly(products[0]), "C2R never routed to msiexec repair");
Check(MicrosoftDownloadPolicy.IsDeploymentTool(new Uri("https://download.microsoft.com/download/a/officedeploymenttool_20326-20112.exe")), "official ODT URL accepted");
foreach (var bad in new[] { "http://download.microsoft.com/officedeploymenttool_1-2.exe", "https://download.microsoft.com.evil.test/officedeploymenttool_1-2.exe",
    "https://evil.test/officedeploymenttool_1-2.exe", "https://user@download.microsoft.com/officedeploymenttool_1-2.exe", "https://download.microsoft.com:444/officedeploymenttool_1-2.exe",
    "https://download.microsoft.com/other.exe", "https://download.microsoft.com/officedeploymenttool_1-2.exe/other" })
    Check(!MicrosoftDownloadPolicy.IsDeploymentTool(new Uri(bad)), "reject unsafe ODT URL: " + bad);
Check(!MicrosoftDownloadPolicy.IsDeploymentTool(new Uri("relative", UriKind.Relative)), "relative download URL rejected");
Check(MicrosoftDownloadPolicy.IsDownloadPage(new Uri("https://www.microsoft.com/en-us/download/details.aspx?id=49117"))
    && !MicrosoftDownloadPolicy.IsDownloadPage(new Uri("https://www.microsoft.com.evil.test/")), "page host allowlist");
Check(MicrosoftDownloadPolicy.FindDeploymentTool("<a href=\"https://download.microsoft.com/download/a/officedeploymenttool_1-2.exe\">Download</a>") is not null, "ODT link extracted from HTML");
Check(MicrosoftDownloadPolicy.FindDeploymentTool("{\"url\":\"https:\\/\\/download.microsoft.com\\/download\\/a\\/officedeploymenttool_1-2.exe\"}") is not null, "ODT link extracted from escaped JSON");
Check(MicrosoftDownloadPolicy.FindDeploymentTool("<a href=\"https://download.microsoft.com/download/a/officedeploymenttool_1-2.exe/other\">") is null, "ODT link not truncated to a trusted-looking prefix");
Check(MicrosoftDownloadPolicy.FindDeploymentTool("<a href=\"https://download.microsoft.com.evil.test/officedeploymenttool_1-2.exe\">") is null, "no fake ODT host extracted");
var manifest = new PreparedDownload(request.ProductId, request.Architecture, request.Language, request.Applications.ToArray());
var roundTrip = System.Text.Json.JsonSerializer.Deserialize<PreparedDownload>(System.Text.Json.JsonSerializer.Serialize(manifest))!.ToRequest();
Check(roundTrip.ProductId == request.ProductId && roundTrip.Applications.SetEquals(request.Applications), "saved package selection round trip");
Reject(() => (manifest with { ProductId = "Injected" }).ToRequest(), "tampered saved edition rejected");
Reject(() => (manifest with { Applications = null! }).ToRequest(), "incomplete saved selection rejected");
Check(OfficeToolIdentity.Matches("Bootstrapper.exe", "Microsoft 365 and Office", "Microsoft Office", true), "official ODT Bootstrapper.exe version metadata accepted");
Check(!OfficeToolIdentity.Matches("Bootstrapper.exe", "Microsoft 365 and Office", "Microsoft Office", false), "Office bootstrapper without deployment help rejected");
Check(!OfficeToolIdentity.Matches("other.exe", "Microsoft Office", "Microsoft Office", true), "unrelated Microsoft executable identity rejected");
Console.WriteLine($"All {passed} checks passed.");
if (args is ["--odt-page", var pagePath])
{
    var link = MicrosoftDownloadPolicy.FindDeploymentTool(File.ReadAllText(pagePath));
    Check(link is not null, "ODT URL from downloaded Microsoft page");
    Console.WriteLine("ODT source: " + link);
}
