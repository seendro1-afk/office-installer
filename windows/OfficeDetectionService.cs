using Microsoft.Win32;

public sealed class OfficeDetectionInfo
{
    public bool ClickToRunDetected { get; set; }

    public string ProductIds { get; set; } = "Не определено";

    public string Edition { get; set; } = "Не определена";

    public string Architecture { get; set; } = "Не определена";

    public string Version { get; set; } = "Не определена";

    public string Channel { get; set; } = "Не определён";

    public string UpdateUrl { get; set; } = "Не определён";

    public bool RequiresCompatibilityCheck { get; set; }
}

public static class OfficeDetectionService
{
    private const string ConfigurationPath =
        @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";

    public static OfficeDetectionInfo Detect()
    {
        var info = new OfficeDetectionInfo();

        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(
                ConfigurationPath
            );

            if (key == null)
            {
                return info;
            }

            info.ClickToRunDetected = true;

            string productIds = ReadString(
                key,
                "ProductReleaseIds"
            );

            string architecture = ReadString(
                key,
                "Platform"
            );

            string version = ReadString(
                key,
                "VersionToReport"
            );

            string updateUrl = ReadString(
                key,
                "UpdateChannel"
            );

            string channel = ReadString(
                key,
                "CDNBaseUrl"
            );

            info.ProductIds = productIds;
            info.Architecture = architecture;
            info.Version = version;
            info.Channel = channel;
            info.UpdateUrl = updateUrl;

            info.Edition = DetectEdition(productIds);

            info.RequiresCompatibilityCheck =
                RequiresCompatibilityCheck(productIds);

            return info;
        }
        catch
        {
            return info;
        }
    }

    private static string ReadString(
        RegistryKey key,
        string name
    )
    {
        object? value = key.GetValue(name);

        return value?.ToString() ?? "Не определено";
    }

    private static string DetectEdition(string productIds)
    {
        if (productIds.Contains(
            "ProPlus2024Volume",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Office LTSC Professional Plus 2024";
        }

        if (productIds.Contains(
            "Standard2024Volume",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Office LTSC Standard 2024";
        }

        if (productIds.Contains(
            "ProPlus2021Volume",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Office LTSC Professional Plus 2021";
        }

        if (productIds.Contains(
            "Standard2021Volume",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Office LTSC Standard 2021";
        }

        if (productIds.Contains(
            "O365",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Microsoft 365 Apps";
        }

        if (productIds.Contains(
            "2019",
            StringComparison.OrdinalIgnoreCase))
        {
            return "Office 2019";
        }

        return "Другая или неизвестная редакция Office";
    }

    private static bool RequiresCompatibilityCheck(
        string productIds
    )
    {
        if (string.IsNullOrWhiteSpace(productIds) ||
            productIds == "Не определено")
        {
            return true;
        }

        string[] ids = productIds.Split(
            ',',
            ';',
            StringSplitOptions.RemoveEmptyEntries |
            StringSplitOptions.TrimEntries
        );

        return ids.Any(id =>
            !id.Equals(
                "ProPlus2024Volume",
                StringComparison.OrdinalIgnoreCase
            )
        );
    }
}