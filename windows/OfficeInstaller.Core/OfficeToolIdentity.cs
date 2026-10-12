namespace OfficeInstaller.Core;

public static class OfficeToolIdentity
{
    public static bool Matches(string? originalFilename, string? description, string? productName, bool hasDeploymentHelp)
    {
        // The official ODT setup.exe identifies itself as Bootstrapper.exe in its
        // version resource. A name on disk alone is not enough to identify the tool.
        return (string.Equals(originalFilename, "Bootstrapper.exe", StringComparison.OrdinalIgnoreCase)
                || string.Equals(originalFilename, "setup.exe", StringComparison.OrdinalIgnoreCase))
            && (description ?? "").Contains("Office", StringComparison.OrdinalIgnoreCase)
            && string.Equals(productName, "Microsoft Office", StringComparison.OrdinalIgnoreCase)
            && hasDeploymentHelp;
    }
}
