using System.Diagnostics;
using System.Text;

public static class InstallerLogger
{
    private static readonly object Sync = new();

    public static string LogDirectory => Path.Combine(
        Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData),
        "OfficeInstaller",
        "Logs");

    public static string LogPath => Path.Combine(
        LogDirectory,
        "office-installer-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");

    public static void Write(string message)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);

            string line =
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}" +
                Environment.NewLine;

            lock (Sync)
            {
                File.AppendAllText(
                    LogPath,
                    line,
                    Encoding.UTF8);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine("Ошибка записи журнала: " + ex.Message);
        }
    }

    public static void Error(string operation, Exception exception)
    {
        Write($"ОШИБКА [{operation}]: {exception}");
    }

    public static void OpenDirectory()
    {
        Directory.CreateDirectory(LogDirectory);

        Process.Start(new ProcessStartInfo
        {
            FileName = LogDirectory,
            UseShellExecute = true
        });
    }
}
