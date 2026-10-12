using System.ComponentModel;
using System.Diagnostics;
using System.IO;

namespace OfficeInstaller.Services;

public static class WindowsProcess
{
    public static async Task<int> RunAsync(string executable, IEnumerable<string> arguments, string workingDirectory, bool elevated)
    {
        if (!Path.IsPathFullyQualified(executable) || !File.Exists(executable))
            throw new FileNotFoundException("Программа для выполнения операции не найдена.", executable);
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = true, WorkingDirectory = workingDirectory
        };
        if (elevated) start.Verb = "runas";
        foreach (var argument in arguments) start.ArgumentList.Add(argument);
        try
        {
            using var process = Process.Start(start) ?? throw new InvalidOperationException("Не удалось запустить программу.");
            await process.WaitForExitAsync();
            return process.ExitCode;
        }
        catch (Win32Exception exception) when (exception.NativeErrorCode == 1223)
        {
            throw new OperationCanceledException("Запрос прав администратора отменён пользователем.", exception);
        }
    }

    public static string SystemExecutable(string name) => Path.Combine(Environment.SystemDirectory, name);
}
