using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeInstaller.Core;

namespace OfficeInstaller.Services;

public static class MicrosoftSignatureVerifier
{
    public static void Verify(string path, bool deploymentTool = false)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("Исполняемый файл не найден.", path);
        var name = Marshal.StringToCoTaskMemUni(path);
        var filePointer = Marshal.AllocHGlobal(Marshal.SizeOf<TrustFile>());
        var file = new TrustFile { Size = (uint)Marshal.SizeOf<TrustFile>(), Path = name };
        Marshal.StructureToPtr(file, filePointer, false);
        var data = new TrustData
        {
            Size = (uint)Marshal.SizeOf<TrustData>(), UiChoice = 2,
            RevocationChecks = 1, UnionChoice = 1, File = filePointer,
            StateAction = 1, ProviderFlags = 0x80 // Revocation check for the chain excluding root.
        };
        var action = new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");
        try
        {
            var status = WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            if (status != 0)
                throw new InvalidOperationException($"Windows не подтвердила подпись файла (0x{status:X8}). Проверьте интернет, дату компьютера и источник файла. Проверка не будет пропущена.");
            using var signed = X509Certificate.CreateFromSignedFile(path);
            using var certificate = new X509Certificate2(signed);
            if (!certificate.GetNameInfo(X509NameType.SimpleName, false).Equals("Microsoft Corporation", StringComparison.Ordinal))
                throw new InvalidOperationException("Файл подписан другим издателем; требуется Microsoft Corporation.");
            if (deploymentTool)
            {
                var version = FileVersionInfo.GetVersionInfo(path);
                if (new FileInfo(path).Length > 64 * 1024 * 1024)
                    throw new InvalidOperationException("Размер setup.exe превышает ожидаемый для ODT. Запуск отклонён.");
                var bytes = File.ReadAllBytes(path);
                var hasHelp = bytes.AsSpan().IndexOf(Encoding.Unicode.GetBytes("Office Deployment Tool")) >= 0;
                if (!OfficeToolIdentity.Matches(version.OriginalFilename, version.FileDescription, version.ProductName, hasHelp))
                    throw new InvalidOperationException("Выберите setup.exe из Office Deployment Tool, а не другой установщик Microsoft.");
            }
        }
        finally
        {
            data.StateAction = 2;
            WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            Marshal.FreeHGlobal(filePointer);
            Marshal.FreeCoTaskMem(name);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct TrustFile
    {
        public uint Size;
        public IntPtr Path, FileHandle, KnownSubject;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct TrustData
    {
        public uint Size;
        public IntPtr PolicyCallback, SipClient;
        public uint UiChoice, RevocationChecks, UnionChoice;
        public IntPtr File;
        public uint StateAction;
        public IntPtr StateData, UrlReference;
        public uint ProviderFlags, UiContext;
        public IntPtr SignatureSettings;
    }

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);
}
