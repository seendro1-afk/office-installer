using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using OfficeInstaller.Core;

namespace OfficeInstaller.Services;

public sealed record DeploymentSession(InstallationRequest Request, string Directory, string SetupPath,
    string DownloadConfigPath, string InstallConfigPath, string SourcePath, string LogPath);

public sealed class OfficeDeploymentService
{
    public const string DownloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=49117";
    public static string SessionsRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeInstaller", "Sessions");
    private readonly OfficeDetectionService detection = new();

    public async Task<DeploymentSession> PrepareAndDownloadAsync(InstallationRequest request, string? selectedTool,
        IProgress<string> progress)
    {
        OfficeDeployment.Validate(request);
        if (!Environment.Is64BitOperatingSystem && request.Architecture == OfficeArchitecture.X64)
            throw new InvalidOperationException("64-разрядный Office нельзя установить на 32-разрядную Windows.");
        var directory = Path.Combine(SessionsRoot, Guid.NewGuid().ToString("N"));
        System.IO.Directory.CreateDirectory(directory);
        var source = Path.Combine(directory, "Packages");
        System.IO.Directory.CreateDirectory(source);
        var setup = Path.Combine(directory, "setup.exe");
        var session = new DeploymentSession(request, directory, setup,
            Path.Combine(directory, "download.xml"), Path.Combine(directory, "install.xml"),
            source, Path.Combine(directory, "operations.log"));
        Log(session, "Создана сессия: " + request.ProductId);
        progress.Report("Подготавливаем Office Deployment Tool…");

        if (selectedTool is not null)
        {
            // Copy first, then verify the exact copy that will run. Never change the original file.
            File.Copy(selectedTool, setup, overwrite: false);
        }
        else
        {
            var bootstrap = Path.Combine(directory, "officedeploymenttool.exe");
            try { await DownloadToolAsync(bootstrap, progress); }
            catch (OperationCanceledException exception)
            {
                throw new TimeoutException("Истекло время ожидания загрузки ODT. Проверьте сеть или выберите setup.exe вручную. Частичные файлы сохранены: " + directory, exception);
            }
            progress.Report("Проверяем подпись и распаковываем ODT…");
            using var bootstrapLock = ReadLock(bootstrap);
            await Task.Run(() => MicrosoftSignatureVerifier.Verify(bootstrap));
            var extractionCode = await WindowsProcess.RunAsync(bootstrap,
                ["/quiet", "/extract:" + directory], directory, elevated: false);
            Log(session, "Распаковка ODT: " + extractionCode);
            if (extractionCode != 0 || !File.Exists(setup))
                throw new InvalidOperationException("Не удалось распаковать ODT. Загрузите его с сайта Microsoft и выберите извлечённый setup.exe вручную. Файлы сессии сохранены: " + directory);
        }

        using var setupLock = ReadLock(setup);
        progress.Report("Проверяем подпись Microsoft у setup.exe…");
        await Task.Run(() => MicrosoftSignatureVerifier.Verify(setup, deploymentTool: true));
        await WriteNewAsync(session.DownloadConfigPath, OfficeDeployment.Configuration(request, source, download: true));
        await WriteNewAsync(session.InstallConfigPath, OfficeDeployment.Configuration(request, source, download: false));
        using var configLock = ReadLock(session.DownloadConfigPath);
        progress.Report("Загружаем пакеты Office с серверов Microsoft. Может появиться запрос прав администратора; установка ещё не запускается…");
        Log(session, "Запуск ODT /download");
        var code = await WindowsProcess.RunAsync(setup, ["/download", session.DownloadConfigPath], directory, elevated: true);
        Log(session, "ODT /download завершился: " + code);
        if (code != 0) throw new InvalidOperationException($"Загрузка Office завершилась с кодом {code}. Журнал: {session.LogPath}");
        CheckPackageFiles(session);
        var manifest = new PreparedDownload(request.ProductId, request.Architecture, request.Language, request.Applications.Order().ToArray());
        await WriteNewAsync(Path.Combine(directory, "prepared.json"), JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
        progress.Report("Пакеты скачаны. Для установки нажмите отдельную кнопку «Установить». Файлы: " + source);
        return session;
    }

    public async Task<DeploymentSession> LoadSessionAsync(string manifestPath)
    {
        var fullPath = Path.GetFullPath(manifestPath);
        var directory = Path.GetDirectoryName(fullPath) ?? throw new ArgumentException("Нет каталога пакетов.");
        if (!Path.GetFileName(fullPath).Equals("prepared.json", StringComparison.OrdinalIgnoreCase)
            || !string.Equals(Path.GetDirectoryName(directory), SessionsRoot, StringComparison.OrdinalIgnoreCase)
            || !Guid.TryParseExact(Path.GetFileName(directory), "N", out _))
            throw new ArgumentException("Выберите prepared.json в каталоге сессии этой программы: " + SessionsRoot);
        using var manifestLock = ReadLock(fullPath);
        if (manifestLock.Length > 65536) throw new ArgumentException("Описание пакетов слишком большое.");
        var manifest = await JsonSerializer.DeserializeAsync<PreparedDownload>(manifestLock)
            ?? throw new ArgumentException("Не удалось прочитать описание скачанного пакета.");
        var request = manifest.ToRequest();
        var source = Path.Combine(directory, "Packages");
        var session = new DeploymentSession(request, directory, Path.Combine(directory, "setup.exe"),
            Path.Combine(directory, "download.xml"), Path.Combine(directory, "install.xml"), source, Path.Combine(directory, "operations.log"));
        using var setupLock = ReadLock(session.SetupPath);
        using var configLock = ReadLock(session.InstallConfigPath);
        await Task.Run(() => MicrosoftSignatureVerifier.Verify(session.SetupPath, deploymentTool: true));
        if (!string.Equals(await File.ReadAllTextAsync(session.InstallConfigPath),
            OfficeDeployment.Configuration(request, source, download: false), StringComparison.Ordinal))
            throw new InvalidOperationException("Конфигурация сохранённой сессии изменилась. Подготовьте пакеты заново.");
        CheckPackageFiles(session);
        return session;
    }

    public async Task<int> InstallAsync(DeploymentSession session, IProgress<string> progress)
    {
        OfficeDeployment.Validate(session.Request);
        using var setupLock = ReadLock(session.SetupPath);
        using var configLock = ReadLock(session.InstallConfigPath);
        progress.Report("Повторно проверяем подпись, конфигурацию и установленный Office…");
        await Task.Run(() => MicrosoftSignatureVerifier.Verify(session.SetupPath, deploymentTool: true));
        var expected = OfficeDeployment.Configuration(session.Request, session.SourcePath, download: false);
        if (!string.Equals(await File.ReadAllTextAsync(session.InstallConfigPath), expected, StringComparison.Ordinal))
            throw new InvalidOperationException("Конфигурация установки была изменена. Подготовьте пакеты заново.");
        var scan = await Task.Run(detection.Scan);
        OfficeDeployment.CheckNewInstallation(scan);
        CheckPackageFiles(session);
        progress.Report("Запускаем установку Office. Следуйте указаниям установщика Microsoft…");
        Log(session, "Запуск ODT /configure");
        var code = await WindowsProcess.RunAsync(session.SetupPath, ["/configure", session.InstallConfigPath], session.Directory, elevated: true);
        Log(session, "ODT /configure завершился: " + code);
        return code;
    }

    private static void CheckPackageFiles(DeploymentSession session)
    {
        var data = Path.Combine(session.SourcePath, "Office", "Data");
        var cab = Path.Combine(data, session.Request.Architecture == OfficeArchitecture.X64 ? "v64.cab" : "v32.cab");
        var stream = session.Request.Architecture == OfficeArchitecture.X64 ? "stream.x64.x-none.dat" : "stream.x86.x-none.dat";
        if (!File.Exists(cab) || new FileInfo(cab).Length == 0 || !System.IO.Directory.Exists(data)
            || !System.IO.Directory.EnumerateFiles(data, stream, SearchOption.AllDirectories).Any(p => new FileInfo(p).Length > 0))
            throw new InvalidOperationException("ODT завершился, но ожидаемые файлы пакета не найдены. Установка не будет запущена. Каталог: " + session.SourcePath);
    }

    private static async Task DownloadToolAsync(string destination, IProgress<string> progress)
    {
        using var handler = new HttpClientHandler { AllowAutoRedirect = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(5) };
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(5));
        var token = deadline.Token;
        client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficeInstaller/0.2");
        using var page = await GetTrustedAsync(client, new Uri(DownloadPage), MicrosoftDownloadPolicy.IsDownloadPage, token);
        var html = await ReadPageAsync(page, 4 * 1024 * 1024, token);
        var link = MicrosoftDownloadPolicy.FindDeploymentTool(html);
        if (link is null)
            throw new InvalidOperationException("Страница Microsoft изменилась или не содержит ссылки ODT. Откройте официальный сайт и выберите извлечённый setup.exe вручную.");

        using var response = await GetTrustedAsync(client, link, MicrosoftDownloadPolicy.IsDeploymentTool, token);
        const long maximum = 64 * 1024 * 1024;
        if (response.Content.Headers.ContentLength > maximum)
            throw new InvalidOperationException("Ответ сервера превышает ожидаемый размер ODT.");
        await using var input = await response.Content.ReadAsStreamAsync(token);
        await using var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        var buffer = new byte[81920];
        long total = 0;
        int count;
        while ((count = await input.ReadAsync(buffer, token)) > 0)
        {
            total += count;
            if (total > maximum) throw new InvalidOperationException("Загрузка ODT превышает допустимый размер.");
            await output.WriteAsync(buffer.AsMemory(0, count), token);
            progress.Report($"Загружаем ODT: {total / 1024:N0} КБ…");
        }
        if (total == 0) throw new InvalidOperationException("Microsoft вернула пустой файл ODT.");
    }

    private static async Task<HttpResponseMessage> GetTrustedAsync(HttpClient client, Uri uri, Func<Uri, bool> allowed, CancellationToken token)
    {
        for (var redirect = 0; redirect <= 5; redirect++)
        {
            if (!allowed(uri)) throw new InvalidOperationException("Ссылка загрузки не принадлежит разрешённому серверу Microsoft.");
            var response = await client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, token);
            if ((int)response.StatusCode is >= 300 and < 400)
            {
                var location = response.Headers.Location;
                response.Dispose();
                if (location is null) throw new InvalidOperationException("Некорректное перенаправление Microsoft.");
                uri = location.IsAbsoluteUri ? location : new Uri(uri, location);
                continue;
            }
            try { response.EnsureSuccessStatusCode(); return response; }
            catch { response.Dispose(); throw; }
        }
        throw new InvalidOperationException("Слишком много перенаправлений при загрузке ODT.");
    }

    private static async Task<string> ReadPageAsync(HttpResponseMessage response, int maximum, CancellationToken token)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(token);
        using var content = new MemoryStream();
        var buffer = new byte[16384];
        int count;
        while ((count = await stream.ReadAsync(buffer, token)) > 0)
        {
            if (content.Length + count > maximum) throw new InvalidOperationException("Страница ODT слишком большая.");
            content.Write(buffer, 0, count);
        }
        return Encoding.UTF8.GetString(content.ToArray());
    }

    private static FileStream ReadLock(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read);
    private static async Task WriteNewAsync(string path, string content)
    {
        await using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        await using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        await writer.WriteAsync(content);
    }
    private static void Log(DeploymentSession session, string message) =>
        File.AppendAllText(session.LogPath, DateTimeOffset.Now.ToString("O") + " " + message + Environment.NewLine);
}
