using System.Diagnostics;
using System.Drawing;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Windows.Forms;
using System.Xml.Linq;

ApplicationConfiguration.Initialize();
Application.Run(new InstallerForm());

public sealed class InstallerForm : Form
{
    private const string OdtUrl =
        "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20326-20112.exe";

    private const string OdtSha256 =
        "fbb64358fd4168acd52ee4efe47ffd032b6231dfb415ae2dce61b0e58ba67f86";

    private readonly Dictionary<string, CheckBox> boxes = new();

    private readonly string[] apps =
    {
        "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access"
    };

    private readonly Label status = new()
    {
        AutoSize = true,
        Text = "Выберите приложения Office LTSC 2024."
    };

    private readonly Button install = new()
    {
        Text = "Скачать и установить Office",
        AutoSize = true
    };

    private readonly Button browse = new()
    {
        Text = "Выбрать ODT вручную",
        AutoSize = true
    };

    private readonly TextBox odtPath = new()
    {
        Width = 490,
        ReadOnly = true
    };

    private readonly ProgressBar progress = new()
    {
        Width = 490,
        Height = 22,
        Visible = false
    };

    public InstallerForm()
    {
        Text = "Office Installer — Office LTSC 2024";
        ClientSize = new Size(560, 560);
        MinimumSize = new Size(580, 580);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);

        var layout = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(18)
        };

        Controls.Add(layout);

        layout.Controls.Add(new Label
        {
            Text = "Office LTSC 2024",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            AutoSize = true
        });

        layout.Controls.Add(new Label
        {
            Text = "Выберите приложения для установки.",
            AutoSize = true
        });

        layout.Controls.Add(new Label
        {
            Text = "Требуется действующая корпоративная лицензия.",
            AutoSize = true
        });

        foreach (var app in apps)
        {
            var box = new CheckBox
            {
                Text = app,
                Checked = app is "Word" or "Excel" or "PowerPoint" or "Outlook",
                AutoSize = true
            };

            boxes[app] = box;
            layout.Controls.Add(box);
        }

        layout.Controls.Add(new Label
        {
            Text = "Office Deployment Tool",
            Font = new Font(Font, FontStyle.Bold),
            AutoSize = true
        });

        layout.Controls.Add(odtPath);

        browse.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Office Deployment Tool (setup.exe)|setup.exe",
                Title = "Выберите setup.exe из Office Deployment Tool"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                odtPath.Text = dialog.FileName;
        };

        layout.Controls.Add(browse);
        layout.Controls.Add(install);
        layout.Controls.Add(progress);

        status.MaximumSize = new Size(490, 0);
        layout.Controls.Add(status);

        install.Click += async (_, _) => await InstallAsync();
    }

    private async Task InstallAsync()
    {
        if (!boxes.Values.Any(x => x.Checked))
        {
            MessageBox.Show("Выберите хотя бы одно приложение.");
            return;
        }

        using var identity = WindowsIdentity.GetCurrent();

        if (!new WindowsPrincipal(identity)
            .IsInRole(WindowsBuiltInRole.Administrator))
        {
            MessageBox.Show(
                "Запустите Office Installer от имени администратора.",
                "Требуются права администратора",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        install.Enabled = false;
        browse.Enabled = false;
        progress.Visible = true;
        progress.Style = ProgressBarStyle.Marquee;

        try
        {
            string setupPath;

            if (File.Exists(odtPath.Text))
            {
                setupPath = odtPath.Text;
                status.Text = "Проверка выбранного setup.exe...";
            }
            else
            {
                status.Text = "Скачивание Office Deployment Tool...";
                setupPath = await DownloadAndExtractOdtAsync();
                odtPath.Text = setupPath;
            }

            status.Text = "Проверка цифровой подписи Microsoft...";

            await VerifyAuthenticodeAsync(setupPath);

            status.Text = "Подготовка конфигурации Office LTSC 2024...";

            string configurationPath = CreateConfiguration();

            status.Text = "Установка Office. Дождитесь завершения...";

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = setupPath,
                Arguments = $"/configure \"{configurationPath}\"",
                UseShellExecute = false,
                WorkingDirectory = Path.GetDirectoryName(setupPath)!
            }) ?? throw new InvalidOperationException(
                "Не удалось запустить Office Deployment Tool."
            );

            await process.WaitForExitAsync();

            if (process.ExitCode == 0)
            {
                status.Text = "ODT завершил установку Office.";

                MessageBox.Show(
                    "Office Deployment Tool завершил работу.\n" +
                    "Проверьте установленные приложения Office.",
                    "Установка завершена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            else
            {
                throw new InvalidOperationException(
                    $"Office Deployment Tool завершился с кодом {process.ExitCode}."
                );
            }
        }
        catch (Exception ex)
        {
            status.Text = "Ошибка: " + ex.Message;

            MessageBox.Show(
                ex.Message,
                "Ошибка Office Installer",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            progress.Visible = false;
            install.Enabled = true;
            browse.Enabled = true;
        }
    }

    private async Task<string> DownloadAndExtractOdtAsync()
    {
        string folder = Path.Combine(
            Path.GetTempPath(),
            "OfficeInstaller",
            "ODT"
        );

        Directory.CreateDirectory(folder);

        string installerPath = Path.Combine(
            folder,
            "OfficeDeploymentTool.exe"
        );

        string setupPath = Path.Combine(folder, "setup.exe");

        if (!File.Exists(setupPath))
        {
            using var client = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(10)
            };

            using var response = await client.GetAsync(
                OdtUrl,
                HttpCompletionOption.ResponseHeadersRead
            );

            response.EnsureSuccessStatusCode();

            await using (var source = await response.Content.ReadAsStreamAsync())
            await using (var destination = File.Create(installerPath))
            {
                await source.CopyToAsync(destination);
            }

            status.Text = "Проверка SHA-256...";

            string actualHash;

            using (var stream = File.OpenRead(installerPath))
            {
                actualHash = Convert.ToHexString(
                    await SHA256.HashDataAsync(stream)
                );
            }

            if (!actualHash.Equals(
                OdtSha256,
                StringComparison.OrdinalIgnoreCase))
            {
                File.Delete(installerPath);

                throw new InvalidOperationException(
                    "SHA-256 скачанного ODT не совпадает с проверенным значением."
                );
            }

            status.Text = "Проверка подписи скачанного ODT...";

            await VerifyAuthenticodeAsync(installerPath);

            status.Text = "Распаковка Office Deployment Tool...";

            using var extraction = Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = $"/quiet /extract:\"{folder}\"",
                UseShellExecute = false,
                WorkingDirectory = folder
            }) ?? throw new InvalidOperationException(
                "Не удалось запустить распаковку ODT."
            );

            await extraction.WaitForExitAsync();

            if (extraction.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"Ошибка распаковки ODT: {extraction.ExitCode}."
                );
            }
        }

        if (!File.Exists(setupPath))
        {
            throw new FileNotFoundException(
                "После распаковки ODT файл setup.exe не найден.",
                setupPath
            );
        }

        return setupPath;
    }

    private static async Task VerifyAuthenticodeAsync(string filePath)
    {
        string escapedPath = filePath.Replace("'", "''");

        string script =
            "$s = Get-AuthenticodeSignature -LiteralPath '" +
            escapedPath +
            "'; " +
            "if ($s.Status -ne 'Valid' -or " +
            "$s.SignerCertificate.Subject -notmatch 'Microsoft Corporation') " +
            "{ exit 1 }";

        string encodedCommand = Convert.ToBase64String(
            System.Text.Encoding.Unicode.GetBytes(script)
        );

        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -NonInteractive -EncodedCommand " +
                        encodedCommand,
            UseShellExecute = false,
            CreateNoWindow = true
        }) ?? throw new InvalidOperationException(
            "Не удалось запустить проверку Authenticode."
        );

        await process.WaitForExitAsync();

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                "Файл не прошёл проверку цифровой подписи Microsoft."
            );
        }
    }

    private string CreateConfiguration()
    {
        var product = new XElement(
            "Product",
            new XAttribute("ID", "ProPlus2024Volume"),
            new XElement(
                "Language",
                new XAttribute("ID", "ru-ru")
            )
        );

        foreach (var app in apps)
        {
            if (!boxes[app].Checked)
            {
                product.Add(new XElement(
                    "ExcludeApp",
                    new XAttribute("ID", app)
                ));
            }
        }

        foreach (var app in new[] { "Lync", "Groove" })
        {
            product.Add(new XElement(
                "ExcludeApp",
                new XAttribute("ID", app)
            ));
        }

        var configuration = new XDocument(
            new XElement(
                "Configuration",
                new XElement(
                    "Add",
                    new XAttribute("OfficeClientEdition", "64"),
                    new XAttribute("Channel", "PerpetualVL2024"),
                    product
                ),
                new XElement(
                    "Display",
                    new XAttribute("Level", "Full"),
                    new XAttribute("AcceptEULA", "FALSE")
                ),
                new XElement(
                    "Updates",
                    new XAttribute("Enabled", "TRUE")
                )
            )
        );

        string folder = Path.Combine(
            Path.GetTempPath(),
            "OfficeInstaller",
            Guid.NewGuid().ToString("N")
        );

        Directory.CreateDirectory(folder);

        string configurationPath = Path.Combine(
            folder,
            "configuration.xml"
        );

        configuration.Save(configurationPath);

        return configurationPath;
    }
}