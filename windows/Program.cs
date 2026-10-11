using System.Diagnostics;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.Windows.Forms;
using System.Xml.Linq;

ApplicationConfiguration.Initialize();
Application.Run(new InstallerForm());

public sealed class InstallerForm : Form
{
    private readonly Dictionary<string, CheckBox> boxes = new();

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
        Style = ProgressBarStyle.Marquee,
        Visible = false
    };

    private readonly string[] apps =
    {
        "Word",
        "Excel",
        "PowerPoint",
        "Outlook",
        "OneNote",
        "Access"
    };

    // Официальная страница загрузки Office Deployment Tool.
    // Прямая ссылка на исполняемый файл может меняться.
    private const string OdtDownloadPage =
        "https://www.microsoft.com/en-us/download/details.aspx?id=49117";

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

        layout.Controls.Add(browse);

        browse.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Office Deployment Tool (setup.exe)|setup.exe",
                Title = "Выберите setup.exe из Office Deployment Tool"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                odtPath.Text = dialog.FileName;
            }
        };

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

        try
        {
            string setupPath;

            if (File.Exists(odtPath.Text))
            {
                setupPath = odtPath.Text;

                status.Text = "Проверка выбранного Office Deployment Tool...";

                await Task.Run(() => VerifyMicrosoftSignature(setupPath));
            }
            else
            {
                status.Text = "Получение Office Deployment Tool...";

                setupPath = await DownloadAndExtractOdtAsync();
                odtPath.Text = setupPath;
            }

            status.Text = "Создание конфигурации Office LTSC 2024...";

            var configurationPath = CreateConfiguration();

            status.Text =
                "Установка Office LTSC 2024. " +
                "Загрузка компонентов может занять некоторое время.";

            var startInfo = new ProcessStartInfo
            {
                FileName = setupPath,
                Arguments = $"/configure \"{configurationPath}\"",
                UseShellExecute = false,
                WorkingDirectory = Path.GetDirectoryName(setupPath)!
            };

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException(
                    "Не удалось запустить Office Deployment Tool."
                );

            await process.WaitForExitAsync();

            if (process.ExitCode == 0)
            {
                status.Text =
                    "Установка завершена. Проверьте приложения Office.";

                MessageBox.Show(
                    "Office Deployment Tool завершил установку.\n" +
                    "Для использования Office необходима действующая лицензия.",
                    "Установка завершена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            else
            {
                status.Text =
                    $"Ошибка установки. Код ODT: {process.ExitCode}.";

                MessageBox.Show(
                    $"Office Deployment Tool завершился с кодом {process.ExitCode}.\n" +
                    "Проверьте журналы установки Office.",
                    "Ошибка установки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
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
        // Microsoft периодически изменяет адрес загрузки ODT.
        // Поэтому не используем неподтверждённую прямую ссылку
        // на исполняемый файл.

        await Task.CompletedTask;

        throw new InvalidOperationException(
            "Для автоматической загрузки ODT пока не настроена " +
            "прямая ссылка на официальный установщик Microsoft.\n\n" +
            "Откройте страницу:\n" +
            OdtDownloadPage +
            "\n\nСкачайте Office Deployment Tool, распакуйте его " +
            "и выберите setup.exe кнопкой «Выбрать ODT вручную»."
        );
    }

    private static void VerifyMicrosoftSignature(string filePath)
    {
        // Проверяем наличие подписи и цепочку сертификатов.
        // Дополнительно проверяем издателя.

        var certificate = X509Certificate.CreateFromSignedFile(filePath);

        using var certificate2 = new X509Certificate2(certificate);

        using var chain = new X509Chain();

        chain.ChainPolicy.RevocationMode = X509RevocationMode.Online;

        if (!chain.Build(certificate2))
        {
            throw new InvalidOperationException(
                "Не удалось проверить сертификат издателя ODT."
            );
        }

        var subject = certificate2.Subject;

        if (!subject.Contains(
            "Microsoft Corporation",
            StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Выбранный файл не имеет сертификата Microsoft Corporation."
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
                product.Add(
                    new XElement(
                        "ExcludeApp",
                        new XAttribute("ID", app)
                    )
                );
            }
        }

        foreach (var app in new[] { "Lync", "Groove" })
        {
            product.Add(
                new XElement(
                    "ExcludeApp",
                    new XAttribute("ID", app)
                )
            );
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

        var folder = Path.Combine(
            Path.GetTempPath(),
            "OfficeInstaller",
            Guid.NewGuid().ToString("N")
        );

        Directory.CreateDirectory(folder);

        var configurationPath = Path.Combine(
            folder,
            "configuration.xml"
        );

        configuration.Save(configurationPath);

        return configurationPath;
    }
}