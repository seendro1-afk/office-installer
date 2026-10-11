using System.Diagnostics;
using System.Drawing;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
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

    private static readonly string OdtFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OfficeInstaller",
        "ODT"
    );

    private static readonly string[] Apps =
    {
        "Word",
        "Excel",
        "PowerPoint",
        "Outlook",
        "OneNote",
        "Access"
    };

    private static readonly Dictionary<string, string> Executables = new()
    {
        ["Word"] = "WINWORD.EXE",
        ["Excel"] = "EXCEL.EXE",
        ["PowerPoint"] = "POWERPNT.EXE",
        ["Outlook"] = "OUTLOOK.EXE",
        ["OneNote"] = "ONENOTE.EXE",
        ["Access"] = "MSACCESS.EXE"
    };

    private readonly Dictionary<string, CheckBox> boxes = new();
    private readonly Dictionary<string, Label> appStatuses = new();

    private readonly Label editionLabel = CreateInfoLabel();
    private readonly Label productLabel = CreateInfoLabel();
    private readonly Label architectureLabel = CreateInfoLabel();
    private readonly Label versionLabel = CreateInfoLabel();
    private readonly Label channelLabel = CreateInfoLabel();
    private readonly Label compatibilityLabel = CreateInfoLabel();

    private OfficeDetectionInfo? detectedOffice;

    private readonly Label status = new()
    {
        AutoSize = true,
        Text = "Проверка установленных приложений Office..."
    };

    private readonly Button refresh = new()
    {
        Text = "Обновить список",
        AutoSize = true
    };

    private readonly Button checkOdt = new()
    {
        Text = "Проверить ODT",
        AutoSize = true
    };

    private readonly Button install = new()
    {
        Text = "Установить Office",
        AutoSize = true
    };

    private readonly Button browse = new()
    {
        Text = "Выбрать setup.exe вручную",
        AutoSize = true
    };

    private readonly TextBox odtPath = new()
    {
        Width = 590,
        ReadOnly = true
    };

    private readonly ProgressBar progress = new()
    {
        Width = 590,
        Height = 22,
        Style = ProgressBarStyle.Marquee,
        Visible = false
    };

    public InstallerForm()
    {
        Text = "Office Installer — Office LTSC 2024";

        ClientSize = new Size(680, 790);
        MinimumSize = new Size(700, 700);

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
            Font = new Font("Segoe UI", 17, FontStyle.Bold),
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

        AddHeading(layout, "Установленные приложения");

        foreach (string app in Apps)
        {
            var row = new TableLayoutPanel
            {
                Width = 610,
                Height = 36,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 2, 0, 2)
            };

            row.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 190)
            );

            row.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100)
            );

            var box = new CheckBox
            {
                Text = app,
                Checked = app is
                    "Word" or
                    "Excel" or
                    "PowerPoint" or
                    "Outlook",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            };

            var appStatus = CreateInfoLabel();
            appStatus.Text = "Проверка...";
            appStatus.Anchor = AnchorStyles.Left;

            boxes[app] = box;
            appStatuses[app] = appStatus;

            row.Controls.Add(box, 0, 0);
            row.Controls.Add(appStatus, 1, 0);

            layout.Controls.Add(row);
        }

        refresh.Click += (_, _) => RefreshOfficeInformation();
        layout.Controls.Add(refresh);

        AddHeading(layout, "Информация об установленном Office");

        AddInfoRow(layout, "Редакция:", editionLabel);
        AddInfoRow(layout, "Продукт:", productLabel);
        AddInfoRow(layout, "Архитектура:", architectureLabel);
        AddInfoRow(layout, "Версия:", versionLabel);
        AddInfoRow(layout, "Канал:", channelLabel);
        AddInfoRow(layout, "Совместимость:", compatibilityLabel);

        AddHeading(layout, "Office Deployment Tool");

        layout.Controls.Add(odtPath);

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

                status.Text =
                    "Выбран setup.exe. Перед использованием будет проверена подпись.";
            }
        };

        layout.Controls.Add(browse);

        var buttonRow = new FlowLayoutPanel
        {
            Width = 610,
            Height = 48,
            WrapContents = false
        };

        buttonRow.Controls.Add(checkOdt);
        buttonRow.Controls.Add(install);

        layout.Controls.Add(buttonRow);
        layout.Controls.Add(progress);

        status.MaximumSize = new Size(590, 0);
        layout.Controls.Add(status);

        checkOdt.Click += async (_, _) => await CheckOdtAsync();
        install.Click += async (_, _) => await InstallAsync();

        Shown += (_, _) => RefreshOfficeInformation();
    }

    private static Label CreateInfoLabel()
    {
        return new Label
        {
            Text = "Не определено",
            AutoSize = true,
            MaximumSize = new Size(400, 0),
            ForeColor = Color.DimGray
        };
    }

    private void AddHeading(
        FlowLayoutPanel layout,
        string text
    )
    {
        layout.Controls.Add(new Label
        {
            Text = text,
            Font = new Font(Font, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(3, 18, 3, 8)
        });
    }

    private static void AddInfoRow(
        FlowLayoutPanel layout,
        string title,
        Label value
    )
    {
        var row = new TableLayoutPanel
        {
            Width = 610,
            Height = 35,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 2, 0, 2)
        };

        row.ColumnStyles.Add(
            new ColumnStyle(SizeType.Absolute, 190)
        );

        row.ColumnStyles.Add(
            new ColumnStyle(SizeType.Percent, 100)
        );

        row.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Anchor = AnchorStyles.Left
        }, 0, 0);

        value.Anchor = AnchorStyles.Left;
        row.Controls.Add(value, 1, 0);

        layout.Controls.Add(row);
    }

    private void RefreshOfficeInformation()
    {
        RefreshInstalledApps();
        RefreshOfficeEdition();

        status.Text =
            "Проверка приложений и конфигурации Office завершена.";
    }

    private void RefreshInstalledApps()
    {
        foreach (string app in Apps)
        {
            string? path = FindOfficeExecutable(Executables[app]);

            Label label = appStatuses[app];

            if (path != null)
            {
                string? version = GetExecutableVersion(path);

                label.Text = version != null
                    ? $"Установлен — версия {version}"
                    : "Установлен";

                label.ForeColor = Color.DarkGreen;
            }
            else
            {
                label.Text = "Не найден";
                label.ForeColor = Color.DarkRed;
            }
        }
    }

    private void RefreshOfficeEdition()
    {
        detectedOffice = OfficeDetectionService.Detect();

        if (!detectedOffice.ClickToRunDetected)
        {
            editionLabel.Text = "Click-to-Run не обнаружен";
            productLabel.Text = "Не определён";
            architectureLabel.Text = "Не определена";
            versionLabel.Text = "Не определена";
            channelLabel.Text = "Не определён";

            compatibilityLabel.Text = "Требуется проверка";
            compatibilityLabel.ForeColor = Color.DarkOrange;

            return;
        }

        editionLabel.Text = detectedOffice.Edition;
        productLabel.Text = detectedOffice.ProductIds;
        architectureLabel.Text = detectedOffice.Architecture;
        versionLabel.Text = detectedOffice.Version;
        channelLabel.Text = detectedOffice.Channel;

        if (detectedOffice.RequiresCompatibilityCheck)
        {
            compatibilityLabel.Text =
                "Требуется проверка перед установкой LTSC 2024";

            compatibilityLabel.ForeColor = Color.DarkOrange;
        }
        else
        {
            compatibilityLabel.Text =
                "Обнаружен ProPlus2024Volume";

            compatibilityLabel.ForeColor = Color.DarkGreen;
        }
    }

    private static string? FindOfficeExecutable(string executableName)
    {
        var roots = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase
        );

        foreach (string variable in new[]
        {
            "ProgramFiles",
            "ProgramFiles(x86)",
            "ProgramW6432"
        })
        {
            string? root = Environment.GetEnvironmentVariable(variable);

            if (!string.IsNullOrWhiteSpace(root))
            {
                roots.Add(root);
            }
        }

        foreach (string root in roots)
        {
            foreach (string relativePath in new[]
            {
                Path.Combine(
                    "Microsoft Office",
                    "root",
                    "Office16",
                    executableName
                ),
                Path.Combine(
                    "Microsoft Office",
                    "Office16",
                    executableName
                ),
                Path.Combine(
                    "Microsoft Office",
                    "Office15",
                    executableName
                )
            })
            {
                string candidate = Path.Combine(root, relativePath);

                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static string? GetExecutableVersion(string path)
    {
        try
        {
            var info = FileVersionInfo.GetVersionInfo(path);

            return string.IsNullOrWhiteSpace(info.ProductVersion)
                ? info.FileVersion
                : info.ProductVersion;
        }
        catch
        {
            return null;
        }
    }

    private void SetBusy(bool busy)
    {
        checkOdt.Enabled = !busy;
        install.Enabled = !busy;
        browse.Enabled = !busy;
        refresh.Enabled = !busy;

        foreach (CheckBox box in boxes.Values)
        {
            box.Enabled = !busy;
        }

        progress.Visible = busy;
    }

    private async Task CheckOdtAsync()
    {
        SetBusy(true);

        try
        {
            string setupPath = await PrepareOdtAsync();

            odtPath.Text = setupPath;

            status.Text =
                "ODT готов. Проверка завершена. Office не устанавливался.";

            MessageBox.Show(
                "Office Deployment Tool успешно подготовлен.\n\n" +
                "Файл setup.exe найден и прошёл проверку цифровой подписи Microsoft.\n\n" +
                "Установка Office НЕ запускалась.",
                "Проверка ODT завершена",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            status.Text = "Ошибка проверки ODT: " + ex.Message;

            MessageBox.Show(
                ex.Message,
                "Ошибка проверки ODT",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task InstallAsync()
    {
        if (!boxes.Values.Any(box => box.Checked))
        {
            MessageBox.Show(
                "Выберите хотя бы одно приложение Office.",
                "Не выбраны приложения",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        using var identity = WindowsIdentity.GetCurrent();

        bool isAdministrator = new WindowsPrincipal(identity)
            .IsInRole(WindowsBuiltInRole.Administrator);

        if (!isAdministrator)
        {
            MessageBox.Show(
                "Для установки Office запустите программу от имени администратора.",
                "Требуются права администратора",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        RefreshOfficeEdition();

        if (detectedOffice == null ||
            !detectedOffice.ClickToRunDetected ||
            detectedOffice.RequiresCompatibilityCheck)
        {
            var compatibilityConfirmation = MessageBox.Show(
                "На компьютере обнаружена другая или неизвестная " +
                "конфигурация Microsoft Office.\n\n" +
                $"Редакция: {detectedOffice?.Edition ?? "Не определена"}\n" +
                $"Продукт: {detectedOffice?.ProductIds ?? "Не определён"}\n" +
                $"Архитектура: {detectedOffice?.Architecture ?? "Не определена"}\n\n" +
                "Установка Office LTSC 2024 может изменить существующий Office.\n" +
                "Некоторые сочетания продуктов и архитектур несовместимы.\n\n" +
                "Продолжить установку?",
                "Предупреждение о совместимости Office",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (compatibilityConfirmation != DialogResult.Yes)
            {
                return;
            }
        }

        string selectedApps = string.Join(
            ", ",
            Apps.Where(app => boxes[app].Checked)
        );

        var confirmation = MessageBox.Show(
            "Будут установлены выбранные приложения:\n\n" +
            selectedApps +
            "\n\nOffice Deployment Tool может изменить существующую " +
            "установку Microsoft Office.\n\nПродолжить?",
            "Подтверждение установки",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        );

        if (confirmation != DialogResult.Yes)
        {
            return;
        }

        SetBusy(true);

        try
        {
            string setupPath = await PrepareOdtAsync();

            odtPath.Text = setupPath;

            status.Text = "Создание конфигурации Office LTSC 2024...";

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

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"Office Deployment Tool завершился с кодом {process.ExitCode}."
                );
            }

            RefreshOfficeInformation();

            status.Text =
                "ODT завершил работу. Проверьте установленные приложения.";

            MessageBox.Show(
                "Office Deployment Tool завершил работу.\n\n" +
                "Список установленных приложений обновлён.\n" +
                "Проверьте установку и лицензирование Office.",
                "Установка завершена",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            status.Text = "Ошибка установки: " + ex.Message;

            MessageBox.Show(
                ex.Message,
                "Ошибка установки Office",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task<string> PrepareOdtAsync()
    {
        if (File.Exists(odtPath.Text))
        {
            status.Text = "Проверка выбранного setup.exe...";

            await VerifyAuthenticodeAsync(odtPath.Text);

            return odtPath.Text;
        }

        Directory.CreateDirectory(OdtFolder);

        string installerPath = Path.Combine(
            OdtFolder,
            "OfficeDeploymentTool.exe"
        );

        string setupPath = Path.Combine(
            OdtFolder,
            "setup.exe"
        );

        if (File.Exists(setupPath))
        {
            status.Text = "Проверка ранее распакованного setup.exe...";

            try
            {
                await VerifyAuthenticodeAsync(setupPath);

                return setupPath;
            }
            catch
            {
                File.Delete(setupPath);
            }
        }

        status.Text = "Скачивание Office Deployment Tool...";

        using (var client = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(10)
        })
        {
            using var response = await client.GetAsync(
                OdtUrl,
                HttpCompletionOption.ResponseHeadersRead
            );

            response.EnsureSuccessStatusCode();

            await using var source =
                await response.Content.ReadAsStreamAsync();

            await using var destination =
                File.Create(installerPath);

            await source.CopyToAsync(destination);
        }

        status.Text = "Проверка SHA-256 скачанного ODT...";

        string actualHash;

        await using (var stream = File.OpenRead(installerPath))
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

        status.Text = "Проверка цифровой подписи ODT...";

        await VerifyAuthenticodeAsync(installerPath);

        status.Text = "Распаковка Office Deployment Tool...";

        using var extraction = Process.Start(new ProcessStartInfo
        {
            FileName = installerPath,
            Arguments = $"/quiet /extract:\"{OdtFolder}\"",
            UseShellExecute = true,
            Verb = "runas",
            WorkingDirectory = OdtFolder
        }) ?? throw new InvalidOperationException(
            "Не удалось запустить распаковку Office Deployment Tool."
        );

        await extraction.WaitForExitAsync();

        if (extraction.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Ошибка распаковки ODT. Код: {extraction.ExitCode}."
            );
        }

        if (!File.Exists(setupPath))
        {
            throw new FileNotFoundException(
                "После распаковки ODT не найден setup.exe.",
                setupPath
            );
        }

        status.Text = "Проверка подписи распакованного setup.exe...";

        await VerifyAuthenticodeAsync(setupPath);

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
            "$null -eq $s.SignerCertificate -or " +
            "$s.SignerCertificate.Subject -notmatch 'Microsoft Corporation') " +
            "{ exit 1 }";

        string encodedCommand = Convert.ToBase64String(
            Encoding.Unicode.GetBytes(script)
        );

        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments =
                "-NoProfile -NonInteractive -EncodedCommand " +
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

        foreach (string app in Apps)
        {
            if (!boxes[app].Checked)
            {
                product.Add(new XElement(
                    "ExcludeApp",
                    new XAttribute("ID", app)
                ));
            }
        }

        foreach (string app in new[] { "Lync", "Groove" })
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