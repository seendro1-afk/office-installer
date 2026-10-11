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

    private const int CardWidth = 640;
    private const int ContentWidth = 590;

    private static readonly string OdtFolder = Path.Combine(
        Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData
        ),
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

    private static readonly Color Background =
        Color.FromArgb(246, 248, 251);

    private static readonly Color Surface =
        Color.White;

    private static readonly Color TextPrimary =
        Color.FromArgb(35, 43, 56);

    private static readonly Color TextSecondary =
        Color.FromArgb(100, 112, 128);

    private static readonly Color Accent =
        Color.FromArgb(37, 99, 235);

    private static readonly Color Success =
        Color.FromArgb(22, 132, 83);

    private static readonly Color Danger =
        Color.FromArgb(190, 48, 55);

    private static readonly Color DisabledButton =
        Color.FromArgb(205, 211, 220);

    private readonly Dictionary<string, CheckBox> boxes = new();
    private readonly Dictionary<string, Label> appStatuses = new();

    private readonly Label officeProductLabel = CreateValueLabel();
    private readonly Label projectProductLabel = CreateValueLabel();
    private readonly Label visioProductLabel = CreateValueLabel();

    private readonly Label architectureLabel = CreateValueLabel();
    private readonly Label versionLabel = CreateValueLabel();
    private readonly Label channelLabel = CreateValueLabel();

    private readonly Label compatibilityTitle = new()
    {
        AutoSize = true,
        Font = new Font("Segoe UI", 11, FontStyle.Bold)
    };

    private readonly Label compatibilityDetails = new()
    {
        AutoSize = true,
        MaximumSize = new Size(550, 0)
    };

    private readonly Panel compatibilityPanel = new()
    {
        Width = ContentWidth,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        Padding = new Padding(14)
    };

    private readonly ComboBox architectureBox = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Width = 220,
        Font = new Font("Segoe UI", 10)
    };

    private readonly Label architectureHint = new()
    {
        AutoSize = true,
        MaximumSize = new Size(550, 0),
        ForeColor = TextSecondary
    };

    private readonly TextBox odtPath = new()
    {
        Width = ContentWidth,
        ReadOnly = true,
        BackColor = Color.White
    };

    private readonly Button refresh =
        CreateButton("Обновить список");

    private readonly Button checkOdt =
        CreateButton("Проверить ODT");

    private readonly Button install =
        CreateButton("Установить Office");

    private readonly Button browse =
        CreateButton("Выбрать setup.exe");

    private readonly Button openLogs =
        CreateButton("Открыть журнал");

    private readonly Button diagnosticsButton =
        CreateButton("Диагностика");

    private readonly ProgressBar progress = new()
    {
        Width = ContentWidth,
        Height = 18,
        Style = ProgressBarStyle.Marquee,
        Visible = false
    };

    private readonly Label status = new()
    {
        AutoSize = true,
        MaximumSize = new Size(550, 0),
        ForeColor = TextSecondary,
        Text = "Проверка конфигурации Office..."
    };

    private OfficeDetectionInfo? detectedOffice;

    private bool officeExecutablesDetected;
    private bool isBusy;

    public InstallerForm()
    {
        Text = "Office Installer — Office LTSC 2024";

        ClientSize = new Size(720, 850);
        MinimumSize = new Size(740, 650);

        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Background;

        Font = new Font("Segoe UI", 10);

        var scroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true
        };

        Controls.Add(scroll);

        var layout = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,

            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,

            Padding = new Padding(24),
            BackColor = Background
        };

        scroll.Controls.Add(layout);

        layout.Controls.Add(new Label
        {
            Text = "Office Installer",
            Font = new Font("Segoe UI", 22, FontStyle.Bold),
            ForeColor = TextPrimary,
            AutoSize = true
        });

        layout.Controls.Add(new Label
        {
            Text = "Microsoft Office LTSC 2024",
            Font = new Font("Segoe UI", 11),
            ForeColor = TextSecondary,
            AutoSize = true,
            Margin = new Padding(3, 0, 3, 12)
        });

        // Приложения Office

        var appsCard = CreateCard("Приложения Office");

        foreach (string app in Apps)
        {
            var row = new TableLayoutPanel
            {
                Width = ContentWidth,
                Height = 34,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 1, 0, 1)
            };

            row.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 180)
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
                ForeColor = TextPrimary,
                Anchor = AnchorStyles.Left
            };

            var appStatus = CreateValueLabel();

            appStatus.Text = "Проверка...";
            appStatus.Anchor = AnchorStyles.Left;

            boxes[app] = box;
            appStatuses[app] = appStatus;

            row.Controls.Add(box, 0, 0);
            row.Controls.Add(appStatus, 1, 0);

            appsCard.Controls.Add(row);
        }

        refresh.Click += (_, _) => RefreshOfficeInformation();

        appsCard.Controls.Add(refresh);
        layout.Controls.Add(appsCard);

        // Установленные продукты

        var productsCard = CreateCard("Установленные продукты");

        AddInfoRow(
            productsCard,
            "Office",
            officeProductLabel
        );

        AddInfoRow(
            productsCard,
            "Project",
            projectProductLabel
        );

        AddInfoRow(
            productsCard,
            "Visio",
            visioProductLabel
        );

        AddInfoRow(
            productsCard,
            "Архитектура",
            architectureLabel
        );

        AddInfoRow(
            productsCard,
            "Версия",
            versionLabel
        );

        AddInfoRow(
            productsCard,
            "Канал",
            channelLabel
        );

        layout.Controls.Add(productsCard);

        // Архитектура установки

        var architectureCard =
            CreateCard("Архитектура установки");

        architectureBox.Items.AddRange(new object[]
        {
            "x64 (64-битная)",
            "x86 (32-битная)"
        });

        architectureBox.SelectedIndex = 0;

        architectureBox.SelectedIndexChanged += (_, _) =>
            UpdateCompatibilityDisplay();

        architectureCard.Controls.Add(architectureBox);

        architectureCard.Controls.Add(architectureHint);

        layout.Controls.Add(architectureCard);

        // Проверка совместимости

        var compatibilityCard =
            CreateCard("Проверка совместимости");

        compatibilityPanel.BackColor =
            Color.FromArgb(255, 240, 240);

        var compatibilityContent = new FlowLayoutPanel
        {
            Width = 550,

            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,

            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        compatibilityContent.Controls.Add(
            compatibilityTitle
        );

        compatibilityContent.Controls.Add(
            compatibilityDetails
        );

        compatibilityPanel.Controls.Add(
            compatibilityContent
        );

        compatibilityCard.Controls.Add(
            compatibilityPanel
        );

        layout.Controls.Add(compatibilityCard);

        // Office Deployment Tool

        var odtCard = CreateCard("Office Deployment Tool");

        odtCard.Controls.Add(new Label
        {
            Text = "Путь к setup.exe",
            ForeColor = TextSecondary,
            AutoSize = true
        });

        odtCard.Controls.Add(odtPath);

        browse.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog
            {
                Filter =
                    "Office Deployment Tool (setup.exe)|setup.exe",

                Title =
                    "Выберите setup.exe из Office Deployment Tool"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                odtPath.Text = dialog.FileName;

                status.Text =
                    "Выбран setup.exe. Подпись будет проверена.";
            }
        };

        odtCard.Controls.Add(browse);

        var buttonRow = new FlowLayoutPanel
        {
            Width = ContentWidth,
            Height = 44,
            WrapContents = false
        };

        install.FlatStyle = FlatStyle.Flat;
        install.FlatAppearance.BorderSize = 0;

        buttonRow.Controls.Add(checkOdt);
        buttonRow.Controls.Add(install);

        odtCard.Controls.Add(buttonRow);
        openLogs.Click += (_, _) =>
        {
            try
            {
                InstallerLogger.OpenDirectory();
            }
            catch (Exception ex)
            {
                InstallerLogger.Error("Открытие журнала", ex);

                MessageBox.Show(
                    ex.Message,
                    "Ошибка открытия журнала",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        };

        odtCard.Controls.Add(openLogs);

        diagnosticsButton.Click += (_, _) => RunDiagnostics();

        odtCard.Controls.Add(diagnosticsButton);



        odtCard.Controls.Add(progress);
        odtCard.Controls.Add(status);

        layout.Controls.Add(odtCard);

        layout.Controls.Add(new Label
        {
            Text =
                "Для установки необходима действующая корпоративная лицензия.",

            ForeColor = TextSecondary,
            AutoSize = true,
            Margin = new Padding(3, 5, 3, 8)
        });

        checkOdt.Click += async (_, _) =>
            await CheckOdtAsync();

        install.Click += async (_, _) =>
            await InstallAsync();

        Shown += (_, _) =>
        {
            InstallerLogger.Write("Запуск Office Installer.");
            RefreshOfficeInformation();
        };
    }

    private static Button CreateButton(string text)
    {
        return new Button
        {
            Text = text,
            AutoSize = true,
            Height = 34,
            Padding = new Padding(10, 3, 10, 3),
            FlatStyle = FlatStyle.System
        };
    }

    private static Label CreateValueLabel()
    {
        return new Label
        {
            Text = "Не определено",
            AutoSize = true,
            MaximumSize = new Size(390, 0),
            ForeColor = TextSecondary
        };
    }

    private static FlowLayoutPanel CreateCard(string title)
    {
        var card = new FlowLayoutPanel
        {
            Width = CardWidth,

            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,

            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,

            MinimumSize = new Size(CardWidth, 0),

            Padding = new Padding(15),
            BackColor = Surface,

            Margin = new Padding(0, 0, 0, 10)
        };

        card.Controls.Add(new Label
        {
            Text = title,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            ForeColor = TextPrimary,
            AutoSize = true,
            Margin = new Padding(3, 0, 3, 9)
        });

        return card;
    }

    private static void AddInfoRow(
        FlowLayoutPanel parent,
        string title,
        Label value
    )
    {
        var row = new TableLayoutPanel
        {
            Width = ContentWidth,
            Height = 36,
            ColumnCount = 2,
            RowCount = 1
        };

        row.ColumnStyles.Add(
            new ColumnStyle(SizeType.Absolute, 170)
        );

        row.ColumnStyles.Add(
            new ColumnStyle(SizeType.Percent, 100)
        );

        row.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = TextPrimary,
            Anchor = AnchorStyles.Left
        }, 0, 0);

        value.Anchor = AnchorStyles.Left;

        row.Controls.Add(value, 1, 0);

        parent.Controls.Add(row);
    }

    private void RunDiagnostics()
    {
        InstallerLogger.Write("Диагностика: кнопка нажата.");

        try
        {
            // Используем уже обнаруженные данные Office.
            // Повторное сканирование при нажатии не выполняем.
            InstallerLogger.Write("Диагностика: подготовка данных.");

            var applications = new Dictionary<string, string>();

            foreach (string app in Apps)
            {
                applications[app] = appStatuses[app].Text;
            }

            string? blockReason = GetInstallationBlockReason();

            string report = OfficeDiagnostics.GenerateReport(
                detectedOffice,
                SelectedArchitecture,
                blockReason,
                applications
            );

            InstallerLogger.Write("Диагностика: отчёт сформирован.");

            using var dialog = new SaveFileDialog
            {
                Title = "Сохранить отчёт диагностики Office",
                Filter = "Текстовый файл (*.txt)|*.txt",
                FileName = "Office-Diagnostics-" +
                    DateTime.Now.ToString("yyyy-MM-dd-HHmmss") +
                    ".txt",
                AddExtension = true,
                DefaultExt = "txt",
                OverwritePrompt = true
            };

            InstallerLogger.Write(
                "Диагностика: открытие окна сохранения.");

            var result = dialog.ShowDialog(this);

            InstallerLogger.Write(
                "Диагностика: окно сохранения закрыто. Результат: " +
                result);

            if (result != DialogResult.OK)
            {
                InstallerLogger.Write(
                    "Диагностика: сохранение отменено.");
                return;
            }

            File.WriteAllText(
                dialog.FileName,
                report,
                Encoding.UTF8
            );

            InstallerLogger.Write(
                "Диагностика: отчёт сохранён: " +
                dialog.FileName);

            status.Text = "Отчёт диагностики сохранён.";

            MessageBox.Show(
                this,
                "Диагностический отчёт сохранён.\\n\\n" +
                dialog.FileName,
                "Диагностика завершена",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            InstallerLogger.Error("Диагностика", ex);

            MessageBox.Show(
                this,
                ex.Message,
                "Ошибка диагностики",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private void RefreshOfficeInformation()
    {
        RefreshInstalledApps();
        RefreshOfficeEdition();

        InstallerLogger.Write(
            "Обнаружен Office: " +
            $"ClickToRun={detectedOffice?.ClickToRunDetected}; " +
            $"Редакция={detectedOffice?.Edition}; " +
            $"Продукты={detectedOffice?.ProductIds}; " +
            $"Архитектура={detectedOffice?.Architecture}; " +
            $"Версия={detectedOffice?.Version}; " +
            $"Канал={detectedOffice?.Channel}.");

        SelectRecommendedArchitecture();
        UpdateCompatibilityDisplay();

        status.Text =
            "Проверка конфигурации Office завершена.";
    }

    private void RefreshInstalledApps()
    {
        officeExecutablesDetected = false;

        foreach (string app in Apps)
        {
            string? path =
                FindOfficeExecutable(Executables[app]);

            Label label = appStatuses[app];

            if (path != null)
            {
                officeExecutablesDetected = true;

                string? version =
                    GetExecutableVersion(path);

                label.Text = version == null
                    ? "Установлен"
                    : $"Установлен · {version}";

                label.ForeColor = Success;
            }
            else
            {
                label.Text = "Не найден";
                label.ForeColor = Danger;
            }
        }
    }

    private void RefreshOfficeEdition()
    {
        detectedOffice = OfficeDetectionService.Detect();

        if (!detectedOffice.ClickToRunDetected)
        {
            officeProductLabel.Text = "Не определён";
            projectProductLabel.Text = "Не определён";
            visioProductLabel.Text = "Не определён";

            architectureLabel.Text = "Не определена";
            versionLabel.Text = "Не определена";
            channelLabel.Text = "Не определён";

            return;
        }

        string[] ids = GetProductIds();

        officeProductLabel.Text = DescribeProduct(
            ids,
            "ProPlus2024Volume",
            "Office LTSC Professional Plus 2024",
            "ProPlus2021Volume",
            "Office LTSC Professional Plus 2021",
            "Standard2024Volume",
            "Office LTSC Standard 2024",
            "Standard2021Volume",
            "Office LTSC Standard 2021"
        );

        projectProductLabel.Text = DescribeProduct(
            ids,
            "ProjectPro2024Volume",
            "Project Professional 2024",
            "ProjectPro2021Volume",
            "Project Professional 2021",
            "ProjectStd2024Volume",
            "Project Standard 2024",
            "ProjectStd2021Volume",
            "Project Standard 2021"
        );

        visioProductLabel.Text = DescribeProduct(
            ids,
            "VisioPro2024Volume",
            "Visio Professional 2024",
            "VisioPro2021Volume",
            "Visio Professional 2021",
            "VisioStd2024Volume",
            "Visio Standard 2024",
            "VisioStd2021Volume",
            "Visio Standard 2021"
        );

        architectureLabel.Text =
            detectedOffice.Architecture;

        versionLabel.Text =
            detectedOffice.Version;

        channelLabel.Text =
            DescribeChannel(detectedOffice.Channel);
    }

    private string[] GetProductIds()
    {
        return (detectedOffice?.ProductIds ?? "").Split(
            new[] { ',', ';' },

            StringSplitOptions.RemoveEmptyEntries |
            StringSplitOptions.TrimEntries
        );
    }

    private static string DescribeProduct(
        string[] ids,
        params string[] mappings
    )
    {
        for (int i = 0; i + 1 < mappings.Length; i += 2)
        {
            if (ids.Any(id => id.Equals(
                mappings[i],
                StringComparison.OrdinalIgnoreCase
            )))
            {
                return mappings[i + 1];
            }
        }

        return "Не обнаружен в Click-to-Run";
    }

    private static string DescribeChannel(string value)
    {
        if (value.Contains(
            "5030841d-c919-4594-8d2d-84ae4f96e58e",
            StringComparison.OrdinalIgnoreCase
        ))
        {
            return "Office LTSC 2021 (PerpetualVL2021)";
        }

        if (value.Contains(
            "7983bac0-e531-40cf-be00-fd24fe66619c",
            StringComparison.OrdinalIgnoreCase
        ))
        {
            return "Office LTSC 2024 (PerpetualVL2024)";
        }

        return value == "Не определено"
            ? "Не определён"
            : "Другой канал Click-to-Run";
    }

    private void SelectRecommendedArchitecture()
    {
        string architecture =
            detectedOffice?.Architecture ?? "";

        if (architecture.Equals(
            "x86",
            StringComparison.OrdinalIgnoreCase
        ))
        {
            architectureBox.SelectedIndex = 1;

            architectureHint.Text =
                "Автоматически выбрана x86.";
        }
        else if (architecture.Equals(
            "x64",
            StringComparison.OrdinalIgnoreCase
        ))
        {
            architectureBox.SelectedIndex = 0;

            architectureHint.Text =
                "Автоматически выбрана x64.";
        }
        else
        {
            architectureBox.SelectedIndex = 0;

            architectureHint.Text =
                "По умолчанию выбрана x64. " +
                "При неизвестной конфигурации Office " +
                "установка может быть заблокирована.";
        }
    }

    private string SelectedArchitecture =>
        architectureBox.SelectedIndex == 1
            ? "x86"
            : "x64";

    private string? GetInstallationBlockReason()
    {
        if (detectedOffice == null)
        {
            return
                "Не удалось проверить установленный Office.";
        }

        string existingArchitecture =
            detectedOffice.Architecture;

        bool knownArchitecture =
            existingArchitecture.Equals(
                "x86",
                StringComparison.OrdinalIgnoreCase
            ) ||
            existingArchitecture.Equals(
                "x64",
                StringComparison.OrdinalIgnoreCase
            );

        if (detectedOffice.ClickToRunDetected &&
            !knownArchitecture)
        {
            return
                "Обнаружен Office Click-to-Run, " +
                "но его архитектура неизвестна.";
        }

        if (knownArchitecture &&
            !existingArchitecture.Equals(
                SelectedArchitecture,
                StringComparison.OrdinalIgnoreCase
            ))
        {
            return
                $"Установлен Office {existingArchitecture}, " +
                $"а выбран {SelectedArchitecture}. " +
                "Архитектуры не совпадают.";
        }

        if (!detectedOffice.ClickToRunDetected)
        {
            return officeExecutablesDetected
                ? "Обнаружены приложения Office, " +
                  "но конфигурация Click-to-Run неизвестна."
                : null;
        }

        string[] ids = GetProductIds();

        if (ids.Length == 0 ||
            ids.Any(id => !id.Equals(
                "ProPlus2024Volume",
                StringComparison.OrdinalIgnoreCase
            )))
        {
            return
                "Обнаружены другие продукты Office, " +
                "Project или Visio. Для установки LTSC 2024 " +
                "необходима дополнительная проверка совместимости.";
        }

        return null;
    }

    private void UpdateCompatibilityDisplay()
    {
        string? reason = GetInstallationBlockReason();

        bool canInstall = reason == null;

        InstallerLogger.Write(
            canInstall
                ? $"Проверка совместимости пройдена. Архитектура: {SelectedArchitecture}."
                : $"Установка заблокирована. Архитектура: {SelectedArchitecture}. Причина: {reason}");

        if (canInstall)
        {
            compatibilityPanel.BackColor =
                Color.FromArgb(231, 248, 239);

            compatibilityTitle.Text =
                "Предварительная проверка пройдена";

            compatibilityTitle.ForeColor = Success;

            compatibilityDetails.Text =
                "Конфликтов по текущим проверкам не обнаружено.";

            compatibilityDetails.ForeColor = Success;
        }
        else
        {
            compatibilityPanel.BackColor =
                Color.FromArgb(255, 240, 240);

            compatibilityTitle.Text =
                "Установка заблокирована";

            compatibilityTitle.ForeColor = Danger;

            compatibilityDetails.Text = reason;

            compatibilityDetails.ForeColor = Danger;
        }

        // Кнопка установки доступна только тогда,
        // когда проверка совместимости пройдена.
        install.Enabled = canInstall && !isBusy;

        install.BackColor = canInstall && !isBusy
            ? Accent
            : DisabledButton;

        install.ForeColor = canInstall && !isBusy
            ? Color.White
            : TextSecondary;
    }

    private static string? FindOfficeExecutable(
        string executableName
    )
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
            string? root =
                Environment.GetEnvironmentVariable(variable);

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
                string candidate =
                    Path.Combine(root, relativePath);

                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static string? GetExecutableVersion(
        string path
    )
    {
        try
        {
            var info =
                FileVersionInfo.GetVersionInfo(path);

            return string.IsNullOrWhiteSpace(
                info.ProductVersion
            )
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
        isBusy = busy;

        checkOdt.Enabled = !busy;
        browse.Enabled = !busy;
        refresh.Enabled = !busy;
        architectureBox.Enabled = !busy;

        foreach (CheckBox box in boxes.Values)
        {
            box.Enabled = !busy;
        }

        progress.Visible = busy;

        UpdateCompatibilityDisplay();
    }

    private async Task CheckOdtAsync()
    {
        InstallerLogger.Write("Начало проверки Office Deployment Tool.");
        SetBusy(true);

        try
        {
            string setupPath =
                await PrepareOdtAsync();

            odtPath.Text = setupPath;

            InstallerLogger.Write(
                "Проверка ODT успешно завершена. Путь: " + setupPath);

            status.Text =
                "ODT готов. Office не устанавливался.";

            MessageBox.Show(
                "Office Deployment Tool успешно подготовлен.\n\n" +
                "Файл setup.exe прошёл проверку " +
                "цифровой подписи Microsoft.\n\n" +
                "Установка Office НЕ запускалась.",

                "Проверка ODT завершена",

                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            InstallerLogger.Error("CheckOdtAsync", ex);

            status.Text =
                "Ошибка проверки ODT: " + ex.Message;

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
        InstallerLogger.Write("Получен запрос на установку Office.");

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

        // Повторная проверка перед установкой.
        RefreshInstalledApps();
        RefreshOfficeEdition();
        UpdateCompatibilityDisplay();

        string? blockReason =
            GetInstallationBlockReason();

        if (blockReason != null)
        {
            InstallerLogger.Write(
                "Установка остановлена проверкой совместимости: " +
                blockReason);

            MessageBox.Show(
                "УСТАНОВКА ЗАБЛОКИРОВАНА\n\n" +
                blockReason + "\n\n" +
                "Office Deployment Tool не будет запущен.",

                "Проверка совместимости Office",

                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );

            status.Text =
                "Установка заблокирована.";

            return;
        }

        using var identity =
            WindowsIdentity.GetCurrent();

        bool isAdministrator =
            new WindowsPrincipal(identity)
                .IsInRole(
                    WindowsBuiltInRole.Administrator
                );

        if (!isAdministrator)
        {
            MessageBox.Show(
                "Запустите программу от имени администратора.",

                "Требуются права администратора",

                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        string selectedApps = string.Join(
            ", ",
            Apps.Where(app => boxes[app].Checked)
        );

        var confirmation = MessageBox.Show(
            $"Приложения: {selectedApps}\n" +
            $"Архитектура: {SelectedArchitecture}\n" +
            "Редакция: Office LTSC Professional Plus 2024\n\n" +
            "Office Deployment Tool может изменить " +
            "существующую установку Microsoft Office.\n\n" +
            "Продолжить?",

            "Подтверждение установки",

            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        if (confirmation != DialogResult.Yes)
        {
            return;
        }

        SetBusy(true);

        try
        {
            string setupPath =
                await PrepareOdtAsync();

            odtPath.Text = setupPath;

            status.Text =
                "Создание конфигурации Office LTSC 2024...";

            string configurationPath =
                CreateConfiguration();

            InstallerLogger.Write(
                $"Запуск установки Office LTSC 2024. " +
                $"Архитектура={SelectedArchitecture}; " +
                $"Приложения={selectedApps}; " +
                $"Конфигурация={configurationPath}.");

            status.Text =
                "Установка Office. Дождитесь завершения...";

            using var process =
                Process.Start(new ProcessStartInfo
                {
                    FileName = setupPath,

                    Arguments =
                        $"/configure \"{configurationPath}\"",

                    UseShellExecute = false,

                    WorkingDirectory =
                        Path.GetDirectoryName(setupPath)!
                })
                ?? throw new InvalidOperationException(
                    "Не удалось запустить Office Deployment Tool."
                );

            await process.WaitForExitAsync();

            InstallerLogger.Write(
                "Office Deployment Tool завершился. Код: " +
                process.ExitCode);

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    "Office Deployment Tool завершился " +
                    $"с кодом {process.ExitCode}."
                );
            }

            RefreshOfficeInformation();

            status.Text =
                "ODT завершил работу. " +
                "Проверьте приложения Office.";

            MessageBox.Show(
                "Office Deployment Tool завершил работу.\n\n" +
                "Проверьте установку и лицензирование Office.",

                "Установка завершена",

                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            InstallerLogger.Error("InstallAsync", ex);

            status.Text =
                "Ошибка установки: " + ex.Message;

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
            status.Text =
                "Проверка выбранного setup.exe...";

            await VerifyAuthenticodeAsync(
                odtPath.Text
            );

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
            status.Text =
                "Проверка ранее распакованного setup.exe...";

            try
            {
                await VerifyAuthenticodeAsync(
                    setupPath
                );

                return setupPath;
            }
            catch
            {
                File.Delete(setupPath);
            }
        }

        InstallerLogger.Write("Начало скачивания ODT: " + OdtUrl);

        status.Text =
            "Скачивание Office Deployment Tool...";

        using (var client = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(10)
        })
        {
            using var response =
                await client.GetAsync(
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

        status.Text =
            "Проверка SHA-256 скачанного ODT...";

        string actualHash;

        await using (var stream =
            File.OpenRead(installerPath))
        {
            actualHash = Convert.ToHexString(
                await SHA256.HashDataAsync(stream)
            );
        }

        if (!actualHash.Equals(
            OdtSha256,
            StringComparison.OrdinalIgnoreCase
        ))
        {
            File.Delete(installerPath);

            throw new InvalidOperationException(
                "SHA-256 скачанного ODT не совпадает " +
                "с проверенным значением."
            );
        }

        InstallerLogger.Write(
            "SHA-256 ODT успешно проверен: " + actualHash);

        status.Text =
            "Проверка цифровой подписи ODT...";

        await VerifyAuthenticodeAsync(
            installerPath
        );

        InstallerLogger.Write(
            "Цифровая подпись скачанного ODT проверена.");

        status.Text =
            "Распаковка Office Deployment Tool...";

        using var extraction =
            Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,

                Arguments =
                    $"/quiet /extract:\"{OdtFolder}\"",

                UseShellExecute = true,
                Verb = "runas",
                WorkingDirectory = OdtFolder
            })
            ?? throw new InvalidOperationException(
                "Не удалось запустить распаковку ODT."
            );

        await extraction.WaitForExitAsync();

        if (extraction.ExitCode != 0)
        {
            throw new InvalidOperationException(
                "Ошибка распаковки ODT. " +
                $"Код: {extraction.ExitCode}."
            );
        }

        if (!File.Exists(setupPath))
        {
            throw new FileNotFoundException(
                "После распаковки ODT не найден setup.exe.",
                setupPath
            );
        }

        status.Text =
            "Проверка подписи распакованного setup.exe...";

        await VerifyAuthenticodeAsync(setupPath);

        InstallerLogger.Write(
            "Распаковка ODT завершена. setup.exe проверен: " +
            setupPath);

        return setupPath;
    }

    private static async Task VerifyAuthenticodeAsync(
        string filePath
    )
    {
        string escapedPath =
            filePath.Replace("'", "''");

        string script =
            "$s = Get-AuthenticodeSignature -LiteralPath '" +
            escapedPath +
            "'; " +
            "if ($s.Status -ne 'Valid' -or " +
            "$null -eq $s.SignerCertificate -or " +
            "$s.SignerCertificate.Subject " +
            "-notmatch 'Microsoft Corporation') " +
            "{ exit 1 }";

        string encodedCommand =
            Convert.ToBase64String(
                Encoding.Unicode.GetBytes(script)
            );

        using var process =
            Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",

                Arguments =
                    "-NoProfile -NonInteractive " +
                    "-EncodedCommand " +
                    encodedCommand,

                UseShellExecute = false,
                CreateNoWindow = true
            })
            ?? throw new InvalidOperationException(
                "Не удалось запустить проверку Authenticode."
            );

        await process.WaitForExitAsync();

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                "Файл не прошёл проверку " +
                "цифровой подписи Microsoft."
            );
        }
    }

    private string CreateConfiguration()
    {
        var product = new XElement(
            "Product",

            new XAttribute(
                "ID",
                "ProPlus2024Volume"
            ),

            new XElement(
                "Language",
                new XAttribute("ID", "ru-ru")
            )
        );

        foreach (string app in Apps)
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

        foreach (string app in new[]
        {
            "Lync",
            "Groove"
        })
        {
            product.Add(
                new XElement(
                    "ExcludeApp",
                    new XAttribute("ID", app)
                )
            );
        }

        int edition =
            SelectedArchitecture == "x86"
                ? 32
                : 64;

        var configuration = new XDocument(
            new XElement(
                "Configuration",

                new XElement(
                    "Add",

                    new XAttribute(
                        "OfficeClientEdition",
                        edition
                    ),

                    new XAttribute(
                        "Channel",
                        "PerpetualVL2024"
                    ),

                    product
                ),

                new XElement(
                    "Display",

                    new XAttribute(
                        "Level",
                        "Full"
                    ),

                    new XAttribute(
                        "AcceptEULA",
                        "FALSE"
                    )
                ),

                new XElement(
                    "Updates",

                    new XAttribute(
                        "Enabled",
                        "TRUE"
                    )
                )
            )
        );

        string folder = Path.Combine(
            Path.GetTempPath(),
            "OfficeInstaller",
            Guid.NewGuid().ToString("N")
        );

        Directory.CreateDirectory(folder);

        string configurationPath =
            Path.Combine(
                folder,
                "configuration.xml"
            );

        configuration.Save(configurationPath);

        return configurationPath;
    }
}