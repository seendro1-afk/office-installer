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
    private readonly CheckBox replacementModeBox = new()
    {
        Text = "Замена существующей версии Office",
        AutoSize = true,
        Checked = false
    };

    private readonly CheckBox projectBox = new()
    {
        Text = "Project Professional",
        AutoSize = true,
        Checked = false
    };

    private readonly CheckBox visioBox = new()
    {
        Text = "Visio Professional",
        AutoSize = true,
        Checked = false
    };


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

    private readonly ComboBox officeVersionBox = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Width = 260,
        Font = new Font("Segoe UI", 10)
    };

    private int SelectedOfficeYear =>
        officeVersionBox.SelectedIndex == 1 ? 2021 : 2024;

    private string SelectedProductId =>
        SelectedOfficeYear == 2021
            ? "ProPlus2021Volume"
            : "ProPlus2024Volume";

    private string SelectedChannel =>
        SelectedOfficeYear == 2021
            ? "PerpetualVL2021"
            : "PerpetualVL2024";

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

    private readonly Button previewButton =
        CreateButton("Предпросмотр XML");

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

        // Дополнительные продукты Office LTSC

        var additionalCard =
            CreateCard("Дополнительные продукты");

        additionalCard.Controls.Add(projectBox);
        additionalCard.Controls.Add(visioBox);

        additionalCard.Controls.Add(new Label
        {
            Text = "Для Project и Visio требуются отдельные лицензии.",
            AutoSize = true,
            ForeColor = TextSecondary
        });

        layout.Controls.Add(additionalCard);

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

        // Выбор версии Office

        var versionCard = CreateCard("Версия Office");

        officeVersionBox.Items.AddRange(new object[]
        {
            "Office LTSC 2024",
            "Office LTSC 2021"
        });

        officeVersionBox.SelectedIndex = 0;

        officeVersionBox.SelectedIndexChanged += (_, _) =>
        {
            Text = $"Office Installer — Office LTSC {SelectedOfficeYear}";

            InstallerLogger.Write(
                $"Выбрана версия Office LTSC {SelectedOfficeYear}.");

            UpdateCompatibilityDisplay();
        };

        versionCard.Controls.Add(officeVersionBox);
        layout.Controls.Add(versionCard);

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
        {
            UpdateArchitectureHint();
            UpdateCompatibilityDisplay();
        };

        architectureCard.Controls.Add(architectureBox);

        architectureCard.Controls.Add(architectureHint);

        layout.Controls.Add(architectureCard);

        // Режим установки

        var installationModeCard = CreateCard("Режим установки");

        installationModeCard.Controls.Add(replacementModeBox);

        installationModeCard.Controls.Add(new Label
        {
            Text =
                "Режим замены пока находится в разработке. " +
                "Удаление существующего Office отключено.",
            AutoSize = true,
            ForeColor = TextSecondary,
            MaximumSize = new Size(500, 0)
        });

        replacementModeBox.CheckedChanged += (_, _) =>
        {
            UpdateCompatibilityDisplay();
        };

        var replacementPlanButton = new Button
        {
            Text = "Посмотреть план замены",
            Width = 240,
            Height = 38,
            AutoSize = false
        };

        replacementPlanButton.Click += (_, _) =>
        {
            try
            {
                RefreshInstalledApps();
                RefreshOfficeEdition();

                string plan = BuildReplacementPlan();

                if (replacementModeBox.Checked)
                {
                    try
                    {
                        ValidateReplacementMode();

                        string removalXml =
                            CreateRemovalConfiguration();

                        plan += Environment.NewLine +
                            Environment.NewLine +
                            "XML ДЛЯ УДАЛЕНИЯ:" +
                            Environment.NewLine +
                            removalXml;
                    }
                    catch (InvalidOperationException ex)
                    {
                        plan += Environment.NewLine +
                            Environment.NewLine +
                            "XML удаления не сформирован: " +
                            ex.Message;
                    }
                }

                InstallerLogger.Write(
                    "Сформирован предварительный план замены Office."
                );

                MessageBox.Show(
                    this,
                    plan,
                    "План замены Office",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                InstallerLogger.Error("BuildReplacementPlan", ex);

                MessageBox.Show(
                    this,
                    ex.Message,
                    "Ошибка формирования плана",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        };

        installationModeCard.Controls.Add(replacementPlanButton);

        layout.Controls.Add(installationModeCard);

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

        previewButton.Click += (_, _) => PreviewConfiguration();

        odtCard.Controls.Add(previewButton);



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
                SelectedOfficeYear,
                blockReason,
                applications
            );

            InstallerLogger.Write("Диагностика: отчёт сформирован.");

            string documents = Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments
            );

            string directory = Path.Combine(
                documents,
                "Office Installer",
                "Diagnostics"
            );

            Directory.CreateDirectory(directory);

            string filename =
                "Office-Diagnostics-" +
                DateTime.Now.ToString("yyyy-MM-dd-HHmmss-fff") +
                ".txt";

            string reportPath = Path.Combine(directory, filename);

            InstallerLogger.Write(
                "Диагностика: сохранение отчёта: " + reportPath
            );

            File.WriteAllText(
                reportPath,
                report,
                Encoding.UTF8
            );

            InstallerLogger.Write(
                "Диагностика: отчёт успешно сохранён."
            );

            status.Text = "Отчёт диагностики сохранён.";

            // Открываем папку с готовым отчётом.
            try
            {
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = directory,
                        UseShellExecute = true
                    }
                );

                InstallerLogger.Write(
                    "Диагностика: папка отчётов открыта."
                );
            }
            catch (Exception openException)
            {
                InstallerLogger.Error(
                    "Открытие папки диагностики",
                    openException
                );
            }

            MessageBox.Show(
                this,
                "Диагностический отчёт сохранён.\n\n" +
                reportPath,
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

        UpdateArchitectureHint();
    }

    private void UpdateArchitectureHint()
    {
        string installedArchitecture =
            detectedOffice?.Architecture ?? "";

        string selectedArchitecture = SelectedArchitecture;

        if (!installedArchitecture.Equals("x86",
                StringComparison.OrdinalIgnoreCase) &&
            !installedArchitecture.Equals("x64",
                StringComparison.OrdinalIgnoreCase))
        {
            architectureHint.Text =
                "Выбрана " + selectedArchitecture +
                ". Архитектура установленного Office не определена.";
            return;
        }

        if (installedArchitecture.Equals(
            selectedArchitecture,
            StringComparison.OrdinalIgnoreCase))
        {
            architectureHint.Text =
                "Обнаружен Office " + installedArchitecture +
                ". Архитектура совпадает.";
        }
        else
        {
            architectureHint.Text =
                "Обнаружен Office " + installedArchitecture +
                ". Выбрана " + selectedArchitecture +
                " — конфликт архитектур.";
        }
    }

    private string SelectedArchitecture =>
        architectureBox.SelectedIndex == 1
            ? "x86"
            : "x64";

    private void ValidateDownloadedOfficeFiles()
    {
        string downloadFolder = GetOfficeDownloadFolder();

        string dataFolder = Path.Combine(
            downloadFolder,
            "Office",
            "Data"
        );

        if (!Directory.Exists(dataFolder))
        {
            throw new InvalidOperationException(
                "Не найден каталог скачанного Office: " +
                dataFolder
            );
        }

        bool hasCabFiles = Directory
            .EnumerateFiles(
                dataFolder,
                "*.cab",
                SearchOption.AllDirectories
            )
            .Any();

        bool hasDatFiles = Directory
            .EnumerateFiles(
                dataFolder,
                "*.dat",
                SearchOption.AllDirectories
            )
            .Any();

        if (!hasCabFiles || !hasDatFiles)
        {
            throw new InvalidOperationException(
                "Дистрибутив Office загружен не полностью. " +
                "Не найдены необходимые CAB или DAT-файлы."
            );
        }

        InstallerLogger.Write(
            "Предварительная проверка файлов Office пройдена: " +
            dataFolder
        );
    }

    private void ValidateRemovalConfiguration(string removalXml)
    {
        // Повторно считываем установленную конфигурацию.
        RefreshOfficeEdition();
        ValidateReplacementMode();

        var document = XDocument.Parse(removalXml);

        var removeElement =
            document.Root?.Element("Remove")
            ?? throw new InvalidOperationException(
                "В XML отсутствует элемент Remove."
            );

        string[] xmlProducts = removeElement
            .Elements("Product")
            .Select(element =>
                (string?)element.Attribute("ID") ?? "")
            .ToArray();

        string[] currentProducts = GetProductIds();

        if (xmlProducts.Length != currentProducts.Length ||
            xmlProducts.Distinct(
                StringComparer.OrdinalIgnoreCase
            ).Count() != xmlProducts.Length ||
            currentProducts.Except(
                xmlProducts,
                StringComparer.OrdinalIgnoreCase
            ).Any())
        {
            throw new InvalidOperationException(
                "Список установленных продуктов Office изменился. " +
                "Операция удаления отменена."
            );
        }

        InstallerLogger.Write(
            "Список продуктов в XML удаления соответствует " +
            "текущей конфигурации Office."
        );
    }

    private bool ConfirmOfficeReplacement(string removalXml)
    {
        ValidateRemovalConfiguration(removalXml);

        var document = XDocument.Parse(removalXml);

        string[] productsToRemove = document.Root!
            .Element("Remove")!
            .Elements("Product")
            .Select(element =>
                (string?)element.Attribute("ID") ?? "")
            .ToArray();

        string products = string.Join(
            Environment.NewLine,
            productsToRemove.Select(id => "  - " + id)
        );

        string message =
            "ВНИМАНИЕ! ЗАМЕНА MICROSOFT OFFICE" +
            Environment.NewLine +
            Environment.NewLine +
            "Будут удалены следующие продукты:" +
            Environment.NewLine +
            products +
            Environment.NewLine +
            Environment.NewLine +
            $"После удаления будет запущена установка Office LTSC {SelectedOfficeYear}." +
            Environment.NewLine +
            Environment.NewLine +
            "Если установка новой версии завершится ошибкой, " +
            "Office может остаться удалённым." +
            Environment.NewLine +
            Environment.NewLine +
            "Автоматический откат не предусмотрен." +
            Environment.NewLine +
            Environment.NewLine +
            "Продолжить замену?";

        return MessageBox.Show(
            this,
            message,
            "Подтверждение удаления Office",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2
        ) == DialogResult.Yes;
    }

    private void ValidateReplacementMode()
    {
        if (!replacementModeBox.Checked)
        {
            throw new InvalidOperationException(
                "Режим замены Office не включён."
            );
        }

        if (SelectedOfficeYear != 2024)
        {
            throw new InvalidOperationException(
                "Сейчас поддерживается только переход " +
                "с Office LTSC 2021 на Office LTSC 2024."
            );
        }

        if (detectedOffice == null ||
            !detectedOffice.ClickToRunDetected)
        {
            throw new InvalidOperationException(
                "Не обнаружен Office Click-to-Run."
            );
        }

        string installedArchitecture =
            detectedOffice.Architecture;

        if (!installedArchitecture.Equals(
            SelectedArchitecture,
            StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Для первого этапа замены архитектура должна " +
                "совпадать с установленной. " +
                $"Установлено: {installedArchitecture}; " +
                $"выбрано: {SelectedArchitecture}."
            );
        }

        string[] ids = GetProductIds();

        if (!ids.Contains(
            "ProPlus2021Volume",
            StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Для замены требуется установленный " +
                "Office LTSC Professional Plus 2021."
            );
        }

        // Генератор XML также проверяет,
        // что неизвестных продуктов нет.
        _ = CreateRemovalConfiguration();
    }

    private string CreateRemovalConfiguration()
    {
        if (detectedOffice == null ||
            !detectedOffice.ClickToRunDetected)
        {
            throw new InvalidOperationException(
                "Не обнаружена установка Office Click-to-Run."
            );
        }

        string[] installedProducts = GetProductIds();

        string[] supportedProducts =
        {
            "ProPlus2021Volume",
            "Standard2021Volume",
            "ProjectPro2021Volume",
            "ProjectStd2021Volume",
            "VisioPro2021Volume",
            "VisioStd2021Volume"
        };

        string[] unsupportedProducts = installedProducts
            .Where(id => !supportedProducts.Contains(
                id,
                StringComparer.OrdinalIgnoreCase
            ))
            .ToArray();

        if (unsupportedProducts.Length > 0)
        {
            throw new InvalidOperationException(
                "Обнаружены продукты, для которых безопасная " +
                "замена пока не реализована: " +
                string.Join(", ", unsupportedProducts) +
                ". Удаление заблокировано."
            );
        }

        string[] productsToRemove = installedProducts
            .Where(id => supportedProducts.Contains(
                id,
                StringComparer.OrdinalIgnoreCase
            ))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (productsToRemove.Length == 0)
        {
            throw new InvalidOperationException(
                "Не обнаружены поддерживаемые продукты Office 2021 для удаления."
            );
        }

        var configuration = new XDocument(
            new XElement(
                "Configuration",
                new XElement(
                    "Remove",
                    productsToRemove.Select(id =>
                        new XElement(
                            "Product",
                            new XAttribute("ID", id)
                        )
                    )
                ),
                new XElement(
                    "Display",
                    new XAttribute("Level", "Full")
                )
            )
        );

        return configuration.ToString();
    }

    private string BuildReplacementPlan()
    {
        string[] installedProducts = GetProductIds();

        var selectedProducts = new List<string>
        {
            SelectedProductId
        };

        if (projectBox.Checked)
        {
            selectedProducts.Add(
                SelectedOfficeYear == 2021
                    ? "ProjectPro2021Volume"
                    : "ProjectPro2024Volume"
            );
        }

        if (visioBox.Checked)
        {
            selectedProducts.Add(
                SelectedOfficeYear == 2021
                    ? "VisioPro2021Volume"
                    : "VisioPro2024Volume"
            );
        }

        var lines = new List<string>
        {
            "ПЛАН ЗАМЕНЫ OFFICE",
            "",
            "Установленная архитектура: " +
                (detectedOffice?.Architecture ?? "Неизвестна"),
            "Выбранная архитектура: " +
                SelectedArchitecture,
            "",
            "Обнаруженные продукты:"
        };

        if (installedProducts.Length == 0)
        {
            lines.Add("  Не обнаружены");
        }
        else
        {
            foreach (string id in installedProducts)
            {
                lines.Add("  - " + id);
            }
        }

        lines.Add("");
        lines.Add("Выбраны для установки:");

        foreach (string id in selectedProducts)
        {
            lines.Add("  + " + id);
        }

        lines.Add("");
        lines.Add("Продукты, которые не будут переустановлены:");

        static string GetProductFamily(string id)
        {
            if (id.StartsWith("Project", StringComparison.OrdinalIgnoreCase))
                return "Project";

            if (id.StartsWith("Visio", StringComparison.OrdinalIgnoreCase))
                return "Visio";

            if (id.StartsWith("ProPlus", StringComparison.OrdinalIgnoreCase) ||
                id.StartsWith("Standard", StringComparison.OrdinalIgnoreCase))
                return "Office";

            return id;
        }

        var selectedFamilies = selectedProducts
            .Select(GetProductFamily)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var notReinstalled = installedProducts
            .Where(id => !selectedFamilies.Contains(GetProductFamily(id)))
            .ToArray();

        if (notReinstalled.Length == 0)
        {
            lines.Add("  Нет");
        }
        else
        {
            foreach (string id in notReinstalled)
            {
                lines.Add("  ! " + id);
            }
        }

        lines.Add("");
        lines.Add(
            "ВНИМАНИЕ: это предварительный план. " +
            "Удаление продуктов пока не выполняется."
        );

        return string.Join(Environment.NewLine, lines);
    }

    private string? GetInstallationBlockReason()
    {
        if (replacementModeBox.Checked)
        {
            try
            {
                ValidateReplacementMode();
                return null;
            }
            catch (Exception ex)
            {
                return "Замена Office заблокирована: " +
                    ex.Message;
            }
        }

        if (detectedOffice == null)
        {
            return "Не удалось проверить установленный Office.";
        }

        string existingArchitecture = detectedOffice.Architecture;

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
                $"Несовместимая архитектура: установлен {existingArchitecture}, " +
                $"а выбран {SelectedArchitecture}. " +
                "Для продуктов Click-to-Run архитектуры должны совпадать.";
        }

        if (!detectedOffice.ClickToRunDetected)
        {
            return officeExecutablesDetected
                ? "Обнаружены приложения Office, " +
                  "но конфигурация Click-to-Run неизвестна."
                : null;
        }

        string[] ids = GetProductIds();

        if (ids.Length == 0)
        {
            return
                "Обнаружен Office Click-to-Run, " +
                "но идентификаторы продуктов определить не удалось.";
        }

        // Повторную установку существующего Office
        // пока не разрешаем без отдельной проверки.
        if (ids.Any(id => id.Equals(
            SelectedProductId,
            StringComparison.OrdinalIgnoreCase)))
        {
            return
                $"Office LTSC {SelectedOfficeYear} уже обнаружен. " +
                "Повторная установка поверх существующей " +
                "конфигурации пока заблокирована.";
        }

        var incompatibleProducts = new List<string>();

        foreach (string id in ids)
        {
            if (id.Equals(
                SelectedProductId,
                StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string productName = id switch
            {
                "ProPlus2021Volume" =>
                    "Office LTSC Professional Plus 2021",

                "Standard2021Volume" =>
                    "Office LTSC Standard 2021",

                "ProjectPro2021Volume" =>
                    "Project Professional 2021",

                "ProjectStd2021Volume" =>
                    "Project Standard 2021",

                "VisioPro2021Volume" =>
                    "Visio Professional 2021",

                "VisioStd2021Volume" =>
                    "Visio Standard 2021",

                "ProPlus2019Volume" =>
                    "Office Professional Plus 2019",

                "Standard2019Volume" =>
                    "Office Standard 2019",

                "ProjectPro2019Volume" =>
                    "Project Professional 2019",

                "VisioPro2019Volume" =>
                    "Visio Professional 2019",

                _ => id
            };

            incompatibleProducts.Add(productName);
        }

        if (incompatibleProducts.Count > 0)
        {
            string products = string.Join(
                "; ",
                incompatibleProducts
            );

            return
                "Обнаружены продукты, совместимость которых " +
                $"с установкой Office LTSC {SelectedOfficeYear} не подтверждена: " +
                products + ". " +
                $"Установленная архитектура: {existingArchitecture}. " +
                $"Выбранная архитектура: {SelectedArchitecture}.";
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

    private void SetInstallationStage(
        int stage,
        string description)
    {
        string message =
            $"Этап {stage}/5 — {description}";

        status.Text = message;
        InstallerLogger.Write(message);
    }

    private void SetBusy(bool busy)
    {
        isBusy = busy;

        checkOdt.Enabled = !busy;
        previewButton.Enabled = !busy;
        browse.Enabled = !busy;
        refresh.Enabled = !busy;
        architectureBox.Enabled = !busy;
        officeVersionBox.Enabled = !busy;
        replacementModeBox.Enabled = !busy;
        projectBox.Enabled = !busy;
        visioBox.Enabled = !busy;

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

        if (replacementModeBox.Checked)
        {
            MessageBox.Show(
                this,
                "Режим замены Office пока находится в разработке.\n\n" +
                "Удаление существующих продуктов отключено. " +
                "Используйте кнопку «Посмотреть план замены».",
                "Замена Office пока недоступна",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );

            return;
        }

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
            $"Редакция: Office LTSC Professional Plus {SelectedOfficeYear}\n\n" +
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
            SetInstallationStage(
                1,
                "Подготовка Office Deployment Tool");

            string setupPath =
                await PrepareOdtAsync();

            odtPath.Text = setupPath;

            SetInstallationStage(
                2,
                $"Создание конфигурации Office LTSC {SelectedOfficeYear}");

            string configurationPath =
                CreateConfiguration();

            InstallerLogger.Write(
                $"Запуск установки Office LTSC {SelectedOfficeYear}. " +
                $"Архитектура={SelectedArchitecture}; " +
                $"Приложения={selectedApps}; " +
                $"Конфигурация={configurationPath}.");

            CheckDownloadDiskSpace();

            SetInstallationStage(
                3,
                $"Загрузка Office LTSC {SelectedOfficeYear}");

            InstallerLogger.Write(
                "Начало загрузки дистрибутива Office.");

            using (var downloadProcess =
                Process.Start(new ProcessStartInfo
                {
                    FileName = setupPath,

                    Arguments =
                        $"/download \"{configurationPath}\"",

                    UseShellExecute = false,

                    WorkingDirectory =
                        Path.GetDirectoryName(setupPath)!
                })
                ?? throw new InvalidOperationException(
                    "Не удалось запустить загрузку Office."
                ))
            {
                await downloadProcess.WaitForExitAsync();

                InstallerLogger.Write(
                    "Загрузка Office завершилась. Код: " +
                    downloadProcess.ExitCode);

                if (downloadProcess.ExitCode != 0)
                {
                    throw new InvalidOperationException(
                        "Не удалось загрузить файлы Office. " +
                        $"Код ODT: {downloadProcess.ExitCode}."
                    );
                }
            }

            InstallerLogger.Write(
                "Проверка локальных файлов Office после загрузки."
            );

            ValidateDownloadedOfficeFiles();

            SetInstallationStage(
                4,
                $"Установка Office LTSC {SelectedOfficeYear}");

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

            SetInstallationStage(
                5,
                "Проверка установленного Office");

            await VerifyOfficeInstallationWithRetryAsync();
            RefreshOfficeInformation();

            status.Text =
                $"Office LTSC {SelectedOfficeYear} установлен. " +
                "Проверьте лицензирование.";

            InstallerLogger.Write(
                $"Все этапы установки Office LTSC {SelectedOfficeYear} завершены.");

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

    private async Task VerifyOfficeInstallationWithRetryAsync()
    {
        InvalidOperationException? lastError = null;

        for (int attempt = 1; attempt <= 6; attempt++)
        {
            try
            {
                VerifyOfficeInstallation();
                return;
            }
            catch (InvalidOperationException ex)
            {
                lastError = ex;

                InstallerLogger.Write(
                    $"Проверка Office: попытка {attempt}/6. " +
                    ex.Message
                );

                if (attempt < 6)
                {
                    status.Text =
                        $"Проверка установки Office: {attempt}/6...";

                    await Task.Delay(TimeSpan.FromSeconds(5));
                }
            }
        }

        throw new InvalidOperationException(
            "Не удалось подтвердить установку Office " +
            "после нескольких проверок. " +
            lastError?.Message
        );
    }

    private void VerifyOfficeInstallation()
    {
        OfficeDetectionInfo installed =
            OfficeDetectionService.Detect();

        if (!installed.ClickToRunDetected)
        {
            throw new InvalidOperationException(
                "После установки Office Click-to-Run не обнаружен."
            );
        }

        string[] productIds = installed.ProductIds
            .Split(
                ',',
                StringSplitOptions.RemoveEmptyEntries |
                StringSplitOptions.TrimEntries
            );

        if (!productIds.Contains(
            SelectedProductId,
            StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Не обнаружен продукт {SelectedProductId}. " +
                $"Найдены: {installed.ProductIds}."
            );
        }

        if (projectBox.Checked)
        {
            string expectedProject =
                SelectedOfficeYear == 2021
                    ? "ProjectPro2021Volume"
                    : "ProjectPro2024Volume";

            if (!productIds.Contains(
                expectedProject,
                StringComparer.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "После установки не найден Project: " +
                    expectedProject
                );
            }
        }

        if (visioBox.Checked)
        {
            string expectedVisio =
                SelectedOfficeYear == 2021
                    ? "VisioPro2021Volume"
                    : "VisioPro2024Volume";

            if (!productIds.Contains(
                expectedVisio,
                StringComparer.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "После установки не найден Visio: " +
                    expectedVisio
                );
            }
        }

        if (!string.Equals(
            installed.Architecture,
            SelectedArchitecture,
            StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Архитектура Office не совпадает. " +
                $"Ожидалась: {SelectedArchitecture}; " +
                $"обнаружена: {installed.Architecture}."
            );
        }

        foreach (string app in Apps)
        {
            if (!boxes[app].Checked)
            {
                continue;
            }

            string? executable =
                FindOfficeExecutable(Executables[app]);

            if (executable == null)
            {
                throw new InvalidOperationException(
                    $"После установки не найдено приложение {app}."
                );
            }
        }

        InstallerLogger.Write(
            $"Проверка Office LTSC {SelectedOfficeYear} успешна. " +
            $"Продукт={SelectedProductId}; " +
            $"Архитектура={SelectedArchitecture}."
        );
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

    private void PreviewConfiguration()
    {
        try
        {
            string configurationPath =
                CreateConfiguration();

            string xml =
                File.ReadAllText(configurationPath);

            string selectedApps = string.Join(
                ", ",
                Apps.Where(app => boxes[app].Checked)
            );

            if (string.IsNullOrWhiteSpace(selectedApps))
            {
                selectedApps = "Не выбраны";
            }

            InstallerLogger.Write(
                $"Открыт предпросмотр Office LTSC {SelectedOfficeYear}. " +
                $"Архитектура={SelectedArchitecture}; " +
                $"Приложения={selectedApps}.");

            using var previewForm = new Form
            {
                Text =
                    $"Предпросмотр — Office LTSC {SelectedOfficeYear}",

                StartPosition =
                    FormStartPosition.CenterParent,

                Size = new Size(850, 680),

                MinimumSize = new Size(650, 480),

                Font = new Font("Segoe UI", 10)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(16)
            };

            layout.RowStyles.Add(
                new RowStyle(SizeType.AutoSize)
            );

            layout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100)
            );

            layout.RowStyles.Add(
                new RowStyle(SizeType.AutoSize)
            );

            var summary = new Label
            {
                AutoSize = true,

                MaximumSize = new Size(780, 0),

                Text =
                    $"Версия: Office LTSC {SelectedOfficeYear}\n" +
                    $"Архитектура: {SelectedArchitecture}\n" +
                    $"Приложения: {selectedApps}\n\n" +
                    "Предпросмотр не запускает установку."
            };

            var xmlBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10),
                Text = xml
            };

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };

            var copyButton = new Button
            {
                Text = "Скопировать XML",
                AutoSize = true
            };

            var saveButton = new Button
            {
                Text = "Сохранить XML",
                AutoSize = true
            };

            var validateButton = new Button
            {
                Text = "Проверить XML",
                AutoSize = true
            };

            var closeButton = new Button
            {
                Text = "Закрыть",
                AutoSize = true,
                DialogResult = DialogResult.OK
            };

            copyButton.Click += (_, _) =>
            {
                try
                {
                    string textToCopy = xmlBox.Text;
                    Exception? clipboardError = null;

                    using var completed =
                        new System.Threading.ManualResetEventSlim(false);

                    var thread = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            Clipboard.SetText(textToCopy);
                        }
                        catch (Exception ex)
                        {
                            clipboardError = ex;
                        }
                        finally
                        {
                            completed.Set();
                        }
                    });

                    thread.SetApartmentState(
                        System.Threading.ApartmentState.STA
                    );

                    thread.Start();

                    if (!completed.Wait(
                        TimeSpan.FromSeconds(10)))
                    {
                        throw new TimeoutException(
                            "Превышено время ожидания буфера обмена."
                        );
                    }

                    if (clipboardError != null)
                    {
                        throw new InvalidOperationException(
                            "Не удалось скопировать XML.",
                            clipboardError
                        );
                    }

                    InstallerLogger.Write(
                        "XML-конфигурация скопирована.");

                    MessageBox.Show(
                        previewForm,
                        "XML скопирован в буфер обмена.",
                        "Готово",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                catch (Exception ex)
                {
                    InstallerLogger.Error("CopyXml", ex);

                    MessageBox.Show(
                        previewForm,
                        ex.Message,
                        "Ошибка копирования",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            };

            saveButton.Click += (_, _) =>
            {
                try
                {
                    string documents =
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments
                        );

                    string folder = Path.Combine(
                        documents,
                        "Office Installer",
                        "Configurations"
                    );

                    Directory.CreateDirectory(folder);

                    string fileName =
                        $"Office-LTSC-{SelectedOfficeYear}-" +
                        $"{SelectedArchitecture}-" +
                        $"{DateTime.Now:yyyyMMdd-HHmmss-fff}-" +
                        $"{Guid.NewGuid():N}.xml";

                    string filePath =
                        Path.Combine(folder, fileName);

                    File.WriteAllText(
                        filePath,
                        xmlBox.Text,
                        new System.Text.UTF8Encoding(false)
                    );

                    InstallerLogger.Write(
                        "XML-конфигурация сохранена: " +
                        filePath
                    );

                    MessageBox.Show(
                        previewForm,
                        "Конфигурация сохранена:\n\n" +
                        filePath,
                        "Готово",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = folder,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    InstallerLogger.Error("SaveXml", ex);

                    MessageBox.Show(
                        previewForm,
                        ex.Message,
                        "Ошибка сохранения",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            };

            validateButton.Click += (_, _) =>
            {
                try
                {
                    ValidateConfigurationXml(xmlBox.Text);

                    InstallerLogger.Write(
                        "Проверка XML-конфигурации пройдена.");

                    MessageBox.Show(
                        previewForm,
                        "XML корректен.\n\n" +
                        "Версия Office, архитектура, " +
                        "канал и приложения соответствуют " +
                        "выбранным параметрам.\n\n" +
                        "Проверка не гарантирует успешную " +
                        "установку или активацию Office.",
                        "Проверка завершена",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                catch (Exception ex)
                {
                    InstallerLogger.Error(
                        "ValidateConfigurationXml",
                        ex
                    );

                    MessageBox.Show(
                        previewForm,
                        ex.Message,
                        "Ошибка проверки XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            };

            buttons.Controls.Add(copyButton);
            buttons.Controls.Add(saveButton);
            buttons.Controls.Add(validateButton);
            buttons.Controls.Add(closeButton);

            layout.Controls.Add(summary, 0, 0);
            layout.Controls.Add(xmlBox, 0, 1);
            layout.Controls.Add(buttons, 0, 2);

            previewForm.Controls.Add(layout);
            previewForm.AcceptButton = closeButton;

            previewForm.ShowDialog(this);
        }
        catch (Exception ex)
        {
            InstallerLogger.Error(
                "PreviewConfiguration",
                ex
            );

            MessageBox.Show(
                ex.Message,
                "Ошибка предпросмотра",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private void ValidateConfigurationXml(string xml)
    {
        var document = XDocument.Parse(xml);

        XElement root = document.Root
            ?? throw new InvalidOperationException(
                "Отсутствует корневой элемент XML."
            );

        if (root.Name != "Configuration")
        {
            throw new InvalidOperationException(
                "Корневой элемент должен называться Configuration."
            );
        }

        XElement add = root.Element("Add")
            ?? throw new InvalidOperationException(
                "Отсутствует элемент Add."
            );

        string expectedEdition =
            SelectedArchitecture == "x86" ? "32" : "64";

        string? actualEdition =
            (string?)add.Attribute("OfficeClientEdition");

        if (actualEdition != expectedEdition)
        {
            throw new InvalidOperationException(
                "Архитектура XML не соответствует выбранной."
            );
        }

        string? actualChannel =
            (string?)add.Attribute("Channel");

        if (actualChannel != SelectedChannel)
        {
            throw new InvalidOperationException(
                "Канал Office не соответствует выбранной версии."
            );
        }

        var expectedProducts = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase
        )
        {
            SelectedProductId
        };

        if (projectBox.Checked)
        {
            expectedProducts.Add(
                SelectedOfficeYear == 2021
                    ? "ProjectPro2021Volume"
                    : "ProjectPro2024Volume"
            );
        }

        if (visioBox.Checked)
        {
            expectedProducts.Add(
                SelectedOfficeYear == 2021
                    ? "VisioPro2021Volume"
                    : "VisioPro2024Volume"
            );
        }

        var products = add.Elements("Product").ToList();

        var actualProducts = products
            .Select(element => (string?)element.Attribute("ID"))
            .ToList();

        if (actualProducts.Any(string.IsNullOrWhiteSpace))
        {
            throw new InvalidOperationException(
                "У одного из продуктов отсутствует Product ID."
            );
        }

        var actualProductIds = actualProducts
            .Select(id => id!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (actualProductIds.Count != products.Count)
        {
            throw new InvalidOperationException(
                "В XML обнаружены повторяющиеся Product ID."
            );
        }

        if (!actualProductIds.SetEquals(expectedProducts))
        {
            throw new InvalidOperationException(
                "Список продуктов XML не соответствует выбранным. " +
                "Ожидаются: " +
                string.Join(", ", expectedProducts) +
                ". Обнаружены: " +
                string.Join(", ", actualProductIds)
            );
        }

        foreach (XElement item in products)
        {
            XElement language = item.Element("Language")
                ?? throw new InvalidOperationException(
                    "У продукта отсутствует элемент Language."
                );

            if ((string?)language.Attribute("ID") != "ru-ru")
            {
                throw new InvalidOperationException(
                    "Язык всех продуктов должен быть ru-ru."
                );
            }
        }

        XElement product = products.Single(element =>
            string.Equals(
                (string?)element.Attribute("ID"),
                SelectedProductId,
                StringComparison.OrdinalIgnoreCase
            )
        );

        var excludedApps = product
            .Elements("ExcludeApp")
            .Select(element =>
                (string?)element.Attribute("ID"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (string app in Apps)
        {
            bool shouldBeExcluded = !boxes[app].Checked;
            bool isExcluded = excludedApps.Contains(app);

            if (shouldBeExcluded != isExcluded)
            {
                throw new InvalidOperationException(
                    $"Параметры приложения {app} не совпадают."
                );
            }
        }

        foreach (string app in new[] { "Lync", "Groove" })
        {
            if (!excludedApps.Contains(app))
            {
                throw new InvalidOperationException(
                    $"Отсутствует исключение {app}."
                );
            }
        }
    }

    private void CheckDownloadDiskSpace()
    {
        string folder = GetOfficeDownloadFolder();

        string root = Path.GetPathRoot(folder)
            ?? throw new InvalidOperationException(
                "Не удалось определить диск загрузки Office."
            );

        var drive = new DriveInfo(root);

        const long requiredBytes = 10L * 1024 * 1024 * 1024;

        long availableBytes = drive.AvailableFreeSpace;

        InstallerLogger.Write(
            $"Свободное место на {root}: " +
            $"{availableBytes / (1024 * 1024 * 1024)} ГБ.");

        if (availableBytes < requiredBytes)
        {
            throw new InvalidOperationException(
                "Недостаточно свободного места для загрузки Office.\\n\\n" +
                $"Диск: {root}\\n" +
                $"Доступно: {availableBytes / (1024 * 1024 * 1024)} ГБ\\n" +
                "Рекомендуется освободить минимум 10 ГБ."
            );
        }
    }

    private string GetOfficeDownloadFolder()
    {
        string programData =
            Environment.GetFolderPath(
                Environment.SpecialFolder.CommonApplicationData
            );

        return Path.Combine(
            programData,
            "OfficeInstaller",
            "Downloads",
            $"Office{SelectedOfficeYear}",
            SelectedArchitecture
        );
    }

    private string CreateConfiguration()
    {
        var product = new XElement(
            "Product",

            new XAttribute(
                "ID",
                SelectedProductId
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

        var products = new List<XElement>
        {
            product
        };

        if (projectBox.Checked)
        {
            products.Add(
                new XElement(
                    "Product",
                    new XAttribute(
                        "ID",
                        SelectedOfficeYear == 2021
                            ? "ProjectPro2021Volume"
                            : "ProjectPro2024Volume"
                    ),
                    new XElement(
                        "Language",
                        new XAttribute("ID", "ru-ru")
                    )
                )
            );
        }

        if (visioBox.Checked)
        {
            products.Add(
                new XElement(
                    "Product",
                    new XAttribute(
                        "ID",
                        SelectedOfficeYear == 2021
                            ? "VisioPro2021Volume"
                            : "VisioPro2024Volume"
                    ),
                    new XElement(
                        "Language",
                        new XAttribute("ID", "ru-ru")
                    )
                )
            );
        }

        int edition =
            SelectedArchitecture == "x86"
                ? 32
                : 64;

        string downloadFolder =
            GetOfficeDownloadFolder();

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
                        SelectedChannel
                    ),

                    new XAttribute(
                        "SourcePath",
                        downloadFolder
                    ),

                    products
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