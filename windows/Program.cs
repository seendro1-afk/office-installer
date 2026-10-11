using System.Diagnostics;
using System.Security.Principal;
using System.Xml.Linq;
using System.Drawing;
using System.Windows.Forms;

ApplicationConfiguration.Initialize();
Application.Run(new InstallerForm());

public sealed class InstallerForm : Form
{
    private readonly Dictionary<string, CheckBox> boxes = new();
    private readonly Label status = new() { AutoSize = true, Text = "Выберите приложения и укажите setup.exe из Office Deployment Tool." };
    private readonly Button install = new() { Text = "Установить выбранное", AutoSize = true };
    private readonly TextBox odtPath = new() { Width = 390, ReadOnly = true };
    private readonly string[] apps = { "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access" };

    public InstallerForm()
    {
        Text = "Office Installer — Office LTSC 2024";
        ClientSize = new Size(560, 475);
        MinimumSize = new Size(580, 500);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(18) };
        Controls.Add(layout);
        layout.Controls.Add(new Label { Text = "Office LTSC 2024 — выбор приложений", Font = new Font(Font, FontStyle.Bold), AutoSize = true });
        layout.Controls.Add(new Label { Text = "Требуется действующая корпоративная лицензия Office LTSC 2024.", AutoSize = true });
        foreach (var app in apps)
        {
            var box = new CheckBox { Text = app, Checked = app is "Word" or "Excel" or "PowerPoint" or "Outlook", AutoSize = true };
            boxes[app] = box;
            layout.Controls.Add(box);
        }
        var row = new FlowLayoutPanel { Width = 515, Height = 44 };
        var browse = new Button { Text = "Выбрать ODT setup.exe", AutoSize = true };
        browse.Click += (_, _) => { using var d = new OpenFileDialog { Filter = "setup.exe|setup.exe", Title = "Выберите setup.exe из официального Office Deployment Tool" }; if (d.ShowDialog() == DialogResult.OK) odtPath.Text = d.FileName; };
        row.Controls.Add(browse);
        layout.Controls.Add(row);
        layout.Controls.Add(odtPath);
        layout.Controls.Add(install);
        status.MaximumSize = new Size(490, 0);
        layout.Controls.Add(status);
        install.Click += async (_, _) => await InstallAsync();
    }

    private async Task InstallAsync()
    {
        if (!boxes.Values.Any(x => x.Checked)) { MessageBox.Show("Выберите хотя бы одно приложение."); return; }
        if (!File.Exists(odtPath.Text) || Path.GetFileName(odtPath.Text).ToLowerInvariant() != "setup.exe") { MessageBox.Show("Укажите setup.exe из Office Deployment Tool."); return; }
        using var identity = WindowsIdentity.GetCurrent();
        if (!new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator)) { MessageBox.Show("Запустите программу от имени администратора."); return; }
        install.Enabled = false;
        try
        {
            var product = new XElement("Product", new XAttribute("ID", "ProPlus2024Volume"), new XElement("Language", new XAttribute("ID", "ru-ru")));
            foreach (var app in apps.Where(a => !boxes[a].Checked)) product.Add(new XElement("ExcludeApp", new XAttribute("ID", app)));
            // Exclude optional apps not offered in this UI.
            foreach (var app in new[] { "Lync", "Groove" }) product.Add(new XElement("ExcludeApp", new XAttribute("ID", app)));
            var xml = new XDocument(new XElement("Configuration",
                new XElement("Add", new XAttribute("OfficeClientEdition", "64"), new XAttribute("Channel", "PerpetualVL2024"), product),
                new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "FALSE")),
                new XElement("Updates", new XAttribute("Enabled", "TRUE"))));
            var folder = Path.Combine(Path.GetTempPath(), "OfficeInstaller", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            var config = Path.Combine(folder, "configuration.xml");
            xml.Save(config);
            status.Text = "Идёт установка из Microsoft CDN. Дождитесь завершения окна установки Office…";
            using var process = Process.Start(new ProcessStartInfo(odtPath.Text, $"/configure \"{config}\"") { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(odtPath.Text)! }) ?? throw new InvalidOperationException("Не удалось запустить ODT.");
            await process.WaitForExitAsync();
            status.Text = process.ExitCode == 0 ? "ODT завершил работу. Проверьте приложения и активируйте Office по действующей лицензии." : $"ODT завершился с кодом {process.ExitCode}. Проверьте журналы установки Office.";
        }
        catch (Exception ex) { status.Text = "Ошибка: " + ex.Message; }
        finally { install.Enabled = true; }
    }
}
