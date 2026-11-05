using System.Diagnostics;
using System.Drawing;

namespace ConciliadorBancario;

internal sealed class RequirementsForm : Form
{
    private readonly ListView _items = new();
    private readonly Label _header = new();
    private readonly Button _install = new();
    public RequirementsForm(bool startup)
    {
        Text = "Verificar requisitos — Conciliador V32";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(775, 430);
        MinimumSize = new Size(620, 370);
        Font = new Font("Segoe UI", 9.5F);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, Padding = new Padding(15) };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
        Controls.Add(layout);
        _header.Dock = DockStyle.Fill;
        _header.TextAlign = ContentAlignment.MiddleLeft;
        _header.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        layout.Controls.Add(_header, 0, 0);
        _items.Dock = DockStyle.Fill;
        _items.View = View.Details;
        _items.FullRowSelect = true;
        _items.GridLines = true;
        _items.Columns.Add("Estado", 105);
        _items.Columns.Add("Componente", 180);
        _items.Columns.Add("Descripción / ruta encontrada", 430);
        layout.Controls.Add(_items, 0, 1);
        var note = new Label { Text = "Los módulos sin OCR pueden utilizarse aunque falten Tesseract o Poppler. Instalar requiere Internet y WinGet.", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.FromArgb(75, 88, 105) };
        layout.Controls.Add(note, 0, 2);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var close = new Button { Text = "Continuar", Width = 105, Height = 30 };
        close.Click += (_, _) => Close();
        buttons.Controls.Add(close);
        var verify = new Button { Text = "Verificar nuevamente", Width = 155, Height = 30 };
        verify.Click += (_, _) => RefreshRequirements();
        buttons.Controls.Add(verify);
        _install.Text = "Instalar faltantes";
        _install.Width = 140;
        _install.Height = 30;
        _install.Click += (_, _) => InstallMissing();
        buttons.Controls.Add(_install);
        layout.Controls.Add(buttons, 0, 3);
        Shown += (_, _) => RefreshRequirements();
    }

    private void RefreshRequirements()
    {
        var state = RequirementsService.Check();
        _items.Items.Clear();
        foreach (var entry in state.Items)
        {
            var item = new ListViewItem(entry.Ready ? "✓ Instalado" : "⚠ Faltante");
            item.SubItems.Add(entry.Name);
            item.SubItems.Add(entry.Detail);
            item.ForeColor = entry.Ready ? Color.DarkGreen : Color.DarkRed;
            _items.Items.Add(item);
        }
        _header.Text = state.AllReady ? "Equipo preparado: todos los componentes detectados" : "Requisitos del equipo: revise los elementos pendientes";
        _install.Enabled = state.Items.Any(x => !x.Ready && (x.Name.Contains("Tesseract") || x.Name.Contains("Poppler") || x.Name.Contains("Idioma")));
    }

    private void InstallMissing()
    {
        if (MessageBox.Show(this, "Se instalarán los componentes OCR faltantes con WinGet. Se necesita conexión a Internet y podrían solicitarse permisos de Windows. ¿Continuar?", "Instalación asistida", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        try
        {
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            string script = Path.Combine(AppContext.BaseDirectory, "INSTALAR_REQUISITOS_V32.ps1");
            if (!File.Exists(script))
            {
                MessageBox.Show(this, "No se encontró INSTALAR_REQUISITOS_V32.ps1 junto al ejecutable. Vuelva a publicar el proyecto con PUBLICAR_EXE_WIN64.bat.");
                return;
            }
            Process.Start(new ProcessStartInfo(powershell) { Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\"", UseShellExecute = true });
            MessageBox.Show(this, "Se abrió el instalador. Una vez que termine, pulse 'Verificar nuevamente'.", "Instalación iniciada", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "No se pudo abrir el instalador", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
}
