using System.Diagnostics;
using System.Drawing;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace ConciliadorBancario;

public sealed class MainForm : Form
{
    private string _archiveFile = "";
    private string _archiveFile2 = "";
    private string _archiveFile3 = "";
    private string _archiveCts = "";
    private string _archiveGratificacion = "";
    private string _archiveDescanso = "";
    private string _archiveLicencia = "";
    private string _archiveSubsidios = "";
    private string _archiveVacaciones = "";
    private string _payflowFile = "";
    private string _ctrFile = "";
    private string _vidaLeyFile = "";
    private string _boletasFile = "";
    private string _excelFile = "";
    private string _destination = "";
    private CancellationTokenSource? _cts;

    private readonly TextBox _archiveBox = new();
    private readonly TextBox _archiveBox2 = new();
    private readonly TextBox _archiveBox3 = new();
    private readonly TextBox _archiveCtsBox = new();
    private readonly TextBox _archiveGratificacionBox = new();
    private readonly TextBox _archiveDescansoBox = new();
    private readonly TextBox _archiveLicenciaBox = new();
    private readonly TextBox _archiveSubsidiosBox = new();
    private readonly TextBox _archiveVacacionesBox = new();
    private readonly Dictionary<TextBox, (Label Badge, bool Required, bool Folder)> _indicators = new();
    private readonly TextBox _payflowBox = new();
    private readonly TextBox _ctrBox = new();
    private readonly TextBox _vidaLeyBox = new();
    private readonly TextBox _boletasBox = new();
    private readonly TextBox _excelBox = new();
    private readonly TextBox _destinationBox = new();
    private readonly RichTextBox _logBox = new();
    private readonly Label _counter = new();
    private readonly Label _status = new();
    private readonly Button _runButton = new();
    private readonly Button _cancelButton = new();
    private readonly Button _requirementsButton = new();
    private readonly Button _orderButton = new();
    private IReadOnlyList<string> _documentOrder = DocumentOrderSettings.Load();
    private readonly ProgressBar _progress = new();

    private static readonly Color Navy = Color.FromArgb(17, 32, 55);
    private static readonly Color Blue = Color.FromArgb(29, 89, 174);
    private static readonly Color Pale = Color.FromArgb(244, 247, 252);
    private static readonly Color Gray = Color.FromArgb(86, 103, 125);

    public MainForm()
    {
        Text = "Conciliador bancario V35 | Dos órdenes configurables";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(980, 720);
        MinimumSize = new Size(760, 530);
        Font = new Font("Segoe UI", 9.3F);
        BackColor = Pale;
        AutoScaleMode = AutoScaleMode.Dpi;
        BuildLayout();
        UpdateState();
        Shown += (_, _) => BeginInvoke(new Action(() => { if (!RequirementsService.Check().AllReady) ShowRequirements(true); }));
    }

    private void ShowRequirements(bool startup)
    {
        using var dialog = new RequirementsForm(startup);
        dialog.ShowDialog(this);
    }

    // Diseño compacto: encabezado y acciones siempre visibles, contenido central desplazable.
    // No se usan posiciones absolutas para los selectores, evitando recortes al ampliar DPI.
    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4,
            Margin = Padding.Empty, Padding = Padding.Empty, BackColor = Pale
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 94));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 176)); // Excel y destino siempre visibles
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        Controls.Add(root);

        var header = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2,
            BackColor = Navy, Padding = new Padding(24, 15, 16, 12)
        };
        header.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
        header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        header.Controls.Add(new Label
        {
            Text = "Conciliador documental", Dock = DockStyle.Fill,
            ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 20F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        header.Controls.Add(new Label
        {
            Text = "Pagos · CTS y gratificación · PAYFLOW · CTR · VIDA LEY · Descansos y licencias · Boletas · Datos y salida",
            Dock = DockStyle.Fill, ForeColor = Color.FromArgb(207, 221, 240),
            Font = new Font("Segoe UI", 9F), AutoEllipsis = true,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 1);
        root.Controls.Add(header, 0, 0);

        // Datos obligatorios visibles en todo momento: nunca ocultos por el scroll.
        var mandatory = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3,
            Margin = Padding.Empty, Padding = new Padding(16, 3, 16, 0), BackColor = Pale
        };
        mandatory.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        mandatory.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        mandatory.RowStyles.Add(new RowStyle(SizeType.Absolute, 68));
        mandatory.RowStyles.Add(new RowStyle(SizeType.Absolute, 68));
        mandatory.Controls.Add(MakeSectionTitle("DATOS Y SALIDA", "Obligatorios · Lista Excel y carpeta de resultados"), 0, 0);
        mandatory.Controls.Add(MakePathRow("LISTA DE TRABAJADORES", "Requerido para pagos y seguros · Opcional para BOLETAS", _excelBox, "Elegir Excel", SelectExcel, required: false), 0, 1);
        mandatory.Controls.Add(MakePathRow("DESTINO", "Requerido · carpeta de resultados", _destinationBox, "Elegir carpeta", SelectDestination, required: true, folder: true), 0, 2);
        root.Controls.Add(mandatory, 0, 1);

        var scroll = new Panel
        {
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(16, 12, 16, 8),
            BackColor = Pale
        };
        root.Controls.Add(scroll, 0, 2);
        var body = new TableLayoutPanel
        {
            ColumnCount = 1, RowCount = 21, Dock = DockStyle.Top,
            AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = Padding.Empty, Padding = Padding.Empty, BackColor = Pale
        };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        int[] rowHeights = { 32, 68, 68, 68, 32, 68, 68, 32, 68, 32, 68, 32, 68, 32, 68, 68, 68, 68, 32, 68, 112 };
        foreach (int h in rowHeights) body.RowStyles.Add(new RowStyle(SizeType.Absolute, h));
        scroll.Controls.Add(body);
        scroll.Resize += (_, _) => body.Width = Math.Max(500, scroll.ClientSize.Width - scroll.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);
        body.Width = Math.Max(500, scroll.ClientSize.Width - scroll.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);

        body.Controls.Add(MakeSectionTitle("01   PAGOS", "PAGO 1, PAGO 2 y PAGO 3 · ZIP/RAR"), 0, 0);
        body.Controls.Add(MakePathRow("PAGO 1", "ZIP/RAR · opcional si procesa otros documentos", _archiveBox, "Elegir archivo", () => SelectArchive(1), () => ClearPayment(1)), 0, 1);
        body.Controls.Add(MakePathRow("PAGO 2", "Opcional · ZIP/RAR", _archiveBox2, "Elegir archivo", () => SelectArchive(2), () => ClearPayment(2)), 0, 2);
        body.Controls.Add(MakePathRow("PAGO 3", "Opcional · ZIP/RAR", _archiveBox3, "Elegir archivo", () => SelectArchive(3), () => ClearPayment(3)), 0, 3);

        body.Controls.Add(MakeSectionTitle("02   CTS Y GRATIFICACIÓN", "Comprobantes adicionales · ZIP/RAR"), 0, 4);
        body.Controls.Add(MakePathRow("PAGO CTS", "Opcional · ZIP/RAR", _archiveCtsBox, "Elegir archivo", () => SelectArchive(4), () => ClearPayment(4)), 0, 5);
        body.Controls.Add(MakePathRow("GRATIFICACIÓN", "Opcional · ZIP/RAR", _archiveGratificacionBox, "Elegir archivo", () => SelectArchive(5), () => ClearPayment(5)), 0, 6);

        body.Controls.Add(MakeSectionTitle("03   PAYFLOW", "Primera hoja del Excel · página individual por trabajador"), 0, 7);
        body.Controls.Add(MakePathRow("PAYFLOW", "Opcional · XLSX/XLSM", _payflowBox, "Elegir Excel", SelectPayflow, ClearPayflow), 0, 8);

        body.Controls.Add(MakeSectionTitle("04   CTR / SCTR", "Constancia PDF completa · solo enmarcar coincidencias"), 0, 9);
        body.Controls.Add(MakePathRow("CTR", "Opcional · PDF · no se divide en individuales", _ctrBox, "Elegir PDF", SelectCtr, ClearCtr), 0, 10);

        body.Controls.Add(MakeSectionTitle("05   VIDA LEY", "Constancia PDF completa · solo enmarcar coincidencias"), 0, 11);
        body.Controls.Add(MakePathRow("VIDA LEY", "Opcional · PDF · no se divide en individuales", _vidaLeyBox, "Elegir PDF", SelectVidaLey, ClearVidaLey), 0, 12);

        body.Controls.Add(MakeSectionTitle("06   DESCANSOS, LICENCIAS Y BENEFICIOS", "Procesamiento de comprobantes · ZIP/RAR · opcionales"), 0, 13);
        body.Controls.Add(MakePathRow("DESCANSO MÉDICO", "Opcional · ZIP/RAR", _archiveDescansoBox, "Elegir archivo", () => SelectArchive(6), () => ClearPayment(6)), 0, 14);
        body.Controls.Add(MakePathRow("LICENCIA", "Opcional · ZIP/RAR", _archiveLicenciaBox, "Elegir archivo", () => SelectArchive(7), () => ClearPayment(7)), 0, 15);
        body.Controls.Add(MakePathRow("SUBSIDIOS", "Opcional · ZIP/RAR", _archiveSubsidiosBox, "Elegir archivo", () => SelectArchive(8), () => ClearPayment(8)), 0, 16);
        body.Controls.Add(MakePathRow("VACACIONES", "Opcional · ZIP/RAR", _archiveVacacionesBox, "Elegir archivo", () => SelectArchive(9), () => ClearPayment(9)), 0, 17);

        body.Controls.Add(MakeSectionTitle("07   BOLETAS DE PAGO", "PDF único escaneado · separar cada página · DNI_NOMBRE.pdf"), 0, 18);
        body.Controls.Add(MakePathRow("BOLETAS", "Opcional · un solo PDF multipágina", _boletasBox, "Elegir archivo", SelectBoletas, ClearBoletas), 0, 19);

        var logPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2,
            Margin = new Padding(0, 7, 0, 4), Padding = new Padding(8, 4, 8, 6),
            BackColor = Color.White
        };
        logPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 23));
        logPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        logPanel.Controls.Add(new Label
        {
            Text = "Registro de actividad", Dock = DockStyle.Fill, ForeColor = Navy,
            Font = new Font("Segoe UI Semibold", 9.5F), TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        _logBox.Dock = DockStyle.Fill;
        _logBox.ReadOnly = true;
        _logBox.Font = new Font("Consolas", 8.8F);
        _logBox.BackColor = Color.FromArgb(249, 251, 254);
        _logBox.ForeColor = Navy;
        _logBox.BorderStyle = BorderStyle.None;
        _logBox.Text = "Seleccione uno o varios procesos, el Excel de trabajadores y la carpeta de salida.\n";
        logPanel.Controls.Add(_logBox, 0, 1);
        body.Controls.Add(logPanel, 0, 20);

        var footer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3,
            Padding = new Padding(20, 7, 20, 9), BackColor = Color.White
        };
        footer.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        footer.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));
        footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.Controls.Add(footer, 0, 3);
        var summary = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
        summary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 64));
        summary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36));
        _counter.Dock = DockStyle.Fill;
        _counter.Font = new Font("Segoe UI", 8.7F);
        _counter.ForeColor = Gray;
        _counter.AutoEllipsis = true;
        _counter.TextAlign = ContentAlignment.MiddleLeft;
        _status.Dock = DockStyle.Fill;
        _status.Font = new Font("Segoe UI", 8.7F);
        _status.TextAlign = ContentAlignment.MiddleRight;
        _status.AutoEllipsis = true;
        _status.ForeColor = Gray;
        summary.Controls.Add(_counter, 0, 0);
        summary.Controls.Add(_status, 1, 0);
        footer.Controls.Add(summary, 0, 0);
        _progress.Dock = DockStyle.Fill;
        _progress.Style = ProgressBarStyle.Continuous;
        _progress.Margin = Padding.Empty;
        footer.Controls.Add(_progress, 0, 1);
        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false, Margin = Padding.Empty, Padding = new Padding(0, 6, 0, 0)
        };
        StyleButton(_runButton, "GENERAR CONCILIACIÓN", true);
        _runButton.Width = 228;
        StyleButton(_cancelButton, "Cancelar", false);
        _cancelButton.Width = 112;
        actions.Controls.Add(_runButton);
        actions.Controls.Add(_cancelButton);
        StyleButton(_requirementsButton, "Verificar requisitos", false);
        _requirementsButton.Width = 153;
        actions.Controls.Add(_requirementsButton);
        _requirementsButton.Click += (_, _) => ShowRequirements(false);
        StyleButton(_orderButton, "Orden pagos + adicionales", false);
        _orderButton.Width = 205;
        actions.Controls.Add(_orderButton);
        _orderButton.Click += (_, _) =>
        {
            using var dialog = new DocumentOrderForm(_documentOrder);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _documentOrder = dialog.SelectedOrder.ToArray();
                AppendLog("Orden de PAGOS y ADICIONALES guardado: " + string.Join(" > ", _documentOrder));
            }
        };
        footer.Controls.Add(actions, 0, 2);
        _runButton.Click += async (_, _) => await PrepareAsync();
        _cancelButton.Click += (_, _) => _cts?.Cancel();
    }

    private static Control MakeSectionTitle(string title, string description)
    {
        var row = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = Padding.Empty };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        row.Controls.Add(new Label
        {
            Text = title, AutoSize = true, Dock = DockStyle.Fill,
            Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
            ForeColor = Navy, TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        row.Controls.Add(new Label
        {
            Text = "    " + description, Dock = DockStyle.Fill, AutoEllipsis = true,
            Font = new Font("Segoe UI", 8.5F), ForeColor = Gray,
            TextAlign = ContentAlignment.MiddleLeft
        }, 1, 0);
        return row;
    }

    private static void StyleButton(Button button, string text, bool primary)
    {
        button.Text = text;
        button.Height = 34;
        button.Margin = new Padding(6, 0, 0, 0);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = primary ? Blue : Color.FromArgb(230, 236, 245);
        button.ForeColor = primary ? Color.White : Navy;
        button.Font = new Font("Segoe UI Semibold", 9F);
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    private Control MakePathRow(string title, string hint, TextBox textBox,
        string buttonText, Action select, Action? clear = null, bool required = false, bool folder = false)
    {
        var outer = new Panel
        {
            Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 3),
            Padding = new Padding(11, 4, 11, 5), BackColor = Color.White
        };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = clear is null ? 3 : 4, RowCount = 2 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 128));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 134));
        if (clear is not null) layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62));
        var caption = new Label
        {
            Text = title + "   ·   " + hint, Dock = DockStyle.Fill,
            Font = new Font("Segoe UI Semibold", 9F), ForeColor = Navy,
            TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true
        };
        layout.Controls.Add(caption, 0, 0);
        layout.SetColumnSpan(caption, layout.ColumnCount);
        textBox.Dock = DockStyle.Fill;
        textBox.ReadOnly = true;
        textBox.BorderStyle = BorderStyle.FixedSingle;
        textBox.PlaceholderText = "Sin archivo seleccionado";
        textBox.BackColor = Color.FromArgb(249, 251, 254);
        textBox.Font = new Font("Segoe UI", 9F);
        textBox.Margin = new Padding(0, 3, 5, 1);
        layout.Controls.Add(textBox, 0, 1);
        var indicator = new Label
        {
            Dock = DockStyle.Fill, Font = new Font("Segoe UI Semibold", 8.0F),
            ForeColor = Gray, BackColor = Color.FromArgb(245, 247, 251),
            TextAlign = ContentAlignment.MiddleCenter,
            Margin = new Padding(3, 3, 5, 1), AutoEllipsis = true
        };
        layout.Controls.Add(indicator, 1, 1);
        _indicators[textBox] = (indicator, required, folder);
        var selectButton = new Button { Dock = DockStyle.Fill };
        StyleButton(selectButton, buttonText, false);
        selectButton.Margin = new Padding(2, 2, 2, 0);
        selectButton.Click += (_, _) => select();
        layout.Controls.Add(selectButton, 2, 1);
        if (clear is not null)
        {
            var clearButton = new Button { Dock = DockStyle.Fill };
            StyleButton(clearButton, "Quitar", false);
            clearButton.Font = new Font("Segoe UI", 8.3F);
            clearButton.Margin = new Padding(4, 2, 0, 0);
            clearButton.Click += (_, _) => clear();
            layout.Controls.Add(clearButton, 3, 1);
        }
        outer.Controls.Add(layout);
        return outer;
    }

    private void ClearPayment(int paymentNumber)
    {
        switch (paymentNumber)
        {
            case 1: _archiveFile = ""; _archiveBox.Clear(); break;
            case 2: _archiveFile2 = ""; _archiveBox2.Clear(); break;
            case 3: _archiveFile3 = ""; _archiveBox3.Clear(); break;
            case 4: _archiveCts = ""; _archiveCtsBox.Clear(); break;
            case 5: _archiveGratificacion = ""; _archiveGratificacionBox.Clear(); break;
            case 6: _archiveDescanso = ""; _archiveDescansoBox.Clear(); break;
            case 7: _archiveLicencia = ""; _archiveLicenciaBox.Clear(); break;
            case 8: _archiveSubsidios = ""; _archiveSubsidiosBox.Clear(); break;
            case 9: _archiveVacaciones = ""; _archiveVacacionesBox.Clear(); break;
        }
        UpdateState();
    }

    private void ClearPayflow()
    {
        _payflowFile = "";
        _payflowBox.Clear();
        UpdateState();
    }

    private void ClearCtr()
    {
        _ctrFile = "";
        _ctrBox.Clear();
        UpdateState();
    }

    private void SelectCtr()
    {
        using var dialog = new OpenFileDialog {
            Title = "Seleccionar constancia CTR/SCTR en PDF (se conservarán todas sus páginas)",
            Filter = "Archivos PDF (*.pdf)|*.pdf", CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _ctrFile = dialog.FileName;
        _ctrBox.Text = _ctrFile;
        UpdateState();
    }

    private void ClearVidaLey()
    {
        _vidaLeyFile = "";
        _vidaLeyBox.Clear();
        UpdateState();
    }

    private void SelectVidaLey()
    {
        using var dialog = new OpenFileDialog {
            Title = "Seleccionar constancia VIDA LEY en PDF (se conservarán todas sus páginas)",
            Filter = "Archivos PDF (*.pdf)|*.pdf", CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _vidaLeyFile = dialog.FileName;
        _vidaLeyBox.Text = _vidaLeyFile;
        UpdateState();
    }

    private void ClearBoletas()
    {
        _boletasFile = "";
        _boletasBox.Clear();
        UpdateState();
    }

    private void SelectBoletas()
    {
        using var dialog = new OpenFileDialog {
            Title = "Seleccionar un PDF con todas las boletas de pago",
            Filter = "Archivos PDF (*.pdf)|*.pdf",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _boletasFile = dialog.FileName;
        _boletasBox.Text = _boletasFile;
        UpdateState();
    }

    private void SelectArchive(int paymentNumber)
    {
        string paymentName = paymentNumber switch { 1 => "PAGO 1", 2 => "PAGO 2", 3 => "PAGO 3", 4 => "PAGO CTS", 5 => "PAGO GRATIFICACIÓN", 6 => "DESCANSO MEDICO", 7 => "LICENCIA", 8 => "SUBSIDIOS", 9 => "VACACIONES", _ => throw new ArgumentOutOfRangeException(nameof(paymentNumber)) };
        using var dialog = new OpenFileDialog {
            Title = $"Seleccionar archivo ZIP/RAR para {paymentName}",
            Filter = "Archivos comprimidos (*.zip;*.rar)|*.zip;*.rar", Multiselect = false, CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        switch (paymentNumber)
        {
            case 1: _archiveFile = dialog.FileName; _archiveBox.Text = _archiveFile; break;
            case 2: _archiveFile2 = dialog.FileName; _archiveBox2.Text = _archiveFile2; break;
            case 3: _archiveFile3 = dialog.FileName; _archiveBox3.Text = _archiveFile3; break;
            case 4: _archiveCts = dialog.FileName; _archiveCtsBox.Text = _archiveCts; break;
            case 5: _archiveGratificacion = dialog.FileName; _archiveGratificacionBox.Text = _archiveGratificacion; break;
            case 6: _archiveDescanso = dialog.FileName; _archiveDescansoBox.Text = _archiveDescanso; break;
            case 7: _archiveLicencia = dialog.FileName; _archiveLicenciaBox.Text = _archiveLicencia; break;
            case 8: _archiveSubsidios = dialog.FileName; _archiveSubsidiosBox.Text = _archiveSubsidios; break;
            case 9: _archiveVacaciones = dialog.FileName; _archiveVacacionesBox.Text = _archiveVacaciones; break;
        }
        UpdateState();
    }

    private void SelectPayflow()
    {
        using var dialog = new OpenFileDialog {
            Title = "Seleccionar Excel PAYFLOW (se procesará su primera hoja)",
            Filter = "Excel PAYFLOW (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _payflowFile = dialog.FileName;
        _payflowBox.Text = _payflowFile;
        UpdateState();
    }

    private void SelectExcel()
    {
        using var dialog = new OpenFileDialog { Title = "Seleccionar lista Excel", Filter = "Libros de Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm", CheckFileExists = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _excelFile = dialog.FileName;
        _excelBox.Text = _excelFile;
        UpdateState();
    }

    private void SelectDestination()
    {
        using var dialog = new FolderBrowserDialog { Description = "Seleccione la carpeta de destino", UseDescriptionForTitle = true, ShowNewFolderButton = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _destination = dialog.SelectedPath;
        _destinationBox.Text = _destination;
        UpdateState();
    }

    private void UpdateState()
    {
        bool running = _cts is not null;
        foreach (var (field, settings) in _indicators)
        {
            string path = field.Text.Trim();
            bool missing = string.IsNullOrWhiteSpace(path);
            bool exists = !missing && (settings.Folder ? Directory.Exists(path) : File.Exists(path));
            settings.Badge.Text = exists ? "✓ Listo" : missing ? (settings.Required ? "⚠ Requerido" : "⚠ Opcional") : "⚠ No existe";
            settings.Badge.ForeColor = exists ? Color.FromArgb(13, 112, 65) : missing && !settings.Required ? Gray : Color.FromArgb(162, 81, 17);
            settings.Badge.BackColor = exists ? Color.FromArgb(227, 247, 236) : missing && !settings.Required ? Color.FromArgb(245, 247, 251) : Color.FromArgb(255, 243, 224);
        }
        var sources = new[] { _archiveFile, _archiveFile2, _archiveFile3, _archiveCts, _archiveGratificacion,
            _archiveDescanso, _archiveLicencia, _archiveSubsidios, _archiveVacaciones,
            _payflowFile, _ctrFile, _vidaLeyFile, _boletasFile };
        int ready = sources.Count(File.Exists);
        int invalid = sources.Count(p => !string.IsNullOrWhiteSpace(p) && !File.Exists(p));
        if (!string.IsNullOrWhiteSpace(_boletasFile) && File.Exists(_boletasFile) &&
            !Path.GetExtension(_boletasFile).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) invalid++;
        bool onlyBoletas = File.Exists(_boletasFile) && sources.Where((p, i) => i != sources.Length - 1).All(string.IsNullOrWhiteSpace);
        bool requiredReady = Directory.Exists(_destination) && (File.Exists(_excelFile) || onlyBoletas);
        _counter.Text = $"Documentos listos: {ready}/13  |  Faltantes opcionales: {sources.Count(string.IsNullOrWhiteSpace)}  |  Rutas inválidas: {invalid}";
        _runButton.Enabled = !running && ready > 0 && invalid == 0 && requiredReady;
        _cancelButton.Enabled = running;
        _orderButton.Enabled = !running;
        _status.Text = running ? "Generando conciliación…" : _runButton.Enabled ? "✓ Preparado para generar" :
            invalid > 0 ? "⚠ Revisar rutas" : !requiredReady ? "⚠ Falta destino o Excel (excepto BOLETAS)" : "⚠ Seleccione un documento";
    }

    private async Task PrepareAsync()
    {
        var archive = _archiveFile;
        var archive2 = _archiveFile2;
        var archive3 = _archiveFile3;
        var archiveCts = _archiveCts;
        var archiveGratificacion = _archiveGratificacion;
        var archiveDescanso = _archiveDescanso;
        var archiveLicencia = _archiveLicencia;
        var archiveSubsidios = _archiveSubsidios;
        var archiveVacaciones = _archiveVacaciones;
        var payflow = _payflowFile;
        var ctr = _ctrFile;
        var vidaLey = _vidaLeyFile;
        var boletas = _boletasFile;
        var excel = _excelFile;
        var dest = _destination;
        var documentOrder = _documentOrder.ToArray();
        if (string.IsNullOrWhiteSpace(dest)) return;
        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        UpdateState();
        _progress.Value = 0;
        AppendLog("--- INICIO " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ---");
        if (!string.IsNullOrWhiteSpace(boletas)) AppendLog("BOLETAS: se leerán DNI y NOMBRE directamente del PDF, sin necesitar Excel. Cada página se guardará individualmente; si falla OCR quedará pendiente.");
        if (!string.IsNullOrWhiteSpace(vidaLey)) AppendLog("VIDA LEY: se conservará el PDF íntegro y se enmarcarán los trabajadores coincidentes por DNI.");
        if (!string.IsNullOrWhiteSpace(ctr)) AppendLog("CTR: el PDF original completo se conservará y solo se enmarcarán filas por DNI coincidente.");
        if (!string.IsNullOrWhiteSpace(payflow)) AppendLog("PAYFLOW seleccionado: se procesará la primera hoja, sin Id de empleado, con PDF general y PDF individuales por DNI.");
        var progress = new Progress<PreparationProgress>(p =>
        {
            _progress.Value = Math.Clamp(p.Percentage, 0, 100);
            if (!string.IsNullOrEmpty(p.Message)) AppendLog(p.Message);
        });
        try
        {
            var result = await Task.Run(() => ProcessingEngine.Process(
                new[] { new PaymentArchive("PAGO 1", archive), new PaymentArchive("PAGO 2", archive2),
                    new PaymentArchive("PAGO 3", archive3), new PaymentArchive("PAGO CTS", archiveCts),
                    new PaymentArchive("PAGO GRATIFICACIÓN", archiveGratificacion),
                    new PaymentArchive("DESCANSO MEDICO", archiveDescanso),
                    new PaymentArchive("LICENCIA", archiveLicencia),
                    new PaymentArchive("SUBSIDIOS", archiveSubsidios),
                    new PaymentArchive("VACACIONES", archiveVacaciones) },
                excel, dest, progress, token, payflow, ctr, vidaLey, boletas, documentOrder), token);
            AppendLog($"COMPLETADO. PDF bancarios: {result.PdfCount}; individuales: {result.IndividualCount}; revisión: {result.ReviewCount}. ZIP: {result.ZipPath}");
            MessageBox.Show(this, $"Conciliación finalizada.\n\nPDF bancarios: {result.PdfCount}\nPDF individuales: {result.IndividualCount}\nCasos para revisión: {result.ReviewCount}\n\nZIP final:\n{result.ZipPath}", "Conciliación terminada", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (OperationCanceledException)
        {
            AppendLog("CANCELADO por el usuario. Puede quedar una carpeta temporal incompleta.");
        }
        catch (Exception ex)
        {
            AppendLog("ERROR: " + ex.Message);
            MessageBox.Show(this, ex.Message, "No fue posible preparar los archivos", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            UpdateState();
        }
    }

    private void AppendLog(string message)
    {
        if (InvokeRequired) { BeginInvoke(() => AppendLog(message)); return; }
        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        _logBox.ScrollToCaret();
    }
}

