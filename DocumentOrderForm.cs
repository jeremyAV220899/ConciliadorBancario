using System.Drawing;

namespace ConciliadorBancario;

internal sealed class DocumentOrderForm : Form
{
    private readonly ListBox _payments = new();
    private readonly ListBox _additional = new();
    internal IReadOnlyList<string> SelectedOrder { get; private set; }

    internal DocumentOrderForm(IReadOnlyList<string> current)
    {
        SelectedOrder = DocumentOrderSettings.Normalize(current);
        Text = "Configurar orden del PDF final | Dos bloques";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(780, 600);
        MinimumSize = new Size(660, 480);
        Font = new Font("Segoe UI", 10F);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(15), ColumnCount = 1, RowCount = 3 };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
        Controls.Add(root);
        root.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            Text = "El PDF final tendrá dos bloques consecutivos. Primero se generan todos los PAGOS por trabajador " +
                   "en orden del Excel; al terminar se agregan los ADICIONALES completos. " +
                   "Ordene los documentos de cada bloque con Subir / Bajar y guarde su preferencia."
        }, 0, 0);

        var columns = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        columns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 61));
        columns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 39));
        root.Controls.Add(columns, 0, 1);
        AddGroup(columns, 0, "1. PAGOS (primera parte)", _payments,
            DocumentOrderSettings.NormalizeGroup(current, DocumentOrderSettings.PaymentOrder),
            DocumentOrderSettings.PaymentOrder);
        AddGroup(columns, 1, "2. DOCUMENTOS ADICIONALES (al final)", _additional,
            DocumentOrderSettings.NormalizeGroup(current, DocumentOrderSettings.AdditionalOrder),
            DocumentOrderSettings.AdditionalOrder);

        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var save = new Button { Text = "Guardar ambos órdenes", Width = 205, Height = 38 };
        var cancel = new Button { Text = "Cancelar", Width = 105, Height = 38, DialogResult = DialogResult.Cancel };
        save.Click += (_, _) =>
        {
            try
            {
                var payments = _payments.Items.Cast<string>().ToArray();
                var additional = _additional.Items.Cast<string>().ToArray();
                DocumentOrderSettings.SaveGroups(payments, additional);
                SelectedOrder = payments.Concat(additional).ToArray();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo guardar el orden: " + ex.Message,
                    "Error de configuración", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        };
        footer.Controls.Add(save);
        footer.Controls.Add(cancel);
        root.Controls.Add(footer, 0, 2);
        AcceptButton = save;
        CancelButton = cancel;
    }

    private static void AddGroup(TableLayoutPanel parent, int col, string title, ListBox list,
        IReadOnlyList<string> current, IReadOnlyList<string> defaults)
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(5), ColumnCount = 1, RowCount = 2 };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.Controls.Add(new Label { Dock = DockStyle.Fill, Text = title, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }, 0, 0);
        var middle = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        list.Dock = DockStyle.Fill;
        list.IntegralHeight = false;
        list.Items.AddRange(current.Cast<object>().ToArray());
        if (list.Items.Count > 0) list.SelectedIndex = 0;
        middle.Controls.Add(list, 0, 0);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        var up = new Button { Text = "↑ Subir", Width = 90, Height = 37 };
        var down = new Button { Text = "↓ Bajar", Width = 90, Height = 37 };
        var reset = new Button { Text = "Restablecer", Width = 90, Height = 37 };
        up.Click += (_, _) => Move(list, -1);
        down.Click += (_, _) => Move(list, 1);
        reset.Click += (_, _) =>
        {
            list.Items.Clear();
            list.Items.AddRange(defaults.Cast<object>().ToArray());
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        };
        actions.Controls.Add(up);
        actions.Controls.Add(down);
        actions.Controls.Add(reset);
        middle.Controls.Add(actions, 1, 0);
        panel.Controls.Add(middle, 0, 1);
        parent.Controls.Add(panel, col, 0);
    }

    private static void Move(ListBox list, int offset)
    {
        int index = list.SelectedIndex, next = index + offset;
        if (index < 0 || next < 0 || next >= list.Items.Count) return;
        var value = list.Items[index];
        list.Items.RemoveAt(index);
        list.Items.Insert(next, value);
        list.SelectedIndex = next;
    }
}
