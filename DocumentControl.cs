using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ConciliadorBancario;

/// <summary>
/// Auditoría documental posterior a la generación de todos los módulos.
/// OK significa presencia de comprobante generado / fila localizada en constancia;
/// NO confirma pago efectivamente realizado ni cobertura vigente.
/// </summary>
internal static class DocumentControl
{
    private static readonly string[] Payments = {
        "PAGO 1", "PAGO 2", "PAGO 3", "PAGO CTS", "PAGO GRATIFICACIÓN",
        "DESCANSO MEDICO", "LICENCIA", "SUBSIDIOS", "VACACIONES"
    };
    private static readonly string[] All = Payments.Concat(new[] { "CTR", "VIDA LEY", "BOLETA", "PAYFLOW" }).ToArray();
    private static readonly Regex LeadingDni = new(@"^(?<dni>\d{8})[_\s-]", RegexOptions.Compiled);
    private static readonly Regex ReportDni = new(@"^\d{8}$", RegexOptions.Compiled);

    public static void AppendControl(string summaryPath, string root, IReadOnlyList<Worker> workers,
        IReadOnlyList<string> selectedPayments, bool payflowSelected, bool ctrSelected,
        bool vidaSelected, bool boletasSelected, IProgress<PreparationProgress> progress, CancellationToken ct)
    {
        var selected = new HashSet<string>(selectedPayments, StringComparer.OrdinalIgnoreCase);
        if (payflowSelected) selected.Add("PAYFLOW");
        if (ctrSelected) selected.Add("CTR");
        if (vidaSelected) selected.Add("VIDA LEY");
        if (boletasSelected) selected.Add("BOLETA");

        // Índice de PDFs: se inspeccionan las rutas de salida efectivamente creadas,
        // no se da OK por el mero hecho de haber cargado el archivo original.
        var files = Directory.EnumerateFiles(root, "*.pdf", SearchOption.AllDirectories).ToList();
        var found = All.ToDictionary(k => k, _ => new HashSet<string>(StringComparer.Ordinal), StringComparer.OrdinalIgnoreCase);
        var names = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var w in workers)
            if (!names.ContainsKey(w.Dni) && ReportDni.IsMatch(w.Dni)) names[w.Dni] = w.Name;
        int pagesWithoutIdentity = 0;

        foreach (string file in files)
        {
            ct.ThrowIfCancellationRequested();
            string rel = Path.GetRelativePath(root, file).Replace('\\', '/');
            string baseName = Path.GetFileNameWithoutExtension(file);
            var leading = LeadingDni.Match(baseName);
            string? dni = leading.Success ? leading.Groups["dni"].Value : null;
            if (rel.StartsWith("02_PDF_INDIVIDUALES/", StringComparison.OrdinalIgnoreCase))
            {
                string payment = rel.Split('/').Skip(1).FirstOrDefault() ?? "";
                if (dni is not null && found.TryGetValue(payment, out var group)) group.Add(dni);
            }
            else if (rel.StartsWith("05_BOLETAS/", StringComparison.OrdinalIgnoreCase))
            {
                if (rel.Contains("/PENDIENTES_DE_IDENTIFICAR/", StringComparison.OrdinalIgnoreCase))
                    pagesWithoutIdentity++;
                else if (dni is not null)
                {
                    found["BOLETA"].Add(dni);
                    AddNameFromFile(names, dni, baseName);
                }
            }
            else if (rel.StartsWith("06_PAYFLOW/PDF_INDIVIDUALES/", StringComparison.OrdinalIgnoreCase))
            {
                if (dni is not null)
                {
                    found["PAYFLOW"].Add(dni);
                    AddNameFromFile(names, dni, baseName.Replace("_PAYFLOW", "", StringComparison.OrdinalIgnoreCase));
                }
            }
        }
        // CTR/VIDA LEY producen un único PDF completo. El reporte del procesador
        // contiene los DNI que NO pudieron marcarse; OK requiere que exista su PDF generado.
        ApplyFullPdfReport(root, "03_CTR_PDF_COMPLETO_ENMARCADO", "CTR", found["CTR"], workers);
        ApplyFullPdfReport(root, "04_VIDA_LEY_PDF_COMPLETO_ENMARCADO", "VIDA_LEY", found["VIDA LEY"], workers);

        using var book = new XLWorkbook(summaryPath);
        if (book.Worksheets.TryGetWorksheet("CONTROL DOCUMENTAL", out var oldSheet)) oldSheet.Delete();
        var ws = book.AddWorksheet("CONTROL DOCUMENTAL");
        ws.Position = 1;
        int last = All.Length + 3;
        ws.Range(1, 1, 1, last).Merge();
        ws.Cell(1, 1).Value = "CONTROL DOCUMENTAL POR TRABAJADOR";
        ws.Range(1, 1, 1, last).Style.Fill.BackgroundColor = XLColor.FromHtml("#14243D");
        ws.Range(1, 1, 1, last).Style.Font.FontColor = XLColor.White;
        ws.Range(1, 1, 1, last).Style.Font.Bold = true;
        ws.Range(1, 1, 1, last).Style.Font.FontSize = 15;
        ws.Row(1).Height = 28;
        ws.Cell(2, 1).Value = "OK = PDF generado para DNI / registro enmarcado en constancia; NO = no encontrado; — = documento no seleccionado.";
        ws.Range(2, 1, 2, last).Merge();
        ws.Cell(3, 1).Value = "Nota: OK no certifica depósito bancario ni cobertura del seguro; verifica solo evidencia documental.";
        ws.Range(3, 1, 3, last).Merge();
        ws.Cell(5, 1).Value = "DNI";
        ws.Cell(5, 2).Value = "NOMBRE COMPLETO";
        for (int i = 0; i < All.Length; i++) ws.Cell(5, i+3).Value = All[i];
        ws.Cell(5, last).Value = "OBSERVACIÓN";
        var header = ws.Range(5, 1, 5, last);
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#3159D8");
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        ws.Row(5).Height = 26;
        int rowNo = 6;
        foreach (var entry in names.OrderBy(x => x.Key, StringComparer.Ordinal))
        {
            ct.ThrowIfCancellationRequested();
            ws.Cell(rowNo, 1).SetValue(entry.Key);
            ws.Cell(rowNo, 1).Style.NumberFormat.Format = "@";
            ws.Cell(rowNo, 2).Value = entry.Value;
            var missing = new List<string>();
            foreach (var (label, idx) in All.Select((v, i) => (v, i)))
            {
                string state = !selected.Contains(label) ? "—" : found[label].Contains(entry.Key) ? "OK" : "NO";
                var cell = ws.Cell(rowNo, idx + 3);
                cell.Value = state;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                if (state == "OK") cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#DCFCE7");
                if (state == "NO")
                {
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FEE2E2");
                    missing.Add(label);
                }
            }
            string note = missing.Count == 0 ? "Sin documentos faltantes entre los seleccionados" :
                "No generado / no encontrado: " + string.Join(", ", missing);
            if (boletasSelected && pagesWithoutIdentity > 0) note += $". Boletas sin identificar en carpeta: {pagesWithoutIdentity} (revisar Excel de boletas)";
            ws.Cell(rowNo, last).Value = note;
            rowNo++;
        }
        if (rowNo == 6)
        {
            ws.Cell(rowNo, 1).Value = "Sin DNI identificables en documentos generados";
            rowNo++;
        }
        ws.Column(1).Width = 15;
        ws.Column(2).Width = 43;
        for (int i = 3; i < last; i++) ws.Column(i).Width = i <= 7 ? 18 : 20;
        ws.Column(last).Width = 85;
        ws.SheetView.FreezeRows(5);
        ws.SheetView.FreezeColumns(2);
        ws.Range(5, 1, rowNo - 1, last).SetAutoFilter();
        book.SaveAs(summaryPath);
        progress.Report(new(99, $"CONTROL DOCUMENTAL: {rowNo - 6} DNI revisados; {files.Count} PDFs rastreados en carpetas de salida."));
    }

    private static void ApplyFullPdfReport(string root, string folder, string label,
        HashSet<string> matched, IReadOnlyList<Worker> workers)
    {
        string dir = Path.Combine(root, folder);
        if (!Directory.Exists(dir) || !Directory.EnumerateFiles(dir, "*.pdf").Any()) return;
        string reportPath = Path.Combine(dir, label + "_REPORTE.txt");
        if (!File.Exists(reportPath)) return;
        var missing = File.ReadLines(reportPath).Select(x => x.Trim())
            .Where(x => ReportDni.IsMatch(x)).ToHashSet(StringComparer.Ordinal);
        foreach (var dni in workers.Select(w => w.Dni).Distinct())
            if (ReportDni.IsMatch(dni) && !missing.Contains(dni)) matched.Add(dni);
    }

    private static void AddNameFromFile(Dictionary<string, string> names, string dni, string baseName)
    {
        if (names.ContainsKey(dni)) return;
        string value = baseName.Length > dni.Length + 1 ? baseName[(dni.Length+1)..] : "";
        value = value.Replace('_', ' ').Trim();
        if (value.Length > 0) names[dni] = value;
    }
}
