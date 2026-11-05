using System.Globalization;
using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace ConciliadorBancario;

// PAYFLOW es una lista de descuentos, no una constancia de pago bancario.
internal static class PayflowProcessor
{
    internal sealed record Record(string Dni, string Name, string Surname, string Company, string Ruc, decimal Amount, int SourceRow);
    internal sealed record Outcome(int Records, int IndividualPdfs, int Warnings);
    private static readonly XColor Blue = XColor.FromArgb(55, 86, 223);
    private static readonly XColor Cyan = XColor.FromArgb(41, 202, 226);

    public static Outcome Generate(string excelPath, string root, IProgress<PreparationProgress> progress, CancellationToken ct)
    {
        if (!File.Exists(excelPath)) throw new FileNotFoundException("No se encontró el Excel PAYFLOW.", excelPath);
        if (!new[] { ".xlsx", ".xlsm" }.Contains(Path.GetExtension(excelPath).ToLowerInvariant()))
            throw new InvalidDataException("PAYFLOW requiere un Excel XLSX o XLSM. No se admiten PDF, CSV ni ZIP.");
        using var wb = new XLWorkbook(excelPath);
        var sheet = wb.Worksheet(1); // Requisito: únicamente primera hoja.
        var header = sheet.RowsUsed().Take(40).Select(row => new
        {
            Row = row.RowNumber(),
            Map = row.CellsUsed().ToDictionary(c => Normalize(c.GetFormattedString()), c => c.Address.ColumnNumber)
        }).FirstOrDefault(x => x.Map.ContainsKey("nombre") && x.Map.ContainsKey("apellidos") &&
            (x.Map.ContainsKey("dnioce") || x.Map.ContainsKey("dni")) && x.Map.ContainsKey("nombreempresa") &&
            x.Map.ContainsKey("rucempresa") && x.Map.Keys.Any(k => k.StartsWith("importetotaladeducir")));
        if (header is null) throw new InvalidDataException("Primera hoja PAYFLOW: no se hallaron los encabezados Nombre, Apellidos, DNI o CE, Nombre empresa, RUC empresa e Importe total a deducir.");
        int Column(string name) => header.Map[name];
        int dniCol = header.Map.TryGetValue("dnioce", out int docColumn) ? docColumn : Column("dni");
        int amtCol = header.Map.First(k => k.Key.StartsWith("importetotaladeducir")).Value;
        var rows = new List<Record>();
        var warnings = new List<string>();
        int last = sheet.LastRowUsed()?.RowNumber() ?? header.Row;
        for (int r = header.Row + 1; r <= last; r++)
        {
            ct.ThrowIfCancellationRequested();
            var row = sheet.Row(r);
            string name = row.Cell(Column("nombre")).GetFormattedString().Trim();
            string surname = row.Cell(Column("apellidos")).GetFormattedString().Trim();
            string rawDni = row.Cell(dniCol).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(surname) && string.IsNullOrWhiteSpace(rawDni)) continue;
            // No convertir DNI/CE a número: se conservan ceros iniciales.
            string dni = rawDni;
            if (row.Cell(dniCol).TryGetValue<long>(out long numeric) && numeric >= 0 && numeric < 100_000_000)
                dni = numeric.ToString("D8", CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(dni) || string.IsNullOrWhiteSpace(name + surname))
            {
                warnings.Add($"Fila {r}: falta documento o nombre, revisar en Excel original.");
                continue;
            }
            decimal amount;
            var value = row.Cell(amtCol);
            if (!value.TryGetValue<decimal>(out amount) &&
                !decimal.TryParse(value.GetFormattedString().Replace("S/", "").Trim(), NumberStyles.Currency,
                    CultureInfo.GetCultureInfo("es-PE"), out amount) &&
                !decimal.TryParse(value.GetFormattedString().Replace("S/", "").Trim(), NumberStyles.Currency,
                    CultureInfo.InvariantCulture, out amount))
            {
                warnings.Add($"Fila {r}: importe no legible; se omite el registro.");
                continue;
            }
            string ruc = row.Cell(Column("rucempresa")).GetFormattedString().Trim();
            string company = row.Cell(Column("nombreempresa")).GetFormattedString().Trim();
            rows.Add(new Record(dni, name, surname, company, ruc, amount, r));
        }
        if (rows.Count == 0) throw new InvalidDataException("No se encontraron trabajadores válidos en la primera hoja PAYFLOW.");
        var duplicateDnis = rows.GroupBy(x => x.Dni, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1).ToList();
        foreach (var duplicate in duplicateDnis) warnings.Add($"DNI/CE repetido {duplicate.Key}: {duplicate.Count()} filas; se genera un PDF independiente por fila.");

        string generalDir = Path.Combine(root, "06_PAYFLOW", "RESUMEN_GENERAL");
        string indivDir = Path.Combine(root, "06_PAYFLOW", "PDF_INDIVIDUALES");
        Directory.CreateDirectory(generalDir);
        Directory.CreateDirectory(indivDir);
        string generalPath = Path.Combine(generalDir, "PAYFLOW_RESUMEN_GENERAL.pdf");
        GeneratePdf(generalPath, rows, -1);
        progress.Report(new(93, $"PAYFLOW: PDF general generado; {rows.Count} registros de la primera hoja."));
        var generated = 0;
        for (int i = 0; i < rows.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var rec = rows[i];
            string baseName = $"{CleanName(rec.Dni)}_{CleanName(rec.Surname + " " + rec.Name)}_PAYFLOW";
            var path = Path.Combine(indivDir, baseName + ".pdf");
            int suffix = 2;
            while (File.Exists(path)) path = Path.Combine(indivDir, $"{baseName}_{suffix++}.pdf");
            // Exportar SOLAMENTE la pagina que contiene esta fila;
            // conserva el resto de trabajadores visibles en esa pagina y encierra la fila elegida.
            GeneratePdf(path, rows, i);
            generated++;
        }
        if (warnings.Count != 0) File.WriteAllLines(Path.Combine(root, "06_PAYFLOW", "PAYFLOW_ADVERTENCIAS.txt"), warnings, Encoding.UTF8);
        progress.Report(new(95, $"PAYFLOW: {generated} PDF individuales enmarcados; {warnings.Count} advertencias."));
        return new Outcome(rows.Count, generated, warnings.Count);
    }

    private static void GeneratePdf(string destination, IReadOnlyList<Record> records, int highlightedIndex)
    {
        // A3 horizontal (en puntos PDF). Nunca exceder el límite de altura de PDFsharp.
        // Si el resumen no cabe, crear automáticamente páginas adicionales, sin perder filas.
        const double width = 1190, pageHeight = 842, margin = 30, top = 108;
        const double headerHeight = 38, rowHeight = 22, bottomMargin = 42;
        int rowsPerPage = Math.Max(1, (int)Math.Floor((pageHeight - top - headerHeight - bottomMargin) / rowHeight));
        int pageCount = (records.Count + rowsPerPage - 1) / rowsPerPage;
        using var doc = new PdfDocument();
        var titleFont = new XFont("Arial", 31, XFontStyleEx.Bold);
        var headerFont = new XFont("Arial", 11, XFontStyleEx.Regular);
        var valueFont = new XFont("Arial", 10.4, XFontStyleEx.Regular);
        var strongFont = new XFont("Arial", 10.5, XFontStyleEx.Bold);
        var footerFont = new XFont("Arial", 9, XFontStyleEx.Regular);
        string[] headers = { "Nombre", "Apellidos", "DNI o CE", "Nombre empresa", "RUC empresa", "Importe total a deducir S/." };
        double[] widths = { 192, 215, 138, 275, 140, 170 };
        double tableLeft = margin;

        // En el PDF general: todas las paginas.
        // En el PDF individual: exclusivamente la pagina original donde cae el trabajador.
        int firstPage = highlightedIndex < 0 ? 0 : highlightedIndex / rowsPerPage;
        int lastPageExclusive = highlightedIndex < 0 ? pageCount : firstPage + 1;
        for (int pageIndex = firstPage; pageIndex < lastPageExclusive; pageIndex++)
        {
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(width);
            page.Height = XUnit.FromPoint(pageHeight);
            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawString("PAYFLOW", titleFont, new XSolidBrush(Cyan),
                new XRect(margin + 38, 22, 310, 50), XStringFormats.CenterLeft);
            gfx.DrawEllipse(new XPen(Cyan, 2), margin, 29, 33, 33);
            gfx.DrawLine(new XPen(XColor.FromArgb(66, 128, 255), 1.5), margin + 3, 45, margin + 30, 45);
            double y = top;
            gfx.DrawRectangle(new XSolidBrush(Blue), tableLeft, y, widths.Sum(), headerHeight);
            double x = tableLeft;
            for (int col = 0; col < headers.Length; col++)
            {
                gfx.DrawString(headers[col], headerFont, XBrushes.White,
                    new XRect(x + 4, y + 4, widths[col] - 8, headerHeight - 8), XStringFormats.CenterLeft);
                x += widths[col];
            }
            y += headerHeight;
            int first = pageIndex * rowsPerPage;
            int lastExclusive = Math.Min(records.Count, first + rowsPerPage);
            for (int rowNo = first; rowNo < lastExclusive; rowNo++)
            {
                var r = records[rowNo];
                x = tableLeft;
                string[] values = { r.Name, r.Surname, r.Dni, r.Company, r.Ruc,
                    r.Amount.ToString("N2", CultureInfo.InvariantCulture) };
                for (int col = 0; col < values.Length; col++)
                {
                    var rect = new XRect(x + 4, y + 1, widths[col] - 8, rowHeight - 2);
                    string cellText = Fit(gfx, values[col], col == 3 ? strongFont : valueFont, rect.Width);
                    gfx.DrawString(cellText, col == 3 ? strongFont : valueFont, XBrushes.Black, rect,
                        col == 5 ? XStringFormats.CenterRight : XStringFormats.CenterLeft);
                    x += widths[col];
                }
                gfx.DrawLine(new XPen(XColors.LightGray, 0.5), tableLeft,
                    y + rowHeight, tableLeft + widths.Sum(), y + rowHeight);
                if (highlightedIndex == rowNo)
                    gfx.DrawRectangle(new XPen(XColors.Black, 2.4), tableLeft + 1,
                        y + 0.8, widths.Sum() - 2, rowHeight - 1.6);
                y += rowHeight;
            }
            gfx.DrawString($"Página {pageIndex + 1} de {pageCount}", footerFont, XBrushes.Gray,
                new XRect(margin, pageHeight - bottomMargin + 8, widths.Sum(), 18), XStringFormats.CenterRight);
        }
        doc.Save(destination);
    }

    private static string Fit(XGraphics gfx, string input, XFont font, double width)
    {
        if (gfx.MeasureString(input, font).Width <= width) return input;
        const string ending = "…";
        int end = input.Length;
        while (end > 0 && gfx.MeasureString(input[..end] + ending, font).Width > width) end--;
        return input[..end] + ending;
    }

    private static string Normalize(string s) => new string(s.Normalize(NormalizationForm.FormD)
        .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
        .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    private static string CleanName(string text)
    {
        var s = new string(text.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
        return s.Trim().Replace(' ', '_').Length > 85 ? s.Trim().Replace(' ', '_')[..85] : s.Trim().Replace(' ', '_');
    }
}
