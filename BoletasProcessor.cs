using System.Diagnostics;
using ClosedXML.Excel;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;

namespace ConciliadorBancario;

internal sealed record BoletasOutcome(int Sources, int Created, int Reviews);

/// <summary>
/// Boletas: conserva la página ORIGINAL de la boleta y la guarda de forma individual.
/// El DNI y nombre se leen DIRECTAMENTE de la boleta (sin consultar Excel); las páginas con identificación ambigua se guardan pendientes.
/// Cuando la página es escaneada intenta OCR local con Tesseract y pdftoppm (si están disponibles).
/// No modifica el motor bancario ni CTR/VIDA LEY/PAYFLOW.
/// </summary>
internal static class BoletasProcessor
{
    private static readonly Regex DniField = new(@"(?:D\s*N\s*I|DOCUMENTO\s+DE\s+IDENTIDAD)\s*[:.\-]?\s*(?:N[°ºo.]\s*)?(\d[\d\s.\-]{6,14})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static BoletasOutcome Generate(string source, string root,
        IProgress<PreparationProgress> progress, CancellationToken ct)
    {
        if (!File.Exists(source)) throw new FileNotFoundException("No existe el PDF de BOLETAS.", source);
        if (!Path.GetExtension(source).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("BOLETAS requiere un único PDF multipágina.");

        string dest = Path.Combine(root, "05_BOLETAS");
        string staging = Path.Combine(dest, "SEPARADAS_SIN_RENOMBRAR");
        string pending = Path.Combine(dest, "PENDIENTES_DE_IDENTIFICAR");
        Directory.CreateDirectory(dest);
        Directory.CreateDirectory(staging);
        Directory.CreateDirectory(pending);
        var report = new List<string> { "ARCHIVO;PAGINA;DNI;TRABAJADOR;ESTADO;PDF_GENERADO;OBSERVACION" };
        var reportRows = new List<(string Source, int Page, string Dni, string Name, string Status, string Pdf, string Note)>();
        var pages = new List<(int Number, string Path)>();
        var tools = ResolveOcrTools();
        int identified = 0, reviews = 0;

        // ETAPA 1: Separar y guardar todas las páginas originales ANTES de ejecutar OCR.
        // El nombre temporal incluye un identificador único para no sobrescribir ejecuciones previas.
        string batch = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N")[..6];
        using (var original = PdfReader.Open(source, PdfDocumentOpenMode.Import))
        {
            int total = original.PageCount;
            progress.Report(new(91, $"BOLETAS: etapa 1/2, separando {total} páginas antes del OCR."));
            for (int i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();
                string part = Path.Combine(staging, $"{batch}_PAGINA_{i + 1:D4}.pdf");
                using (var single = new PdfSharp.Pdf.PdfDocument())
                {
                    single.AddPage(original.Pages[i]);
                    single.Save(part);
                }
                pages.Add((i + 1, part));
                if ((i + 1) % 25 == 0 || i + 1 == total)
                    progress.Report(new(92, $"BOLETAS: separadas {i + 1}/{total}; analizando DNI y nombre a continuación."));
            }
        }

        // ETAPA 2: Analizar cada PDF individual, leyendo por separado el campo DNI
        // y el campo Nombres y Apellidos, siempre en la tabla superior PLAME.
        if (tools.Renderer is null || tools.Ocr is null)
            progress.Report(new(93, "BOLETAS: faltan Tesseract o Poppler. Se conservarán todas las boletas en pendientes."));
        else
            progress.Report(new(93, $"BOLETAS: etapa 2/2, OCR por zonas fijas ({tools.Language})."));

        string temp = Path.Combine(Path.GetTempPath(), "BOLETAS_OCR_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        try
        {
            foreach (var page in pages)
            {
                ct.ThrowIfCancellationRequested();
                Identity? identity = null;
                string observation = "";
                try
                {
                    if (tools.Renderer is null || tools.Ocr is null)
                    {
                        observation = "OCR no disponible; verificar Tesseract y Poppler";
                    }
                    else
                    {
                        identity = ReadIdentityFromIndividualPdf(page.Path, temp, tools, ct);
                        if (identity is null)
                            observation = "El DNI o nombre de la cabecera no pudo verificarse; revisión manual";
                    }
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    observation = "Fallo de OCR: " + ex.Message;
                }

                string fileName = identity is null
                    ? $"{batch}_PAGINA_{page.Number:D4}_SIN_IDENTIFICAR.pdf"
                    : SafeFile(identity.Dni + "_" + identity.Name) + ".pdf";
                string target = UniquePath(Path.Combine(identity is null ? pending : dest, fileName));
                File.Move(page.Path, target);
                if (identity is null)
                {
                    reviews++;
                    report.Add(Row(Path.GetFileName(source), page.Number, "", "", "PENDIENTE", Path.GetRelativePath(dest, target), observation));
                    reportRows.Add((Path.GetFileName(source), page.Number, "", "", "PENDIENTE", Path.GetRelativePath(dest, target), observation));
                }
                else
                {
                    identified++;
                    report.Add(Row(Path.GetFileName(source), page.Number, identity.Dni, identity.Name, "IDENTIFICADA", Path.GetRelativePath(dest, target), "OCR de DNI y nombre en zonas independientes"));
                    reportRows.Add((Path.GetFileName(source), page.Number, identity.Dni, identity.Name, "IDENTIFICADA", Path.GetRelativePath(dest, target), "OCR de DNI y nombre en zonas independientes"));
                }
                if (page.Number % 10 == 0 || page.Number == pages.Count)
                    progress.Report(new(94, $"BOLETAS: {page.Number}/{pages.Count}, renombradas {identified}, pendientes {reviews}."));
            }
        }
        finally
        {
            try { Directory.Delete(temp, true); } catch { }
            try { if (!Directory.EnumerateFileSystemEntries(staging).Any()) Directory.Delete(staging); } catch { }
            File.WriteAllLines(Path.Combine(dest, "BOLETAS_REPORTE.csv"), report, new UTF8Encoding(true));
            WriteExcelReport(Path.Combine(dest, "BOLETAS_REPORTE.xlsx"), reportRows);
        }
        progress.Report(new(95, $"BOLETAS: {pages.Count} PDF separados; {identified} renombrados DNI_NOMBRE; {reviews} pendientes."));
        return new(1, pages.Count, reviews);
    }

    private static void WriteExcelReport(string path,
        IReadOnlyList<(string Source, int Page, string Dni, string Name, string Status, string Pdf, string Note)> records)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Boletas");
        ws.Cell("A1").Value = "REPORTE DE BOLETAS DE PAGO";
        ws.Range("A1:G1").Merge();
        ws.Range("A1:G1").Style.Font.Bold = true;
        ws.Range("A1:G1").Style.Font.FontSize = 16;
        ws.Range("A1:G1").Style.Font.FontColor = XLColor.White;
        ws.Range("A1:G1").Style.Fill.BackgroundColor = XLColor.FromHtml("#14243D");
        ws.Row(1).Height = 30;
        ws.Cell("A2").Value = "Páginas procesadas";
        ws.Cell("B2").Value = records.Count;
        ws.Cell("C2").Value = "Identificadas";
        ws.Cell("D2").Value = records.Count(r => r.Status == "IDENTIFICADA");
        ws.Cell("E2").Value = "Pendientes";
        ws.Cell("F2").Value = records.Count(r => r.Status == "PENDIENTE");
        var headers = new[] { "PDF ORIGINAL", "PÁGINA", "DNI", "NOMBRE COMPLETO", "ESTADO", "PDF GENERADO", "OBSERVACIÓN" };
        for (int col = 0; col < headers.Length; col++) ws.Cell(4, col + 1).Value = headers[col];
        var head = ws.Range(4, 1, 4, 7);
        head.Style.Fill.BackgroundColor = XLColor.FromHtml("#3159D8");
        head.Style.Font.FontColor = XLColor.White;
        head.Style.Font.Bold = true;
        for (int i = 0; i < records.Count; i++)
        {
            var r = records[i];
            int row = i + 5;
            ws.Cell(row, 1).Value = r.Source;
            ws.Cell(row, 2).Value = r.Page;
            ws.Cell(row, 3).SetValue(r.Dni);
            ws.Cell(row, 3).Style.NumberFormat.Format = "@"; // preservar ceros iniciales
            ws.Cell(row, 4).Value = r.Name;
            ws.Cell(row, 5).Value = r.Status;
            ws.Cell(row, 6).Value = r.Pdf;
            ws.Cell(row, 7).Value = r.Note;
            if (r.Status == "PENDIENTE")
                ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF0C2");
            else
                ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#DFF5E6");
        }
        ws.Column(1).Width = 29;
        ws.Column(2).Width = 10;
        ws.Column(3).Width = 15;
        ws.Column(4).Width = 43;
        ws.Column(5).Width = 17;
        ws.Column(6).Width = 65;
        ws.Column(7).Width = 66;
        ws.SheetView.FreezeRows(4);
        ws.Range(4, 1, Math.Max(4, records.Count + 4), 7).SetAutoFilter();
        workbook.SaveAs(path);
    }

    private sealed record Identity(string Dni, string Name);

    private static Identity? ReadIdentityFromIndividualPdf(string pdf, string temp, OcrTools tools, CancellationToken ct)
    {
        // Boleta A4 PLAME: en la muestra verificada las coordenadas de ambos
        // campos se mantienen estables. Se escalan proporcionalmente al tamaño de página.
        using var doc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
        var page = doc.Pages[0];
        const int dpi = 300;
        double width = page.Width.Point * dpi / 72.0;
        double height = page.Height.Point * dpi / 72.0;

        // Intento preciso: DNI aprox. x 5%-27%, nombres x 26%-67%, y 25%-31%.
        // Intento amplio: tolera variaciones de escáner y posición.
        foreach (bool wide in new[] { false, true })
        {
            ct.ThrowIfCancellationRequested();
            double y = wide ? .22 : .248;
            double h = wide ? .145 : .062;
            string dniText = OcrRegion(pdf, temp, tools, ct, dpi, width, height,
                wide ? .04 : .055, y, wide ? .31 : .225, h, "dni", wide ? 11 : 6);
            string nameText = OcrRegion(pdf, temp, tools, ct, dpi, width, height,
                wide ? .235 : .26, y, wide ? .55 : .43, h, "nombre", wide ? 11 : 6);
            string? dni = ParseDniRegion(dniText);
            string? name = ParseNameRegion(nameText);
            if (dni is not null && name is not null) return new Identity(dni, name);
        }
        return null;
    }

    private static string OcrRegion(string pdf, string temp, OcrTools tools, CancellationToken ct,
        int dpi, double totalWidth, double totalHeight,
        double xFraction, double yFraction, double wFraction, double hFraction,
        string region, int psm)
    {
        if (tools.Renderer is null || tools.Ocr is null) return "";
        string stem = Path.Combine(temp, Guid.NewGuid().ToString("N") + "_" + region);
        string png = stem + ".png";
        try
        {
            int x = (int)Math.Round(totalWidth * xFraction);
            int y = (int)Math.Round(totalHeight * yFraction);
            int w = Math.Min((int)Math.Round(totalWidth * wFraction), (int)totalWidth - x);
            int h = Math.Min((int)Math.Round(totalHeight * hFraction), (int)totalHeight - y);
            Run(tools.Renderer, $"-f 1 -l 1 -singlefile -r {dpi} -x {x} -y {y} -W {w} -H {h} -png {Q(pdf)} {Q(stem)}", ct);
            if (!File.Exists(png)) throw new IOException("No se generó la imagen de la zona " + region);
            return Run(tools.Ocr, $"{Q(png)} stdout -l {tools.Language} --psm {psm}", ct);
        }
        finally { try { File.Delete(png); } catch { } }
    }

    private static string? ParseDniRegion(string raw)
    {
        // No confundir los 11 dígitos del RUC, números de fecha o de boleta.
        var lines = raw.ToUpperInvariant().Split('\n');
        foreach (string line in lines)
        {
            if (!Regex.IsMatch(line, @"\bD\s*N\s*I\b", RegexOptions.IgnoreCase)) continue;
            var m = Regex.Match(line, @"(?<!\d)\d{8}(?!\d)");
            if (m.Success) return m.Value;
        }
        var matches = Regex.Matches(raw, @"(?<!\d)\d{8}(?!\d)")
            .Select(m => m.Value).Distinct().ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }

    private static string? ParseNameRegion(string raw)
    {
        foreach (var original in raw.Split('\n'))
        {
            string line = Regex.Replace(original.ToUpperInvariant().Normalize(NormalizationForm.FormD), @"[\u0300-\u036f]", "");
            line = Regex.Replace(line, @"\bNOMBRES?\s+Y\s+APELLIDOS?\b", " ");
            line = Regex.Replace(line, @"\b(?:ACTIVO|SUBSIDIADO|SUSPENDIDO|SITUACION|TIPO|TRABAJADOR|REGIMEN|PENSIONARIO|EMPLEADO|OBRERO|AFP|ONP|CUSPP)\b.*", " ");
            line = Regex.Replace(line, @"[^A-ZÑ\s'-]", " ");
            line = Regex.Replace(line, @"\s+", " ").Trim(' ', '-', '\'');
            var words = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length >= 3 && line.Length >= 12 && line.Length <= 90 &&
                words.All(w => w.Length >= 2) &&
                !line.Contains("FECHA DE INGRESO") && !line.Contains("NOMBRES Y APELLIDOS"))
                return line;
        }
        return null;
    }

    /// <summary>
    /// Las boletas PLAME escaneadas repiten su cabecera en la parte superior.
    /// Se limita el OCR al cuadro de identidad, no a los importes ni a las firmas.
    /// Las coordenadas se calculan en función del tamaño A4 y se permite lectura
    /// general en archivos con formatos diferentes.
    /// </summary>
    private sealed record OcrTools(string? Renderer, string? Ocr, string Language);

    private static OcrTools ResolveOcrTools()
    {
        // Primero las herramientas colocadas junto al ejecutable; luego instalaciones de Windows.
        string baseDir = AppContext.BaseDirectory;
        var renderer = LocateExe("pdftoppm.exe", new[] {
            Path.Combine(baseDir, "OCR", "Poppler", "Library", "bin", "pdftoppm.exe"),
            Path.Combine(baseDir, "OCR", "Poppler", "bin", "pdftoppm.exe"),
            Path.Combine(baseDir, "OCR", "pdftoppm.exe"),
            @"C:\poppler\Library\bin\pdftoppm.exe", @"C:\poppler\bin\pdftoppm.exe",
            @"C:\Program Files\poppler\Library\bin\pdftoppm.exe"
        });
        var ocr = LocateExe("tesseract.exe", new[] {
            Path.Combine(baseDir, "OCR", "Tesseract-OCR", "tesseract.exe"),
            Path.Combine(baseDir, "OCR", "tesseract.exe"),
            @"C:\Program Files\Tesseract-OCR\tesseract.exe",
            @"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        });
        string lang = "eng";
        if (ocr is not null)
        {
            try
            {
                var installed = Run(ocr, "--list-langs", CancellationToken.None);
                bool spa = Regex.IsMatch(installed, @"(?m)^spa\s*$", RegexOptions.IgnoreCase);
                bool eng = Regex.IsMatch(installed, @"(?m)^eng\s*$", RegexOptions.IgnoreCase);
                lang = spa && eng ? "spa+eng" : spa ? "spa" : eng ? "eng" : "";
                if (lang.Length == 0) ocr = null;
            }
            catch { ocr = null; }
        }
        return new(renderer, ocr, lang);
    }

    private static string? LocateExe(string name, IEnumerable<string> candidates)
    {
        foreach (var path in candidates)
            if (File.Exists(path)) return path;
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            try { var path = Path.Combine(dir.Trim('"'), name); if (File.Exists(path)) return path; }
            catch { }
        }
        // WinGet puede instalar Poppler en paquetes locales cuyo directorio cambia con la version.
        var packages = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "WinGet", "Packages");
        if (Directory.Exists(packages))
        {
            try
            {
                foreach (var folder in Directory.EnumerateDirectories(packages, "*Poppler*"))
                {
                    foreach (var path in Directory.EnumerateFiles(folder, name, SearchOption.AllDirectories))
                        if (File.Exists(path)) return path;
                }
            }
            catch { }
        }
        return RequirementsService.Find(name);
    }

    private static string Run(string exe, string args, CancellationToken ct)
    {
        var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true };
        using var process = Process.Start(psi) ?? throw new IOException("No se inició OCR: " + exe);
        using var registration = ct.Register(() => { try { if (!process.HasExited) process.Kill(true); } catch { } });
        string output = process.StandardOutput.ReadToEnd();
        string stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (ct.IsCancellationRequested) ct.ThrowIfCancellationRequested();
        if (process.ExitCode != 0) throw new IOException(exe + ": " + stderr);
        return output;
    }
    private static string Q(string s) => "\"" + s.Replace("\"", "\\\"") + "\"";

    private static string Row(params object[] fields) => string.Join(";", fields.Select(v => "\"" + v.ToString()!.Replace("\"", "\"\"") + "\""));
    private static string SafeFile(string s)
    {
        var clean = new string(s.Select(c => Path.GetInvalidFileNameChars().Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim(' ', '.');
        return clean[..Math.Min(clean.Length, 140)];
    }
    private static string UniquePath(string p)
    {
        if (!File.Exists(p)) return p;
        for (var n = 2; n < 100000; n++)
        {
            string candidate = Path.Combine(Path.GetDirectoryName(p)!, Path.GetFileNameWithoutExtension(p) + "_" + n + Path.GetExtension(p));
            if (!File.Exists(candidate)) return candidate;
        }
        throw new IOException("Demasiados nombres repetidos de boletas.");
    }
}
