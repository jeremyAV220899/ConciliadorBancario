using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SharpCompress.Archives;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace ConciliadorBancario;

internal sealed record PaymentArchive(string Payment, string Archive);
internal sealed record PreparationProgress(int Percentage, string Message);
internal sealed record PreparationResult(string OutputFolder, int PdfCount, int IndividualCount, int ReviewCount, string ZipPath);
internal sealed record Worker(string Dni, string Name, decimal? Amount, int ExcelRow);
internal sealed record PageFrame(int Page, Box Frame, double BorderWidth);
internal sealed record Match(Worker Worker, string Pdf, string Bank, int Page, Box Frame, double BorderWidth, Box DniBox, Box? AmountBox, decimal? BankAmount, string BankStatus, string Payment)
{
    public IReadOnlyList<PageFrame> Continuations { get; init; } = Array.Empty<PageFrame>();
}
internal sealed record Box(double Left, double Top, double Right, double Bottom)
{
    public double Width => Right - Left;
    public double Height => Bottom - Top;
    public static Box Union(IEnumerable<Box> boxes) => new(boxes.Min(x => x.Left), boxes.Min(x => x.Top), boxes.Max(x => x.Right), boxes.Max(x => x.Bottom));
    public bool Overlaps(Box b, double gap = 0) => Left < b.Right + gap && Right > b.Left - gap && Top < b.Bottom + gap && Bottom > b.Top - gap;
}
internal sealed record Token(string Text, Box Bounds, double MidY);
internal sealed record PdfAudit(string File, string Bank, int Pages, int TextTokens, int DniOccurrences, int DniInExcel, int Generated, int Reviews, string Status, string Payment);

internal static class ProcessingEngine
{
    private static readonly Regex DniRegex = new(@"(?<!\d)\d{8}(?!\d)", RegexOptions.Compiled);
    private static readonly Regex NumericRegex = new(@"[^\d]", RegexOptions.Compiled);
    private const long PdfLimit = 200L * 1024 * 1024;
    private const long TotalLimit = 3L * 1024 * 1024 * 1024;

    public static PreparationResult Process(IReadOnlyList<PaymentArchive> selectedArchives, string excelPath, string destination,
        IProgress<PreparationProgress> progress, CancellationToken ct, string? payflowPath = null, string? ctrPath = null, string? vidaLeyPath = null, string? boletasPath = null, IReadOnlyList<string>? documentOrder = null)
    {
        var archives = selectedArchives.Where(x => !string.IsNullOrWhiteSpace(x.Archive)).ToList();
        if (archives.Count == 0 && string.IsNullOrWhiteSpace(payflowPath) && string.IsNullOrWhiteSpace(ctrPath) && string.IsNullOrWhiteSpace(vidaLeyPath) && string.IsNullOrWhiteSpace(boletasPath))
            throw new InvalidDataException("Seleccione al menos un archivo de pagos, PAYFLOW, CTR, VIDA LEY o BOLETAS.");
        if (archives.GroupBy(x => x.Payment, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
            throw new InvalidDataException("Cada tipo de pago debe incluirse una sola vez.");
        foreach (var archive in archives)
            if (!File.Exists(archive.Archive)) throw new FileNotFoundException($"No existe el ZIP/RAR de {archive.Payment}.", archive.Archive);
        // Solo BOLETAS no requiere Excel; otros módulos conservan la lista obligatoria.
        bool needsExcel = archives.Count > 0 || !string.IsNullOrWhiteSpace(ctrPath) || !string.IsNullOrWhiteSpace(vidaLeyPath);
        if (needsExcel && !File.Exists(excelPath)) throw new FileNotFoundException("No existe el Excel seleccionado.", excelPath);
        var workers = File.Exists(excelPath) ? ReadWorkers(excelPath) : new List<Worker>();
        if (needsExcel && workers.Count == 0) throw new InvalidDataException("No se detectaron DNI de 8 dígitos en el Excel. Debe existir una columna DNI/documento y una de nombres.");
        var workerIndex = workers.GroupBy(x => x.Dni).ToDictionary(x => x.Key, x => x.ToList());
        Directory.CreateDirectory(destination);
        string root = Path.Combine(destination, "CONCILIACION_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"));
        string general = Path.Combine(root, "01_PDF_BANCARIOS_SUBRAYADOS");
        string individual = Path.Combine(root, "02_PDF_INDIVIDUALES");
        string staging = Path.Combine(Path.GetTempPath(), "CONCILIADOR_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(general);
        Directory.CreateDirectory(individual);
        Directory.CreateDirectory(staging);
        foreach (var item in archives)
        {
            string paymentFolder = Path.Combine(individual, item.Payment);
            foreach (var bankName in new[] { "GENERAL", "BBVA", "BCP", "INTERBANK", "SCOTIABANK", "BANCO_NO_IDENTIFICADO" })
                Directory.CreateDirectory(Path.Combine(paymentFolder, bankName));
            Directory.CreateDirectory(Path.Combine(general, item.Payment));
        }
        var matches = new List<Match>();
        var audits = new List<PdfAudit>();
        var reviews = new List<(string dni, string pdf, int page, string reason)>();
        var seenRecords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int pdfCount = 0, individualCount = 0;
        try
        {
            var jobs = new List<(string Payment, string ArchivePdf)>();
            foreach (var item in archives)
            {
                string work = Path.Combine(staging, item.Payment);
                Directory.CreateDirectory(work);
                foreach (var pdf in ExtractPdfs(item.Archive, work, ct)) jobs.Add((item.Payment, pdf));
            }
            progress.Report(new(8, $"Excel: {workers.Count} registros; PDF encontrados: {jobs.Count}; pagos: {archives.Count}."));
            for (int i = 0; i < jobs.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var (payment, path) = jobs[i];
                string paymentGeneral = Path.Combine(general, payment);
                string paymentIndividual = Path.Combine(individual, payment);
                string bank = DetectBank(path);
                var fileMatches = new List<Match>();
                var pageTokens = new Dictionary<int, List<Token>>();
                var pageSizes = new Dictionary<int, (double Width, double Height)>();
                int pageCount = 0, textTokenCount = 0, dniFound = 0, dniInExcel = 0, filesBefore = individualCount, reviewsBefore = reviews.Count;
                Box? paymentDateFrame = null;
                var safeName = SafeFile(Path.GetFileName(path));
                try
                {
                    using (var doc = UglyToad.PdfPig.PdfDocument.Open(path))
                    {
                        pageCount = doc.NumberOfPages;
                        if (bank == "BANCO_NO_IDENTIFICADO" && doc.NumberOfPages > 0)
                            bank = DetectBankFromText(doc.GetPage(1).Text);
                        for (int pageNo = 1; pageNo <= doc.NumberOfPages; pageNo++)
                        {
                            ct.ThrowIfCancellationRequested();
                            var page = doc.GetPage(pageNo);
                            var tokens = NearestNeighbourWordExtractor.Instance.GetWords(page.Letters)
                                .Select(w => new Token(w.Text, new Box(w.BoundingBox.Left, page.Height - w.BoundingBox.Top,
                                    w.BoundingBox.Right, page.Height - w.BoundingBox.Bottom),
                                    page.Height - (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2)).ToList();
                            textTokenCount += tokens.Count;
                            if (pageNo == 1)
                            {
                                paymentDateFrame = FindBankPaymentDateFrame(tokens, bank, page.Width, page.Height);
                                if (paymentDateFrame is not null)
                                    progress.Report(new(8 + (int)(84.0 * (i + 1) / Math.Max(1, jobs.Count)),
                                        $"FECHA UBICADA: {bank} ({payment}), página 1, X={paymentDateFrame.Left:F1}, Y={paymentDateFrame.Top:F1}."));
                                if (paymentDateFrame is null)
                                    progress.Report(new(8 + (int)(84.0 * (i + 1) / Math.Max(1, jobs.Count)),
                                        $"AVISO: {bank} ({payment}): fecha de pago sin ubicación inequívoca, no se dibujará marca de fecha."));
                            }
                            if (bank == "BBVA")
                            {
                                pageTokens[pageNo] = tokens;
                                pageSizes[pageNo] = (page.Width, page.Height);
                            }
                            var dniWords = FindDniTokens(tokens, bank);
                            dniFound += dniWords.Count;
                            foreach (var occurrence in dniWords)
                            {
                                string dni = DniRegex.Match(occurrence.Text).Value;
                                if (!workerIndex.TryGetValue(dni, out var candidates)) continue;
                                dniInExcel++;
                                string recordKey = $"{path}|{pageNo}|{dni}|{Math.Round(occurrence.MidY, 1)}";
                                if (!seenRecords.Add(recordKey)) continue;
                                var frame = CalculateSafeFrame(tokens, occurrence, dniWords, bank, page.Width, page.Height, out var borderWidth, out var reason);
                                if (frame is null)
                                {
                                    reviews.Add((dni, path, pageNo, reason));
                                    continue;
                                }
                                // If the same DNI appears twice in one physical region, prefer a single unambiguous Excel row.
                                if (candidates.Count != 1)
                                {
                                    reviews.Add((dni, path, pageNo, "DNI duplicado en Excel; seleccionar fila manualmente"));
                                    continue;
                                }
                                var worker = candidates[0];
                                var rowTokens = tokens.Where(t => IsWithin(t.Bounds, frame)).ToList();
                                var amountToken = FindBankAmount(rowTokens, bank, page.Width);
                                decimal? bankAmount = amountToken is null ? null : ParseAmount(amountToken.Text);
                                string state = FindBankStatus(rowTokens, bank);
                                if (bankAmount is null || state == "NO VERIFICADO")
                                {
                                    reviews.Add((dni, path, pageNo, "Importe o estado bancario sin lectura segura; verificar comprobante"));
                                }
                                fileMatches.Add(new Match(worker, path, bank, pageNo, frame, borderWidth, occurrence.Bounds,
                                    amountToken?.Bounds, bankAmount, state, payment));
                            }
                        }
                    }
                    // V11: en BBVA ciertos nombres continúan en una página distinta
                    // de la que contiene su DNI. Inspeccionar únicamente páginas
                    // ADYACENTES y añadir fragmentos vinculados al nombre del Excel.
                    if (bank == "BBVA" && fileMatches.Count > 0)
                    {
                        for (int j = 0; j < fileMatches.Count; j++)
                        {
                            var item = fileMatches[j];
                            var continuation = FindBbvaPageContinuations(item, pageTokens, pageSizes);
                            if (continuation.Count > 0)
                            {
                                fileMatches[j] = item with { Continuations = continuation };
                                progress.Report(new(8 + (int)(84.0 * (i + 1) / Math.Max(1, jobs.Count)),
                                    $"BBVA: DNI {item.Worker.Dni}: {continuation.Count} fragmento(s) recuperado(s) de páginas adyacentes."));
                            }
                        }
                    }
                    // Produce annotated full bank PDF, preserving original page order.
                    var destGeneral = UniquePath(Path.Combine(paymentGeneral, safeName));
                    try
                    {
                        if (fileMatches.Count == 0) File.Copy(path, destGeneral);
                        else WriteGeneral(path, destGeneral, fileMatches);
                    }
                    catch (Exception ex)
                    {
                        reviews.Add(("", path, 0, "No se pudo subrayar PDF general: " + ex.Message));
                        if (File.Exists(destGeneral)) File.Delete(destGeneral);
                        File.Copy(path, destGeneral);
                    }
                    foreach (var m in fileMatches)
                    {
                        ct.ThrowIfCancellationRequested();
                        try
                        {
                            string filename = $"{m.Worker.Dni}_{SafeFile(m.Worker.Name)}_{SafeFile(m.Bank)}.pdf";
                            string bankFolder = Path.Combine(paymentIndividual, SafeFile(m.Bank));
                            Directory.CreateDirectory(bankFolder);
                            string output = UniquePath(Path.Combine(bankFolder, filename));
                            WriteIndividual(path, output, m, paymentDateFrame);
                            // GENERAL contiene todos los comprobantes individuales del pago, sin cambiar el original.
                            string generalCopy = UniquePath(Path.Combine(paymentIndividual, "GENERAL", filename));
                            File.Copy(output, generalCopy);
                            matches.Add(m);
                            individualCount++;
                        }
                        catch (Exception ex)
                        {
                            reviews.Add((m.Worker.Dni, path, m.Page, "No se pudo crear PDF individual: " + ex.Message));
                        }
                    }
                    pdfCount++;
                    string status = individualCount > filesBefore ? "GENERADO" : textTokenCount == 0 ? "PDF SIN TEXTO (ESCANEADO)" :
                        dniFound == 0 ? "SIN DNI LEGIBLE" : dniInExcel == 0 ? "SIN COINCIDENCIAS CON EXCEL" : "COINCIDENCIAS EN REVISIÓN";
                    audits.Add(new PdfAudit(Path.GetFileName(path), bank, pageCount, textTokenCount, dniFound, dniInExcel, individualCount - filesBefore, reviews.Count - reviewsBefore, status, payment));
                    progress.Report(new(8 + (int)(84.0 * (i + 1) / jobs.Count),
                        $"{i + 1}/{jobs.Count} | {payment} | {bank} | DNI leídos: {dniFound} | Excel: {dniInExcel} | PDF generados: {individualCount-filesBefore} | {status} | {Path.GetFileName(path)}"));
                }
                catch (Exception ex)
                {
                    reviews.Add(("", path, 0, "No fue posible leer el PDF (escaneado, cifrado o dañado): " + ex.Message));
                    // Preserve the original even if its content cannot be parsed.
                    File.Copy(path, UniquePath(Path.Combine(paymentGeneral, safeName)));
                    audits.Add(new PdfAudit(Path.GetFileName(path), bank, pageCount, textTokenCount, dniFound, dniInExcel, 0, reviews.Count - reviewsBefore, "ERROR DE LECTURA", payment));
                    progress.Report(new(8 + (int)(84.0 * (i + 1) / jobs.Count), "REVISIÓN: " + Path.GetFileName(path) + " — " + ex.Message));
                }
            }
            if (!string.IsNullOrWhiteSpace(payflowPath))
            {
                ct.ThrowIfCancellationRequested();
                var payflow = PayflowProcessor.Generate(payflowPath, root, progress, ct);
                individualCount += payflow.IndividualPdfs;
                reviews.AddRange(Enumerable.Range(0, payflow.Warnings).Select(_ => ("", payflowPath!, 0, "Advertencia de PAYFLOW; consultar PAYFLOW_ADVERTENCIAS.txt")));
                pdfCount++; // PDF general PAYFLOW; individuales se cuentan por separado.
            }
            if (!string.IsNullOrWhiteSpace(ctrPath))
            {
                ct.ThrowIfCancellationRequested();
                var ctr = CtrProcessor.Generate(ctrPath, workers, root, progress, ct);
                pdfCount++; // CTR: un solo PDF íntegro, jamás individuales.
                if (ctr.Unmatched > 0)
                    reviews.Add(("", ctrPath!, 0, $"CTR: {ctr.Unmatched} documentos del Excel sin coincidencia; ver CTR_REPORTE.txt"));
            }
            if (!string.IsNullOrWhiteSpace(vidaLeyPath))
            {
                ct.ThrowIfCancellationRequested();
                var vidaLey = CtrProcessor.Generate(vidaLeyPath, workers, root, progress, ct, "VIDA_LEY");
                pdfCount++; // VIDA LEY: un PDF íntegro enmarcado, sin individuales.
                if (vidaLey.Unmatched > 0)
                    reviews.Add(("", vidaLeyPath!, 0, $"VIDA LEY: {vidaLey.Unmatched} documentos del Excel sin coincidencia; ver VIDA_LEY_REPORTE.txt"));
            }
            if (!string.IsNullOrWhiteSpace(boletasPath))
            {
                ct.ThrowIfCancellationRequested();
                var boletas = BoletasProcessor.Generate(boletasPath, root, progress, ct);
                pdfCount += boletas.Sources;
                individualCount += boletas.Created;
                if (boletas.Reviews > 0)
                    reviews.Add(("", boletasPath!, 0, $"BOLETAS: {boletas.Reviews} páginas requieren revisión; ver 05_BOLETAS/BOLETAS_REPORTE.xlsx"));
            }
            // Primero terminar todos los módulos; luego comprobar documentos realmente producidos.
            string summaryPath = Path.Combine(root, "RESUMEN_CONCILIACION_PAGOS.xlsx");
            WriteSummary(summaryPath, workers, matches, reviews, audits, archives.Select(a => a.Payment).ToList());
            DocumentControl.AppendControl(summaryPath, root, workers, archives.Select(a => a.Payment).ToList(),
                !string.IsNullOrWhiteSpace(payflowPath), !string.IsNullOrWhiteSpace(ctrPath),
                !string.IsNullOrWhiteSpace(vidaLeyPath), !string.IsNullOrWhiteSpace(boletasPath), progress, ct);
            // ETAPA FINAL V31: un PDF único en el orden físico del Excel, luego comprimir.
            var consolidated = ConsolidatedPdfBuilder.Generate(root, workers,
                !string.IsNullOrWhiteSpace(ctrPath), !string.IsNullOrWhiteSpace(vidaLeyPath),
                !string.IsNullOrWhiteSpace(boletasPath), !string.IsNullOrWhiteSpace(payflowPath),
                archives.Select(a => a.Payment).ToList(), documentOrder ?? DocumentOrderSettings.Load(), progress, ct);
            if (consolidated.Pages > 0) pdfCount++;
            string zipPath = UniquePath(Path.Combine(destination, "CONCILIACION_PAGOS_DNI.zip"));
            ZipFile.CreateFromDirectory(root, zipPath, CompressionLevel.Optimal, false);
            progress.Report(new(100, $"Finalizado: {individualCount} PDF individuales; {reviews.Count} incidencias para revisar."));
            return new PreparationResult(root, pdfCount, individualCount, reviews.Count, zipPath);
        }
        finally
        {
            try { Directory.Delete(staging, true); } catch { /* temp may be locked */ }
        }
    }

    private static List<Worker> ReadWorkers(string path)
    {
        if (Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Guarde el Excel antiguo .xls como .xlsx antes de importarlo.");
        var result = new List<Worker>();
        using var wb = new XLWorkbook(path);
        foreach (var sheet in wb.Worksheets)
        {
            var rows = sheet.RowsUsed().Take(200000).ToList();
            var header = rows.Take(25).Select(r => new {
                Row = r,
                Cells = r.CellsUsed().ToDictionary(c => c.Address.ColumnNumber,
                    c => NormalizeHeader(c.GetFormattedString()))
            }).Select(x => new {
                x.Row,
                Dni = x.Cells.FirstOrDefault(c => c.Value is "dni" or "documento" or "numerodedocumento" or "numerodni" or "nrodocumento" or "docidentidad" or "identificacion").Key,
                Name = x.Cells.FirstOrDefault(c => c.Value.Contains("nombre") || c.Value.Contains("apellidos") || c.Value == "trabajador" || c.Value == "beneficiario").Key,
                Amount = x.Cells.FirstOrDefault(c => c.Value.Contains("monto") || c.Value.Contains("importe") || c.Value.Contains("netopagar") || c.Value == "pago").Key
            }).FirstOrDefault(x => x.Dni > 0 && x.Name > 0);
            if (header is null) continue;
            foreach (var row in rows.Where(r => r.RowNumber() > header.Row.RowNumber()))
            {
                var dniCell = row.Cell(header.Dni);
                string dni = NumericRegex.Replace(dniCell.GetFormattedString(), "");
                if (dni.Length < 8 && dni.Length > 0 && dniCell.TryGetValue<long>(out var numericDni) && numericDni >= 0)
                    dni = numericDni.ToString("D8", CultureInfo.InvariantCulture);
                if (dni.Length != 8) continue;
                string name = row.Cell(header.Name).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(name)) name = "SIN_NOMBRE";
                decimal? amount = null;
                if (header.Amount > 0)
                {
                    var cell = row.Cell(header.Amount);
                    if (cell.TryGetValue<decimal>(out var number)) amount = number;
                    else if (decimal.TryParse(cell.GetFormattedString().Replace("S/", "").Replace(",", ""),
                                 NumberStyles.Number, CultureInfo.InvariantCulture, out number)) amount = number;
                }
                result.Add(new(dni, name, amount, row.RowNumber()));
            }
        }
        return result;
    }

    private static string NormalizeHeader(string s) => new string(s.Normalize(NormalizationForm.FormD)
        .Where(ch => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
        .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static List<string> ExtractPdfs(string archivePath, string output, CancellationToken ct)
    {
        var found = new List<string>();
        using var archive = ArchiveFactory.Open(archivePath);
        long cumulative = 0;
        int index = 0;
        foreach (var entry in archive.Entries)
        {
            ct.ThrowIfCancellationRequested();
            if (entry.IsDirectory || !entry.Key.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) continue;
            if (++index > 20000) throw new InvalidDataException("Más de 20 000 PDF dentro del archivo.");
            if (entry.Size > PdfLimit) throw new InvalidDataException("Un PDF supera el límite de 200 MB: " + entry.Key);
            string bankName = SafeFile(entry.Key.Replace('\\', '/').Replace('/', '_'));
            string bankLabel = DetectBank(entry.Key);
            string dest = Path.Combine(output, $"{index:00000}_{bankLabel}_{bankName}");
            using (var input = entry.OpenEntryStream())
            using (var file = new FileStream(dest, FileMode.CreateNew))
            {
                byte[] buff = new byte[65536];
                long length = 0;
                int n;
                while ((n = input.Read(buff, 0, buff.Length)) > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    length += n;
                    cumulative += n;
                    if (length > PdfLimit || cumulative > TotalLimit) throw new InvalidDataException("Archivo comprimido demasiado grande.");
                    file.Write(buff, 0, n);
                }
            }
            using var stream = File.OpenRead(dest);
            byte[] first = new byte[5];
            if (stream.Read(first) != 5 || Encoding.ASCII.GetString(first) != "%PDF-")
                throw new InvalidDataException("Archivo .pdf sin firma PDF válida: " + entry.Key);
            found.Add(dest);
        }
        if (found.Count == 0) throw new InvalidDataException("El ZIP/RAR no contiene archivos PDF.");
        return found;
    }

    // Plantillas de ubicación por banco, basadas en los PDF originales recibidos.
    // No son coordenadas absolutas: las bandas se escalan al ancho de la página.
    // El filtro de columna evita identificar cuentas como DNI y el duplicado "Ref Emp" de BCP.
    private static List<Token> FindDniTokens(List<Token> tokens, string bank)
    {
        var result = new List<Token>();
        var seen = new HashSet<string>();
        double pageWidth = tokens.Select(t => t.Bounds.Right).DefaultIfEmpty(595).Max();
        // Tamaños base: ancho real del contenido puede ser inferior al MediaBox.
        (double lo, double hi) = bank switch
        {
            "BBVA" => (0.63, 0.81),
            "BCP" => (0.19, 0.36),
            "INTERBANK" => (0.045, 0.23),
            "SCOTIABANK" => (0.035, 0.19),
            _ => (0, 1)
        };
        // El PDF puede tener márgenes extra o rotación; tolerar una anchura base mínima.
        double scaleWidth = Math.Max(pageWidth, bank == "BBVA" ? 585 : 555);
        bool InDniColumn(Token t) => t.Bounds.Left >= lo * scaleWidth && t.Bounds.Left <= hi * scaleWidth;
        foreach (var t in tokens)
        {
            var match = DniRegex.Match(t.Text);
            if (match.Success && InDniColumn(t)) Add(t, match.Value);
        }
        var lines = tokens.GroupBy(t => (int)Math.Round(t.MidY / 3.0));
        foreach (var line in lines)
        {
            var words = line.OrderBy(t => t.Bounds.Left).ToList();
            for (int i = 0; i < words.Count; i++)
            {
                if (!Regex.IsMatch(words[i].Text, @"^\d{1,7}$") || !InDniColumn(words[i])) continue;
                string digits = "";
                var parts = new List<Token>();
                for (int j = i; j < Math.Min(i + 5, words.Count); j++)
                {
                    var next = words[j];
                    if (!Regex.IsMatch(next.Text, @"^\d{1,7}$")) break;
                    if (parts.Count > 0 && (next.Bounds.Left - parts[^1].Bounds.Right > 8 ||
                        Math.Abs(next.MidY - words[i].MidY) > 2.5)) break;
                    digits += next.Text;
                    parts.Add(next);
                    if (digits.Length == 8)
                    {
                        var bounds = Box.Union(parts.Select(x => x.Bounds));
                        Add(new Token(digits, bounds, (bounds.Top + bounds.Bottom) / 2), digits);
                        break;
                    }
                    if (digits.Length > 8) break;
                }
            }
        }
        return result.OrderBy(x => x.MidY).ToList();
        void Add(Token token, string dni)
        {
            string key = dni + "|" + Math.Round(token.MidY / 2.0);
            if (seen.Add(key)) result.Add(token);
        }
    }

    // V5: buscar límites a partir de los textos completos, no únicamente del centro del DNI.
    // Nunca se fuerza un marco cuya tinta pudiera tocar otra palabra.
    private static Box? CalculateSafeFrame(List<Token> tokens, Token dni, List<Token> anchors,
        string bank, double pageWidth, double pageHeight, out double borderWidth, out string reason)
    {
        reason = "";
        borderWidth = 1.35;
        // Interbank: el DNI se imprime DEBAJO del beneficiario. Un intervalo
        // construido a partir de los centros de dos DNI no describe sus filas.
        // Delimitamos por el comienzo real del nombre propio y el comienzo del
        // siguiente beneficiario, y después verificamos todos los bordes.
        if (bank == "INTERBANK")
            return CalculateInterbankFrame(tokens, dni, anchors, pageWidth, pageHeight,
                out borderWidth, out reason);
        // V9: BBVA puede imprimir 3-5 renglones de titular. El punto medio
        // entre los DNI NO es necesariamente la separación real de dos filas:
        // p. ej. el titular de la fila 116 termina muy por debajo de su DNI.
        // Primero intentar una banda adaptativa y validar todos sus bordes.
        if (bank == "BBVA")
        {
            var bbvaFrame = CalculateBbvaFrame(tokens, dni, anchors, pageWidth, pageHeight,
                out borderWidth, out reason);
            if (bbvaFrame is not null) return bbvaFrame;
        }
        // V8: BCP imprime el mismo DNI en "Documento" y "Ref Emp".
        // No son trabajadores diferentes: la fila se delimita por el N° de registro.
        if (bank == "BCP")
            return CalculateBcpFrame(tokens, dni, anchors, pageWidth, pageHeight,
                out borderWidth, out reason);
        var ordered = anchors.OrderBy(a => a.MidY).ToList();
        int idx = ordered.FindIndex(a => ReferenceEquals(a, dni));
        if (idx < 0) { reason = "DNI sin ancla única"; return null; }

        var prev = idx > 0 ? ordered[idx - 1] : null;
        var next = idx + 1 < ordered.Count ? ordered[idx + 1] : null;
        // Los intervalos se derivan de las anclas reales de los registros próximos.
        // Interbank imprime el DNI debajo del nombre: su fila empieza antes del DNI.
        double head = bank switch { "BBVA" => 26, "INTERBANK" => 35, "BCP" => 15, _ => 17 };
        double tail = bank switch { "BBVA" => 27, "INTERBANK" => 8, "BCP" => 15, _ => 17 };
        double upper = prev is null ? dni.MidY - head :
            bank == "INTERBANK" ? Math.Max(prev.MidY + 4, (prev.MidY + dni.MidY) / 2 - 11) :
            (prev.MidY + dni.MidY) / 2;
        double lower = next is null ? dni.MidY + tail :
            bank == "INTERBANK" ? Math.Min(next.MidY - 12, (dni.MidY + next.MidY) / 2 - 11) :
            (dni.MidY + next.MidY) / 2;
        upper = Math.Max(0, upper);
        lower = Math.Min(pageHeight, lower);
        if (lower - upper < 9 || lower - upper > 135)
        { reason = "Intervalo vertical ambiguo entre registros"; return null; }

        // Retener únicamente palabras completas dentro de la banda de la fila.
        // Esto evita que un fragmento de la fila anterior o siguiente se convierta
        // accidentalmente en parte del rectángulo de este trabajador.
        var row = tokens.Where(t => t.Bounds.Top >= upper + 0.1 &&
                                    t.Bounds.Bottom <= lower - 0.1 &&
                                    t.Bounds.Left >= 0 && t.Bounds.Right <= pageWidth).ToList();
        if (!row.Any(t => ReferenceEquals(t, dni))) row.Add(dni);
        if (row.Count < 3)
        { reason = "Fila sin importe/nombre/estado suficiente"; return null; }
        var content = Box.Union(row.Select(t => t.Bounds));
        if (content.Width < 70 || content.Height > 120)
        { reason = "Extensión de fila inválida"; return null; }

        // No reducir el contenido para que quepa el borde. Primero se intenta
        // margen 3.5pt y grosor 1.35pt; luego se reducen progresivamente.
        // Además, cada candidato comprueba intersección de los cuatro trazos
        // con TODOS los tokens, incluidos los vecinos y los encabezados.
        foreach (double stroke in new[] { 1.35, 1.0, 0.7, 0.45 })
        {
            foreach (double margin in new[] { 3.5, 2.0, 1.1, 0.65, 0.35 })
            {
                double edge = stroke / 2 + 0.18;
                var frame = new Box(
                    Math.Max(edge, content.Left - margin),
                    Math.Max(edge, content.Top - margin),
                    Math.Min(pageWidth - edge, content.Right + margin),
                    Math.Min(pageHeight - edge, content.Bottom + margin));
                if (frame.Width <= 0 || frame.Height <= 0) continue;
                // La línea física ocupa stroke/2 a cada lado de la coordenada.
                double clearance = stroke / 2 + 0.12;
                if (content.Left - frame.Left <= clearance ||
                    frame.Right - content.Right <= clearance ||
                    content.Top - frame.Top <= clearance ||
                    frame.Bottom - content.Bottom <= clearance) continue;

                // Ningún DNI ajeno puede quedar encerrado (también se controla
                // el caso de las filas altas de BBVA).
                if (ordered.Any(a => !ReferenceEquals(a, dni) && a.Bounds.Overlaps(frame, stroke / 2)))
                    continue;
                bool safe = true;
                foreach (var token in tokens)
                {
                    var t = token.Bounds;
                    // Si está totalmente dentro, no queda debajo de ningún borde.
                    bool inside = t.Left > frame.Left + clearance &&
                                  t.Right < frame.Right - clearance &&
                                  t.Top > frame.Top + clearance &&
                                  t.Bottom < frame.Bottom - clearance;
                    if (inside) continue;
                    // Detectar trazos que crucen tinta (no confundir con que
                    // el rectángulo completo abarque el área de texto).
                    var topEdge = new Box(frame.Left - clearance, frame.Top - clearance,
                                          frame.Right + clearance, frame.Top + clearance);
                    var bottomEdge = new Box(frame.Left - clearance, frame.Bottom - clearance,
                                             frame.Right + clearance, frame.Bottom + clearance);
                    var leftEdge = new Box(frame.Left - clearance, frame.Top - clearance,
                                           frame.Left + clearance, frame.Bottom + clearance);
                    var rightEdge = new Box(frame.Right - clearance, frame.Top - clearance,
                                            frame.Right + clearance, frame.Bottom + clearance);
                    if (t.Overlaps(topEdge) || t.Overlaps(bottomEdge) ||
                        t.Overlaps(leftEdge) || t.Overlaps(rightEdge))
                    { safe = false; break; }
                }
                if (!safe) continue;
                borderWidth = stroke;
                return frame;
            }
        }
        reason = "Fila sin espacio suficiente para un borde libre de texto (todos los grosores/márgenes comprobados)";
        return null;
    }

    // V9: BBVA. Agrupar cada registro mediante las anclas de DNI en la
    // columna Doc.Identidad y el número de fila en la columna No.
    // La banda se desplaza respecto del punto medio entre DNI porque los
    // nombres de 3-5 líneas pueden extenderse bajo el documento.
    // Todos los candidatos pasan por TrySafeBorder, que examina los trazos
    // contra TODOS los textos de la página y evita DNI de otras filas.
    private static Box? CalculateBbvaFrame(List<Token> tokens, Token dni,
        List<Token> anchors, double pageWidth, double pageHeight,
        out double borderWidth, out string reason)
    {
        borderWidth = 1.35;
        reason = "BBVA: no se pudo delimitar con seguridad la fila completa";
        var ordered = anchors.OrderBy(t => t.MidY).ToList();
        int i = ordered.FindIndex(t => ReferenceEquals(t, dni));
        if (i < 0) return null;
        var prev = i > 0 ? ordered[i - 1] : null;
        var next = i + 1 < ordered.Count ? ordered[i + 1] : null;
        // Las primeras/últimas filas de página también se procesan con
        // bandas reales; no se envían al estimador de puntos medios.
        double before = prev is null ? Math.Min(65, Math.Max(15, dni.MidY - 28)) : dni.MidY - prev.MidY;
        double after = next is null ? Math.Min(65, Math.Max(18, pageHeight - dni.MidY - 12)) : next.MidY - dni.MidY;
        if (before < 12 || after < 12 || before > 110 || after > 110)
            return null;
        // Probar primero franjas estrechas, expandiéndolas lo necesario.
        // El límite de la fila 116 (DNI 25633881) queda aprox. entre
        // 600.4 y 661.1 pt, no entre los puntos medios de los DNI.
        foreach (double fraction in new[] { 0.62, 0.65, 0.69, 0.73, 0.77 })
        {
            double top = prev is null ? Math.Max(0, dni.MidY - before) :
                Math.Max(0, dni.MidY - before * fraction);
            double bottom = next is null ? Math.Min(pageHeight - 2, dni.MidY + after) :
                Math.Min(pageHeight, dni.MidY + after * fraction);
            if (bottom - top > 120) continue;
            var row = tokens.Where(t => t.Bounds.Left >= pageWidth * 0.09 &&
                t.Bounds.Right <= pageWidth * 0.955 &&
                t.Bounds.Top >= top && t.Bounds.Bottom <= bottom).ToList();
            if (!row.Any(t => ReferenceEquals(t, dni))) row.Add(dni);
            // Exigir el número correlativo, el nombre repartido en columnas,
            // el importe y el estado de la operación.
            bool hasIndex = row.Any(t =>
                t.Bounds.Left >= pageWidth * 0.11 && t.Bounds.Left <= pageWidth * 0.16 &&
                Regex.IsMatch(t.Text, @"^\d{1,4}$") &&
                Math.Abs(t.MidY - dni.MidY) < 4);
            bool hasNames = row.Count(t => t.Bounds.Left >= pageWidth * 0.38 &&
                t.Bounds.Left < pageWidth * 0.65 && t.Text.Any(char.IsLetter)) >= 3;
            if (!hasIndex || !hasNames || FindBankAmount(row, "BBVA", pageWidth) is null ||
                FindBankStatus(row, "BBVA") != "ABONO CORRECTO") continue;
            var content = Box.Union(row.Select(t => t.Bounds));
            if (content.Width < 250 || content.Height < 13 || content.Height > 105)
                continue;
            var candidate = TrySafeBorder(tokens, ordered, dni, content, bottom,
                pageWidth, pageHeight, out borderWidth, out reason);
            if (candidate is not null) return candidate;
        }
        reason = "BBVA: fila multilínea sin borde seguro; revisar delimitación";
        return null;
    }

    // V8: BCP Telecrédito. El número del registro, en la primera columna,
    // delimita horizontalmente la fila completa y evita el falso vecino que
    // producía el DNI duplicado en la referencia del empleador.
    private static Box? CalculateBcpFrame(List<Token> tokens, Token dni,
        List<Token> anchors, double pageWidth, double pageHeight,
        out double borderWidth, out string reason)
    {
        borderWidth = 1.35;
        reason = "";
        // Telecrédito utiliza índices consecutivos en la columna izquierda.
        // Comprobar alineación horizontal con el DNI de Documento/Referencia.
        var numbers = tokens.Where(t =>
            t.Bounds.Left >= pageWidth * 0.047 && t.Bounds.Right <= pageWidth * 0.095 &&
            Regex.IsMatch(t.Text.Trim(), @"^\d{1,3}$") &&
            t.Bounds.Top >= 70)
            .OrderBy(t => t.Bounds.Top).ToList();
        var current = numbers.Where(t => Math.Abs(t.Bounds.Top - dni.Bounds.Top) <= 3.0)
            .OrderBy(t => Math.Abs(t.Bounds.Top - dni.Bounds.Top)).FirstOrDefault();
        if (current is null)
        {
            reason = "BCP: no se localizó el número de fila alineado con el DNI";
            return null;
        }
        var following = numbers.FirstOrDefault(t => t.Bounds.Top > current.Bounds.Top + 4.0);
        // El nombre puede ocupar dos o más líneas. No admitir el pie de página.
        double limit = Math.Min(pageHeight - 13.0, current.Bounds.Top + 23.0);
        if (following is not null) limit = Math.Min(limit, following.Bounds.Top - 0.7);
        var row = tokens.Where(t => t.Bounds.Top >= current.Bounds.Top - 0.6 &&
            t.Bounds.Bottom <= limit &&
            t.Bounds.Left >= pageWidth * 0.045 &&
            t.Bounds.Right <= pageWidth * 0.96).ToList();
        if (!row.Any(t => ReferenceEquals(t, dni))) row.Add(dni);
        // El registro debe contener monto, estado y nombre del beneficiario.
        bool nameFound = row.Any(t => t.Bounds.Left >= pageWidth * 0.08 &&
            t.Bounds.Left < pageWidth * 0.22 && t.Text.Any(char.IsLetter));
        if (!nameFound || FindBankAmount(row, "BCP", pageWidth) is null ||
            !row.Any(t => t.Text.Equals("PROCESADA", StringComparison.OrdinalIgnoreCase)))
        {
            reason = "BCP: nombre, monto o estado no identificados en la fila";
            return null;
        }
        var content = Box.Union(row.Select(t => t.Bounds));
        // La búsqueda robusta intenta primero el borde normal y después
        // márgenes y grosores menores; nunca permite contacto con texto.
        return TrySafeBorder(tokens, new List<Token> { dni }, dni, content,
            following?.Bounds.Top ?? pageHeight - 10.0,
            pageWidth, pageHeight, out borderWidth, out reason);
    }

    // V6: segmentación geométrica específica para las planillas Interbank.
    // Importante: no confundir el DNI (última línea de la columna Beneficiario)
    // con la primera línea de la fila. El siguiente nombre marca el límite real.
    private static Box? CalculateInterbankFrame(List<Token> tokens, Token dni,
        List<Token> anchors, double pageWidth, double pageHeight,
        out double borderWidth, out string reason)
    {
        borderWidth = 1.35;
        reason = "";
        var ordered = anchors.OrderBy(a => a.MidY).ToList();
        int index = ordered.FindIndex(x => ReferenceEquals(x, dni));
        if (index < 0) { reason = "Interbank: ancla DNI no localizada"; return null; }

        Token? previous = index > 0 ? ordered[index - 1] : null;
        Token? next = index + 1 < ordered.Count ? ordered[index + 1] : null;
        double beneficiaryRight = pageWidth * 0.28;
        // Los nombres pueden tener hasta cuatro renglones antes del DNI.
        // La ancla previa solo sirve para no recuperar nombres ajenos.
        double searchTop = previous is null ? Math.Max(0, dni.Bounds.Top - 65)
            : Math.Max(previous.Bounds.Bottom + 0.4, dni.Bounds.Top - 65);
        bool BeneficiaryText(Token t) =>
            t.Bounds.Left >= pageWidth * 0.035 && t.Bounds.Left < beneficiaryRight &&
            !Regex.IsMatch(t.Text, @"^(?:DNI\d*|\d{8})$", RegexOptions.IgnoreCase) &&
            t.Text.Any(char.IsLetter);

        var nameLines = tokens.Where(t => BeneficiaryText(t) &&
            t.Bounds.Top >= searchTop && t.Bounds.Top < dni.Bounds.Top - 0.2)
            .OrderBy(t => t.Bounds.Top).ToList();
        if (nameLines.Count == 0)
        { reason = "Interbank: no se identificó el nombre encima del DNI"; return null; }
        // Quedarse con las líneas contiguas inmediatamente anteriores al DNI.
        // En la primera fila de cada página descarta la cabecera Beneficiario.
        var nameLineTops = nameLines.Select(t => t.Bounds.Top)
            .OrderByDescending(y => y).ToList();
        double rowStart = nameLineTops[0];
        double lastTop = rowStart;
        foreach (double top in nameLineTops.Skip(1))
        {
            if (lastTop - top > 17.5) break;
            rowStart = top;
            lastTop = top;
        }

        // El nombre del siguiente titular determina el corte de la fila actual.
        // No se presupone que la distancia entre DNI consecutivos sea fija.
        double nextStart = pageHeight;
        if (next is not null)
        {
            var followingNames = tokens.Where(t => BeneficiaryText(t) &&
                    t.Bounds.Top > dni.Bounds.Bottom + 0.5 &&
                    t.Bounds.Top < next.Bounds.Top - 0.2)
                .OrderBy(t => t.Bounds.Top).ToList();
            if (followingNames.Count == 0)
            { reason = "Interbank: no se pudo ubicar el inicio del siguiente beneficiario"; return null; }
            nextStart = followingNames[0].Bounds.Top;
        }

        // Evita incluir el pie de página cuando se trata del último registro.
        double bandBottom = Math.Min(Math.Min(nextStart - 0.5, dni.Bounds.Bottom + 9), pageHeight - 3);
        var row = tokens.Where(t => t.Bounds.Top >= rowStart - 0.4 &&
                                    t.Bounds.Bottom <= bandBottom + 0.3).ToList();
        if (!row.Any(t => ReferenceEquals(t, dni))) row.Add(dni);
        var name = row.Where(BeneficiaryText).ToList();
        var amount = FindBankAmount(row, "INTERBANK", pageWidth);
        bool hasStatus = row.Any(t => t.Text.Equals("Procesado", StringComparison.OrdinalIgnoreCase));
        if (name.Count == 0 || amount is null || !hasStatus)
        { reason = "Interbank: nombre, importe o estado incompleto en la fila"; return null; }
        var content = Box.Union(row.Select(t => t.Bounds));
        if (content.Width < 180 || content.Height < 8 || content.Height > 75)
        { reason = "Interbank: geometría del registro fuera de límites esperados"; return null; }
        return TrySafeBorder(tokens, ordered, dni, content, nextStart,
            pageWidth, pageHeight, out borderWidth, out reason);
    }

    // Misma protección de letras para cada uno de los cuatro trazos.
    // Si hay poco espacio, elige progresivamente un borde más fino.
    private static Box? TrySafeBorder(List<Token> tokens, List<Token> anchors,
        Token dni, Box content, double nextRowStart, double pageWidth, double pageHeight,
        out double borderWidth, out string reason)
    {
        borderWidth = 1.35;
        reason = "No hay espacio seguro para un rectángulo completo";
        foreach (double stroke in new[] { 1.35, 1.0, 0.7, 0.45, 0.3 })
        foreach (double margin in new[] { 3.5, 2.0, 1.2, 0.7, 0.4, 0.28 })
        {
            double half = stroke / 2.0;
            double inset = half + 0.14;
            var b = new Box(Math.Max(inset, content.Left - margin),
                Math.Max(inset, content.Top - margin),
                Math.Min(pageWidth - inset, content.Right + margin),
                Math.Min(pageHeight - inset, content.Bottom + margin));
            if (b.Width <= 0 || b.Height <= 0 || b.Bottom + half >= nextRowStart) continue;
            if (content.Left - b.Left <= inset || b.Right - content.Right <= inset ||
                content.Top - b.Top <= inset || b.Bottom - content.Bottom <= inset) continue;
            if (anchors.Any(x => !ReferenceEquals(x, dni) && x.Bounds.Overlaps(b, half))) continue;
            var top = new Box(b.Left - inset, b.Top - inset, b.Right + inset, b.Top + inset);
            var bottom = new Box(b.Left - inset, b.Bottom - inset, b.Right + inset, b.Bottom + inset);
            var left = new Box(b.Left - inset, b.Top - inset, b.Left + inset, b.Bottom + inset);
            var right = new Box(b.Right - inset, b.Top - inset, b.Right + inset, b.Bottom + inset);
            if (tokens.Any(t => t.Bounds.Overlaps(top) || t.Bounds.Overlaps(bottom) ||
                                t.Bounds.Overlaps(left) || t.Bounds.Overlaps(right))) continue;
            borderWidth = stroke;
            reason = "";
            return b;
        }
        return null;
    }

    // V14: BBVA puede imprimir las últimas palabras del registro en la página
    // siguiente, o anticipar parte del titular al pie de la página ANTERIOR.
    // No modifica/corta/pega texto: guarda todas las páginas originales y
    // dibuja un rectángulo independiente alrededor de cada fragmento comprobado.
    // Se verifica: pertenencia al nombre Excel, ubicación fuera de la última /
    // primera fila completa de esa página, y separación de los demás textos.
    private static List<PageFrame> FindBbvaPageContinuations(Match m,
        Dictionary<int, List<Token>> pages,
        Dictionary<int, (double Width, double Height)> sizes)
    {
        var found = new List<PageFrame>();
        if (!sizes.TryGetValue(m.Page, out var size)) return found;
        var workerParts = Regex.Matches(NormalizePerson(m.Worker.Name), @"[A-Z]{3,}")
            .Select(x => x.Value).Distinct().ToHashSet(StringComparer.Ordinal);
        if (workerParts.Count < 2) return found;

        // Inicialmente se comprueba el límite que realmente contiene el registro.
        // Se admiten continuaciones en 3+ páginas, pero solo mientras exista un
        // fragmento identificable en el borde opuesto y la cadena siga contigua.
        Scan(m.Page, -1, m.Frame.Top < 125);
        Scan(m.Page, +1, m.Frame.Bottom > size.Height - 135);
        return found;

        void Scan(int start, int direction, bool adjacentToEdge)
        {
            int page = start + direction;
            int hops = 0;
            while (adjacentToEdge && hops++ < pages.Count && pages.ContainsKey(page))
            {
                if (!TryFragment(page, direction < 0, out var frame)) break;
                found.Add(frame!);
                var dimensions = sizes[page];
                adjacentToEdge = direction < 0
                    ? frame!.Frame.Top < 110
                    : frame!.Frame.Bottom > dimensions.Height - 110;
                page += direction;
            }
        }

        bool TryFragment(int page, bool atBottom, out PageFrame? result)
        {
            result = null;
            if (!pages.TryGetValue(page, out var words) ||
                !sizes.TryGetValue(page, out var dimensions)) return false;
            double width = dimensions.Width, height = dimensions.Height;

            // El DNI es una referencia de inicio de registro, pero para una
            // continuación NO se exige otro DNI en la página adyacente.
            var dniAnchors = FindDniTokens(words, "BBVA").OrderBy(t => t.MidY).ToList();
            if (dniAnchors.Count == 0) return false;
            double firstDni = dniAnchors[0].MidY;
            double lastDni = dniAnchors[^1].MidY;
            // Solo la franja de texto que queda separada de las filas con DNI.
            // Arriba: texto antes de la primera fila identificada.
            // Abajo: texto después del último DNI del registro final.
            var names = words.Where(t =>
                t.Bounds.Left >= width * 0.37 && t.Bounds.Left < width * 0.65 &&
                t.Bounds.Right < width * 0.67 && t.Text.Any(char.IsLetter) &&
                (atBottom
                    ? t.Bounds.Top > lastDni + 9 && t.Bounds.Top > height - 145
                    : t.Bounds.Bottom < firstDni - 17 && t.Bounds.Top < 110))
                .ToList();
            if (names.Count == 0) return false;

            // En BBVA se usan dos columnas de titulares. Un único apellido
            // puede repetirse en ambas (p.ej. SUCASAIRE/SUCASAIRE), y la
            // anterior regla de 'dos palabras diferentes' lo omitía.
            bool Similar(string token, string part)
            {
                if (token == part) return true;
                // La columna Titular(Banco) puede truncar apellidos o nombres.
                // Solo permitir prefijos largos, no coincidencias de 2-3 letras.
                return token.Length >= 5 && part.Length >= 5 &&
                    (part.StartsWith(token, StringComparison.Ordinal) ||
                     token.StartsWith(part, StringComparison.Ordinal));
            }
            bool Belongs(Token t) => Regex.Matches(NormalizePerson(t.Text), @"[A-Z]{3,}")
                .Select(x => x.Value).Any(p => workerParts.Any(w => Similar(p, w)));
            var related = names.Where(Belongs).OrderBy(t => t.Bounds.Top).ToList();
            if (related.Count == 0) return false;

            // No basta un solo token común: exigir dos columnas de titulares,
            // o al menos dos términos distintos del nombre confirmado.
            bool leftColumn = related.Any(t => t.Bounds.Left < width * 0.515);
            bool rightColumn = related.Any(t => t.Bounds.Left >= width * 0.515);
            int differentWords = related.SelectMany(t => Regex.Matches(NormalizePerson(t.Text), @"[A-Z]{3,}")
                .Select(x => x.Value)).Distinct().Count();
            if (!(leftColumn && rightColumn) && differentWords < 2) return false;

            // Exigir que no haya nombres ajenos intercalados en el fragmento.
            double top = related.Min(t => t.Bounds.Top);
            double bottom = related.Max(t => t.Bounds.Bottom);
            if (bottom - top > 65) return false;
            var contentWords = names.Where(t => t.Bounds.Top >= top - 3 &&
                t.Bounds.Bottom <= bottom + 3).ToList();
            if (contentWords.Any(t => !Belongs(t))) return false;
            if (contentWords.Count < 2) return false;
            // V16: para fragmentos divididos, usar la separación horizontal
            // entre filas como frontera geométrica (inferida del hueco real entre
            // las palabras de dos registros). No limitar el marco al nombre.
            // Incluir cuenta, estado, número de registro, titulares y demás
            // columnas impresas EN ESA MISMA FRACCIÓN.
            // V17: al final de página el número de cuenta y el estado ABONO
            // pueden quedar 2-3 puntos más abajo que la última palabra del
            // titular. Incluir la banda COMPLETA del registro; de otro modo
            // se excluyen cuenta/estado del contenido y la validación del
            // borde inferior los interpreta como texto ajeno, descartándolo.
            // Solo ampliar hacia abajo en la página ANTERIOR (atBottom): en
            // la página siguiente hay otra fila inmediatamente después.
            double segmentTop = top - 1.5;
            double segmentBottom = bottom + (atBottom ? 10.0 : 2.0);
            var sameFragment = words.Where(t =>
                t.Bounds.Left >= width * 0.105 &&
                t.Bounds.Right <= width * 0.955 &&
                t.Bounds.Top >= segmentTop &&
                t.Bounds.Bottom <= segmentBottom).ToList();
            if (!contentWords.All(t => sameFragment.Contains(t))) return false;
            if (dniAnchors.Any(t => sameFragment.Contains(t))) return false;
            if (sameFragment.Count == 0) return false;
            var content = Box.Union(sameFragment.Select(t => t.Bounds));
            if (content.Width > width * 0.90 || content.Height > 65) return false;

            // La primera/última línea divisoria de la tabla delimita la fracción
            // incluso cuando el resto del registro queda en la otra página.
            // Cuando no se pueden extraer trazos vectoriales, el punto medio del
            // hueco entre filas es una referencia equivalente y conservadora.
            var otherContent = words.Where(t =>
                t.Bounds.Left >= width * 0.105 &&
                t.Bounds.Right <= width * 0.955 &&
                !sameFragment.Contains(t)).ToList();
            double before = otherContent.Where(t => t.Bounds.Bottom <= content.Top - 0.1)
                .Select(t => t.Bounds.Bottom).DefaultIfEmpty(double.NaN).Max();
            double after = otherContent.Where(t => t.Bounds.Top >= content.Bottom + 0.1)
                .Select(t => t.Bounds.Top).DefaultIfEmpty(double.NaN).Min();
            double topDivider = double.IsNaN(before)
                ? content.Top - 2.0 : (before + content.Top) / 2;
            double bottomDivider = double.IsNaN(after)
                ? content.Bottom + 2.0 : (content.Bottom + after) / 2;
            if (topDivider >= content.Top - 0.35 || bottomDivider <= content.Bottom + 0.35)
                return false;
            // BBVA tiene columnas desde aprox. 12% a 93% de la anchura.
            // En una fila partida, el borde debe abarcar TODAS las columnas,
            // no quedar como un pequeño cuadro alrededor de dos apellidos.
            double left = Math.Max(1.0, Math.Min(content.Left - 2, width * 0.12));
            double right = Math.Min(width - 1.0, Math.Max(content.Right + 2, width * 0.93));
            foreach (double stroke in new[] { 1.15, 0.8, 0.5, 0.3 })
            {
                double safety = stroke / 2 + 0.12;
                double frameTop = Math.Max(safety + 0.15, topDivider);
                double frameBottom = Math.Min(height - safety - 0.15, bottomDivider);
                var candidate = new Box(left, frameTop, right, frameBottom);
                if (candidate.Height < 10 || candidate.Top + safety >= content.Top ||
                    candidate.Bottom - safety <= content.Bottom) continue;
                var edgeTop = new Box(left - safety, frameTop - safety, right + safety, frameTop + safety);
                var edgeBottom = new Box(left - safety, frameBottom - safety, right + safety, frameBottom + safety);
                var edgeLeft = new Box(left - safety, frameTop - safety, left + safety, frameBottom + safety);
                var edgeRight = new Box(right - safety, frameTop - safety, right + safety, frameBottom + safety);
                if (words.Any(t => t.Bounds.Overlaps(edgeTop) || t.Bounds.Overlaps(edgeBottom) ||
                                   t.Bounds.Overlaps(edgeLeft) || t.Bounds.Overlaps(edgeRight))) continue;
                if (dniAnchors.Any(t => t.Bounds.Overlaps(candidate, stroke / 2))) continue;
                result = new PageFrame(page, candidate, stroke);
                return true;
            }
            return false;
        }
    }

    private static string NormalizePerson(string s)
    {
        var decomposed = s.ToUpperInvariant().Normalize(NormalizationForm.FormD);
        return new string(decomposed.Where(c =>
            CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
    }

    private static Token? FindBankAmount(List<Token> row, string bank, double width)
    {
        (double lo, double hi) = bank switch
        {
            "BBVA" => (0.75, 0.86),
            "BCP" => (0.31, 0.40),
            "INTERBANK" => (0.62, 0.72),
            "SCOTIABANK" => (0.45, 0.51),
            _ => (0, 1)
        };
        return row.Where(t => t.Bounds.Left >= lo * width && t.Bounds.Left <= hi * width &&
            ParseAmount(t.Text) is not null).OrderBy(t => t.Bounds.Left).FirstOrDefault();
    }

    private static decimal? ParseAmount(string input)
    {
        string s = input.Replace("S/", "").Replace(" ", "").Trim();
        if (!Regex.IsMatch(s, @"^\d{1,3}(?:,\d{3})*\.\d{2}$|^\d+\.\d{2}$")) return null;
        return decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    private static string FindBankStatus(List<Token> row, string bank)
    {
        var words = row.OrderBy(t => t.MidY).ThenBy(t => t.Bounds.Left).Select(t => t.Text.ToUpperInvariant()).ToList();
        if (bank == "BBVA" && words.Contains("ABONO") && words.Contains("CORRECTO")) return "ABONO CORRECTO";
        if (bank == "BCP" && words.Contains("PROCESADA")) return "PROCESADA";
        if (bank == "INTERBANK" && words.Contains("PROCESADO")) return "PROCESADO";
        if (bank == "SCOTIABANK" && words.Any(w => w == "O.K." || w == "OK" || w == "O.K")) return "O.K.";
        return "NO VERIFICADO";
    }

    private static bool IsWithin(Box inner, Box outer) =>
        inner.Left >= outer.Left && inner.Right <= outer.Right && inner.Top >= outer.Top && inner.Bottom <= outer.Bottom;

    private static bool IsAmount(string token, decimal amount)
    {
        var plain = token.Replace("S/", "").Replace(" ", "");
        if (decimal.TryParse(plain, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)) return parsed == amount;
        if (decimal.TryParse(plain.Replace(".", "").Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out parsed)) return parsed == amount;
        return false;
    }

    // V13: localizar la fecha POR SU ETIQUETA y marcar únicamente el valor.
    // Evita los falsos rechazos de V12 cuando se pretendía encerrar la etiqueta
    // y algunas palabras intermedias quedaban fuera del grupo.
    // Todas las posiciones provienen de las palabras reales del PDF.
    private static Box? FindBankPaymentDateFrame(List<Token> tokens, string bank, double width, double height)
    {
        var dates = tokens.Where(t => Regex.IsMatch(t.Text.Trim(),
            @"^\d{1,2}[/.-]\d{1,2}[/.-]\d{4}$")).ToList();
        if (dates.Count == 0) return null;

        Token? selected = null;
        if (bank == "BCP" || bank == "INTERBANK")
        {
            // BCP: "Fecha de proceso:"; Interbank: "Programación de envío".
            // El texto puede venir separado en varias palabras o agrupado.
            var labels = tokens.Where(t => bank == "BCP"
                ? Regex.IsMatch(t.Text, @"proceso\b", RegexOptions.IgnoreCase)
                    && t.MidY < height * 0.55
                : (t.Text.Contains("Programaci", StringComparison.OrdinalIgnoreCase)
                   && t.MidY < height * 0.55)).ToList();
            var possible = (from label in labels
                            from date in dates
                            let dy = Math.Abs(label.MidY - date.MidY)
                            let dx = date.Bounds.Left - label.Bounds.Right
                            where dy <= 9.0 && dx >= -2.0 && dx <= width * 0.46
                            orderby dy * 15.0 + Math.Abs(dx - width * 0.12)
                            select date).ToList();
            if (possible.Count > 0) selected = possible[0];
            // No usar "fecha de envío" de BCP ni "Creación" de Interbank
            // cuando falta la etiqueta correcta; es preferible informar.
        }
        else
        {
            (double x, double y) = bank switch
            {
                "BBVA" => (0.665, 0.325),
                "SCOTIABANK" => (0.756, 0.142),
                _ => (-1, -1)
            };
            if (x < 0) return null;
            var best = dates.Select(t => new
            {
                Token = t,
                Distance = Math.Abs(t.Bounds.Left / width - x) +
                           2.0 * Math.Abs(t.MidY / height - y)
            }).OrderBy(a => a.Distance).ToList();
            if (best.Count == 0 || best[0].Distance > 0.15) return null;
            if (best.Count > 1 && Math.Abs(best[1].Distance - best[0].Distance) < 0.007)
                return null;
            selected = best[0].Token;
        }
        if (selected is null) return null;

        // Solo la fecha (más la hora contigua en BBVA). No unirla con una
        // etiqueta distante, pues la caja puede cruzar otras palabras.
        var group = new List<Token> { selected };
        if (bank == "BBVA")
        {
            group.AddRange(tokens.Where(t => t.Bounds.Left > selected.Bounds.Right &&
                t.Bounds.Left < width * 0.83 &&
                Math.Abs(t.MidY - selected.MidY) <= 4.5 &&
                Regex.IsMatch(t.Text.Trim(), @"^\d{1,2}:\d{2}(:\d{2})?$")));
        }
        var content = Box.Union(group.Select(t => t.Bounds));
        var otherWords = tokens.Where(t => !group.Contains(t)).ToList();
        // Adaptar margen: la seguridad visual es más importante que el grosor.
        foreach (double margin in new[] { 3.0, 2.0, 1.2, 0.6, 0.25 })
        {
            var frame = new Box(Math.Max(0.5, content.Left - margin),
                Math.Max(0.5, content.Top - margin),
                Math.Min(width - 0.5, content.Right + margin),
                Math.Min(height - 0.5, content.Bottom + margin));
            if (otherWords.Any(t => frame.Overlaps(t.Bounds, 0.35))) continue;
            return frame;
        }
        return null;
    }

    private static void WriteGeneral(string source, string target, List<Match> matches)
    {
        using var doc = PdfReader.Open(source, PdfDocumentOpenMode.Modify);
        foreach (var pageGroup in matches.GroupBy(m => m.Page))
        {
            var page = doc.Pages[pageGroup.Key - 1];
            using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            foreach (var m in pageGroup)
            {
                Underline(graphics, m.DniBox);
                if (m.AmountBox is not null) Underline(graphics, m.AmountBox);
            }
        }
        doc.Save(target);
    }

    private static void Underline(XGraphics gfx, Box box)
    {
        var y = box.Bottom + 1.3;
        // V7: los PDF de Interbank pueden terminar el contenido original con
        // el color de trazo BLANCO. PDFsharp puede omitir un "0 0 0 RG"
        // si cree que negro es el color inicial. Forzar primero la transición
        // de color (sin trazo visible), y luego pedir negro explícitamente.
        ForceBlackStroke(gfx, box.Left, y);
        gfx.DrawLine(new XPen(XColors.Black, 1.15), box.Left, y, box.Right, y);
    }

    private static void ForceBlackStroke(XGraphics gfx, double x, double y)
    {
        // Un segmento prácticamente invisible con color blanco fuerza el cambio
        // interno de estado del lápiz. El siguiente lápiz negro debe emitir
        // explícitamente el comando "0 0 0 RG", incluso si el contenido
        // importado terminó con "1 1 1 RG" (trazo blanco).
        // Esta marca no dibuja un segmento visible ni altera el contenido.
        gfx.DrawLine(new XPen(XColors.White, 0.1), x, y, x + 0.05, y);
    }

    private static void WriteIndividual(string source, string target, Match m, Box? paymentDateFrame)
    {
        using var input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
        using var doc = new PdfSharp.Pdf.PdfDocument();
        // Scotiabank: la constancia individual incluye solamente la hoja del DNI.
        // Los demás bancos conservan la portada y las hojas de continuación.
        bool isScotiabank = string.Equals(m.Bank, "SCOTIABANK", StringComparison.OrdinalIgnoreCase);
        var requiredPages = isScotiabank
            ? new SortedSet<int> { m.Page }
            : new SortedSet<int> { 1, m.Page };
        if (!isScotiabank)
            foreach (var extra in m.Continuations) requiredPages.Add(extra.Page);
        var position = new Dictionary<int, int>();
        foreach (int pageNumber in requiredPages)
        {
            position[pageNumber] = doc.PageCount;
            doc.AddPage(input.Pages[pageNumber - 1]);
        }
        // V12: rodear la fecha de ejecución/programación del banco en la portada.
        // Se dibuja sólo si la fecha fue encontrada con suficiente confianza.
        if (!isScotiabank && paymentDateFrame is not null)
        {
            var cover = doc.Pages[position[1]];
            using var dateGraphics = XGraphics.FromPdfPage(cover, XGraphicsPdfPageOptions.Append);
            ForceBlackStroke(dateGraphics, paymentDateFrame.Left, paymentDateFrame.Top);
            dateGraphics.DrawRectangle(new XPen(XColors.Black, 1.25),
                paymentDateFrame.Left, paymentDateFrame.Top,
                paymentDateFrame.Width, paymentDateFrame.Height);
        }
        var frames = new List<PageFrame> { new(m.Page, m.Frame, m.BorderWidth) };
        frames.AddRange(m.Continuations);
        foreach (var mark in frames)
        {
            var page = doc.Pages[position[mark.Page]];
            using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            ForceBlackStroke(gfx, mark.Frame.Left, mark.Frame.Top);
            gfx.DrawRectangle(new XPen(XColors.Black, mark.BorderWidth),
                mark.Frame.Left, mark.Frame.Top, mark.Frame.Width, mark.Frame.Height);
        }
        doc.Save(target);
    }

    private static string PaymentFromFile(string file)
    {
        var normalized = file.Replace('\\', '/');
        foreach (var payment in new[] { "DESCANSO MEDICO", "LICENCIA", "SUBSIDIOS", "VACACIONES", "PAGO GRATIFICACIÓN", "PAGO CTS", "PAGO 3", "PAGO 2", "PAGO 1" })
            if (normalized.Contains("/" + payment + "/", StringComparison.OrdinalIgnoreCase)) return payment;
        return "SIN IDENTIFICAR";
    }

    private static void WriteSummary(string path, List<Worker> workers, List<Match> matches,
        List<(string dni, string pdf, int page, string reason)> reviews, List<PdfAudit> audits, List<string> payments)
    {
        using var book = new XLWorkbook();
        var ws = book.AddWorksheet("RESUMEN");
        string[] head = ["DNI", "NOMBRE COMPLETO", "MONTO EXCEL", "MONTO BANCO", "ESTADO BANCO", "BANCO", "RESULTADO", "PDF ORIGEN", "PÁGINA", "OBSERVACIÓN", "PAGO"];
        for (int c = 0; c < head.Length; c++) ws.Cell(1, c + 1).Value = head[c];
        int r = 2;
        foreach (var payment in payments)
        foreach (var w in workers)
        {
            var hits = matches.Where(x => x.Payment == payment && x.Worker.ExcelRow == w.ExcelRow && x.Worker.Dni == w.Dni).ToList();
            var issues = reviews.Where(x => x.dni == w.Dni && PaymentFromFile(x.pdf) == payment).ToList();
            if (hits.Count == 0) hits.Add(null!);
            foreach (var hit in hits)
            {
                ws.Cell(r, 1).Value = w.Dni;
                ws.Cell(r, 2).Value = w.Name;
                if (w.Amount.HasValue) ws.Cell(r, 3).Value = w.Amount.Value;
                if (hit?.BankAmount is not null) ws.Cell(r, 4).Value = hit.BankAmount.Value;
                ws.Cell(r, 5).Value = hit?.BankStatus ?? "";
                ws.Cell(r, 6).Value = hit?.Bank ?? "";
                bool mismatch = hit is not null && w.Amount.HasValue && hit.BankAmount.HasValue && w.Amount.Value != hit.BankAmount.Value;
                ws.Cell(r, 7).Value = hit is null ? (issues.Count > 0 ? "REVISIÓN MANUAL" : "NO ENCONTRADO") :
                    mismatch ? "MONTO DISTINTO (REVISAR)" : hit.BankAmount is null || hit.BankStatus == "NO VERIFICADO" ? "PDF GENERADO / REVISAR" : "PDF GENERADO";
                ws.Cell(r, 8).Value = hit is not null ? Path.GetFileName(hit.Pdf) : "";
                if (hit is not null) ws.Cell(r, 9).Value = hit.Page;
                ws.Cell(r, 10).Value = hit is null ? string.Join("; ", issues.Select(i => i.reason).Distinct()) :
                    mismatch ? "Monto Excel y monto bancario no coinciden" : "Confirmar nombre y conformidad del pago.";
                ws.Cell(r, 11).Value = payment;
                r++;
            }
        }
        ws.Range(1, 1, 1, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#17365D");
        ws.Range(1, 1, 1, 11).Style.Font.FontColor = XLColor.White;
        ws.Range(1, 1, 1, 11).Style.Font.Bold = true;
        ws.Column(1).Width = 15; ws.Column(2).Width = 38; ws.Column(3).Width = 18;
        ws.Column(4).Width = 18; ws.Column(5).Width = 22; ws.Column(6).Width = 17;
        ws.Column(7).Width = 30; ws.Column(8).Width = 50; ws.Column(9).Width = 13; ws.Column(10).Width = 65; ws.Column(11).Width = 15;
        ws.Column(3).Style.NumberFormat.Format = "#,##0.00";
        ws.Column(4).Style.NumberFormat.Format = "#,##0.00";
        ws.SheetView.FreezeRows(1);
        ws.Range(1, 1, Math.Max(r - 1, 1), 11).SetAutoFilter();
        var log = book.AddWorksheet("REVISION_MANUAL");
        log.Cell(1, 1).Value = "DNI"; log.Cell(1, 2).Value = "PDF";
        log.Cell(1, 3).Value = "PÁGINA"; log.Cell(1, 4).Value = "MOTIVO"; log.Cell(1, 5).Value = "PAGO";
        for (int i = 0; i < reviews.Count; i++)
        {
            var issue = reviews[i];
            log.Cell(i + 2, 1).Value = issue.dni;
            log.Cell(i + 2, 2).Value = Path.GetFileName(issue.pdf);
            log.Cell(i + 2, 3).Value = issue.page;
            log.Cell(i + 2, 4).Value = issue.reason;
            log.Cell(i + 2, 5).Value = PaymentFromFile(issue.pdf);
        }
        log.Columns(1, 5).AdjustToContents(10, 75);
        var audit = book.AddWorksheet("DIAGNOSTICO_POR_BANCO");
        string[] cols = ["ARCHIVO PDF", "BANCO", "PÁGINAS", "PALABRAS EXTRAÍDAS", "DNI LEÍDOS", "DNI EN EXCEL", "INDIVIDUALES", "INCIDENCIAS", "ESTADO", "PAGO"];
        for (int j = 0; j < cols.Length; j++) audit.Cell(1, j + 1).Value = cols[j];
        for (int i = 0; i < audits.Count; i++)
        {
            var a = audits[i]; int row = i + 2;
            audit.Cell(row, 1).Value = a.File; audit.Cell(row, 2).Value = a.Bank;
            audit.Cell(row, 3).Value = a.Pages; audit.Cell(row, 4).Value = a.TextTokens;
            audit.Cell(row, 5).Value = a.DniOccurrences; audit.Cell(row, 6).Value = a.DniInExcel;
            audit.Cell(row, 7).Value = a.Generated; audit.Cell(row, 8).Value = a.Reviews;
            audit.Cell(row, 9).Value = a.Status; audit.Cell(row, 10).Value = a.Payment;
        }
        audit.Range(1, 1, 1, cols.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("#17365D");
        audit.Range(1, 1, 1, cols.Length).Style.Font.FontColor = XLColor.White;
        audit.Range(1, 1, 1, cols.Length).Style.Font.Bold = true;
        audit.Columns(1, cols.Length).AdjustToContents(10, 60);
        audit.SheetView.FreezeRows(1);
        book.SaveAs(path);
    }

    private static string DetectBankFromText(string text)
    {
        string t = text.ToUpperInvariant();
        if (t.Contains("BBVA") || t.Contains("CONSULTA DE PAGOS MASIVOS")) return "BBVA";
        if (t.Contains("TELECRÉDITO") || t.Contains("TELECREDITO") || t.Contains("DETALLE DE OPERACIÓN - PLANILLA")) return "BCP";
        if (t.Contains("INTERBANK") || t.Contains("DETALLE DE SOLICITUD")) return "INTERBANK";
        if (t.Contains("SCOTIABANK") || t.Contains("DETALLE DE ORDEN NO.")) return "SCOTIABANK";
        return "BANCO_NO_IDENTIFICADO";
    }

    private static string DetectBank(string path)
    {
        string s = Path.GetFileName(path).ToUpperInvariant();
        foreach (string bank in new[] { "BBVA", "BCP", "INTERBANK", "SCOTIABANK", "BANBIF", "PICHINCHA", "NACION", "CREDITO" })
            if (s.Contains(bank)) return bank;
        return "BANCO_NO_IDENTIFICADO";
    }
    private static string SafeFile(string s)
    {
        var name = new string(s.Select(c => Path.GetInvalidFileNameChars().Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim(' ', '.');
        return string.IsNullOrWhiteSpace(name) ? "SIN_NOMBRE" : name[..Math.Min(100, name.Length)];
    }
    private static string UniquePath(string path)
    {
        if (!File.Exists(path)) return path;
        string dir = Path.GetDirectoryName(path)!;
        string stem = Path.GetFileNameWithoutExtension(path);
        string ext = Path.GetExtension(path);
        for (int i = 2; i < 100000; i++)
        {
            string candidate = Path.Combine(dir, $"{stem}_{i}{ext}");
            if (!File.Exists(candidate)) return candidate;
        }
        throw new IOException("Demasiados PDF con el mismo nombre.");
    }
}
