using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ConciliadorBancario;

/// <summary>
/// Consolida por el orden de secciones elegido; dentro de los documentos individuales,
/// los trabajadores mantienen el orden físico del Excel.
/// CTR y VIDA LEY se insertan como PDFs completos en la posición seleccionada.
/// </summary>
internal static class ConsolidatedPdfBuilder
{
    private static readonly string[] SectionOrder = DocumentOrderSettings.DefaultOrder;
    private static readonly Regex DniPrefix = new(@"^(?<dni>\d{8})[_\s-]", RegexOptions.Compiled);
    private static readonly Regex BareDni = new(@"^\d{8}$", RegexOptions.Compiled);

    internal sealed record Outcome(int Pages, int Employees, int IncludedFiles, int Missing, string PdfPath);
    private sealed record InputItem(string Dni, string Section, string PdfPath);

    public static Outcome Generate(string root, IReadOnlyList<Worker> workers,
        bool ctrSelected, bool vidaSelected, bool boletasSelected, bool payflowSelected,
        IReadOnlyList<string> selectedPayments, IReadOnlyList<string> documentOrder,
        IProgress<PreparationProgress> progress, CancellationToken ct)
    {
        string finalPath = Path.Combine(root, "CONSOLIDADO_FINAL.pdf");
        string csvPath = Path.Combine(root, "CONSOLIDADO_DETALLE.csv");
        string notePath = Path.Combine(root, "CONSOLIDADO_LEEME.txt");
        var enabled = new HashSet<string>(selectedPayments, StringComparer.OrdinalIgnoreCase);
        if (ctrSelected) enabled.Add("CTR");
        if (vidaSelected) enabled.Add("VIDA LEY");
        var sections = DocumentOrderSettings.Normalize(documentOrder);
        if (boletasSelected) enabled.Add("BOLETA");
        if (payflowSelected) enabled.Add("PAYFLOW");
        var catalog = BuildCatalog(root, ct);
        var ordered = new List<string>();
        var seenWorkers = new HashSet<string>(StringComparer.Ordinal);
        // El Excel es la autoridad para el ORDEN de trabajadores: se respeta el orden de lectura,
        // no el orden de las carpetas ni el orden alfabético de DNI.
        foreach (var worker in workers)
            if (BareDni.IsMatch(worker.Dni) && seenWorkers.Add(worker.Dni)) ordered.Add(worker.Dni);
        // BOLETAS puede procesarse sin Excel. En ese caso, respetar orden de página
        // reportado por su propio procesador y luego agregar cualquier otro DNI detectado.
        if (ordered.Count == 0)
        {
            foreach (var dni in ReadBoletaOrder(root))
                if (seenWorkers.Add(dni)) ordered.Add(dni);
            foreach (var dni in catalog.Keys.Select(x => x.Split('\t')[0]).Distinct()
                         .OrderBy(x => x, StringComparer.Ordinal))
                if (seenWorkers.Add(dni)) ordered.Add(dni);
        }

        var log = new List<string> { "ORDEN;DNI;SECCION;ESTADO;ARCHIVO;PAGINAS;OBSERVACION" };
        int pages = 0, included = 0, missing = 0;
        using (var output = new PdfDocument())
        {
            output.Info.Title = "Consolidado documental en orden configurado";
            output.Info.Subject = "Trabajadores en orden Excel; secciones en orden del usuario";
            // BLOQUE 1: todos los pagos, trabajador por trabajador según Excel.
            // BLOQUE 2: adicionales completos; solo después de finalizar los pagos.
            var paymentSections = sections.Where(s => DocumentOrderSettings.PaymentOrder.Contains(s)).ToArray();
            var additionalSections = sections.Where(s => DocumentOrderSettings.AdditionalOrder.Contains(s)).ToArray();
            foreach (string dni in ordered)
            {
                ct.ThrowIfCancellationRequested();
                foreach (string section in paymentSections)
                {
                    if (!enabled.Contains(section)) continue;
                    string key = dni + "\t" + section;
                    if (!catalog.TryGetValue(key, out var paths) || paths.Count == 0)
                    {
                        missing++;
                        log.Add(Line("PAGOS", dni, section, "NO ENCONTRADO", "", 0,
                            "Documento no encontrado entre los PDFs individuales identificados"));
                        continue;
                    }
                    foreach (var file in paths)
                    {
                        ct.ThrowIfCancellationRequested();
                        try
                        {
                            int count = AppendPdf(output, file, ct);
                            pages += count;
                            included++;
                            log.Add(Line("PAGOS", dni, section, "AGREGADO", Path.GetRelativePath(root, file), count, ""));
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException)
                        {
                            missing++;
                            log.Add(Line("PAGOS", dni, section, "ERROR", Path.GetRelativePath(root, file), 0, ex.Message));
                        }
                    }
                }
            }
            // Se insertan una sola vez, cada uno como PDF completo.
            foreach (string section in additionalSections)
            {
                ct.ThrowIfCancellationRequested();
                if (!enabled.Contains(section)) continue;
                AddFull(section, section == "CTR" ? "03_CTR_PDF_COMPLETO_ENMARCADO" : "04_VIDA_LEY_PDF_COMPLETO_ENMARCADO");
            }
            if (pages > 0)
            {
                ct.ThrowIfCancellationRequested();
                output.Save(finalPath);
            }
            else
                log.Add(Line("GENERAL", "", "", "SIN DOCUMENTOS", "", 0,
                    "No se encontró ningún PDF para consolidar"));

            void AddFull(string section, string folder)
            {
                ct.ThrowIfCancellationRequested();
                string dir = Path.Combine(root, folder);
                var files = Directory.Exists(dir)
                    ? Directory.EnumerateFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly)
                        .OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList()
                    : new List<string>();
                if (files.Count == 0)
                {
                    missing++;
                    log.Add(Line("CABECERA", "", section, "NO ENCONTRADO", "", 0, "No se generó el PDF completo"));
                }
                // CTR y VIDA LEY: exactamente un archivo PDF por sección, nunca duplicados.
                if (files.Count > 1)
                {
                    missing += files.Count - 1;
                    foreach (string extra in files.Skip(1))
                        log.Add(Line("CABECERA", "", section, "OMITIDO DUPLICADO", Path.GetRelativePath(root, extra), 0,
                            "Solo se permite un PDF completo por sección CTR / VIDA LEY"));
                }
                foreach (string file in files.Take(1))
                {
                    try
                    {
                        int count = AppendPdf(output, file, ct);
                        pages += count;
                        included++;
                        log.Add(Line("CABECERA", "", section, "AGREGADO", Path.GetRelativePath(root, file), count, ""));
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        missing++;
                        log.Add(Line("CABECERA", "", section, "ERROR", Path.GetRelativePath(root, file), 0, ex.Message));
                    }
                }
            }
        }
        File.WriteAllLines(csvPath, log, new UTF8Encoding(true));
        File.WriteAllLines(notePath, new[]
        {
            "PDF FINAL - ORDEN DE COMPAGINACIÓN CONFIGURADO POR EL USUARIO",
            "BLOQUE 1 - PAGOS: " + string.Join(" > ", sections.Where(s => DocumentOrderSettings.PaymentOrder.Contains(s))),
            "BLOQUE 2 - ADICIONALES: " + string.Join(" > ", sections.Where(s => DocumentOrderSettings.AdditionalOrder.Contains(s))),
            "Solo se añaden secciones seleccionadas en esta ejecución.",
            "Primero se añaden TODOS los pagos por trabajador según el orden Excel.",
            "Después de terminar todos los pagos se agregan CTR y VIDA LEY, completos, una sola vez cada uno.",
            "Los documentos seleccionados pero ausentes se informan en CONSOLIDADO_DETALLE.csv.",
            "Las boletas sin DNI identificado quedan en 05_BOLETAS/PENDIENTES_DE_IDENTIFICAR.",
            "Si no hay Excel, las boletas mantienen el orden de páginas reconocido en BOLETAS_REPORTE.csv.",
            $"Páginas: {pages}. Archivos incluidos: {included}. Incidencias: {missing}.",
            "Este documento reúne evidencias, no certifica pagos ni coberturas."
        }, new UTF8Encoding(true));
        File.WriteAllLines(Path.Combine(root, "CONSOLIDADO_ORDEN_DOCUMENTOS.txt"),
            new[] { "BLOQUE 1 - PAGOS" }.Concat(sections.Where(s => DocumentOrderSettings.PaymentOrder.Contains(s)).Select((section, i) => $"{i + 1:D2}. {section}"))
                .Concat(new[] { "", "BLOQUE 2 - DOCUMENTOS ADICIONALES" })
                .Concat(sections.Where(s => DocumentOrderSettings.AdditionalOrder.Contains(s)).Select((section, i) => $"{i + 1:D2}. {section}")), new UTF8Encoding(true));
        progress.Report(new(99, $"CONSOLIDADO: {pages} páginas en un PDF; {ordered.Count} DNI en orden Excel; {missing} incidencias."));
        return new Outcome(pages, ordered.Count, included, missing, finalPath);
    }

    private static Dictionary<string, List<string>> BuildCatalog(string root, CancellationToken ct)
    {
        var catalog = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        // GENERAL es la copia única de cada PDF bancario individual;
        // no tomar las carpetas por banco, que duplican los mismos archivos.
        string individual = Path.Combine(root, "02_PDF_INDIVIDUALES");
        foreach (var section in SectionOrder.Where(s => s is not "BOLETA" and not "PAYFLOW" and not "CTR" and not "VIDA LEY"))
        {
            string dir = Path.Combine(individual, section, "GENERAL");
            if (!Directory.Exists(dir)) continue;
            foreach (var file in Directory.EnumerateFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly))
                Add(section, file);
        }
        string boletas = Path.Combine(root, "05_BOLETAS");
        if (Directory.Exists(boletas))
            foreach (var file in Directory.EnumerateFiles(boletas, "*.pdf", SearchOption.TopDirectoryOnly))
                Add("BOLETA", file);
        string payflow = Path.Combine(root, "06_PAYFLOW", "PDF_INDIVIDUALES");
        if (Directory.Exists(payflow))
            foreach (var file in Directory.EnumerateFiles(payflow, "*.pdf", SearchOption.TopDirectoryOnly))
                Add("PAYFLOW", file);
        foreach (var list in catalog.Values)
            list.Sort(StringComparer.OrdinalIgnoreCase);
        return catalog;

        void Add(string section, string file)
        {
            ct.ThrowIfCancellationRequested();
            var match = DniPrefix.Match(Path.GetFileNameWithoutExtension(file));
            if (!match.Success) return;
            string key = match.Groups["dni"].Value + "\t" + section;
            if (!catalog.TryGetValue(key, out var paths)) catalog[key] = paths = new List<string>();
            paths.Add(file);
        }
    }

    private static IEnumerable<string> ReadBoletaOrder(string root)
    {
        string report = Path.Combine(root, "05_BOLETAS", "BOLETAS_REPORTE.csv");
        if (!File.Exists(report)) yield break;
        // La cabecera es ARCHIVO;PAGINA;DNI;TRABAJADOR;ESTADO;...
        foreach (string line in File.ReadLines(report).Skip(1))
        {
            var parts = line.Split(';');
            if (parts.Length > 2 && BareDni.IsMatch(parts[2].Trim().Trim('"')))
                yield return parts[2].Trim().Trim('"');
        }
    }

    private static int AppendPdf(PdfDocument dest, string file, CancellationToken ct)
    {
        using var source = PdfReader.Open(file, PdfDocumentOpenMode.Import);
        int count = source.PageCount;
        for (int i = 0; i < count; i++)
        {
            ct.ThrowIfCancellationRequested();
            dest.AddPage(source.Pages[i]);
        }
        return count;
    }

    private static string Line(string order, string dni, string section, string state, string file, int pages, string note) =>
        string.Join(";", new[] { order, dni, section, state, file, pages.ToString(CultureInfo.InvariantCulture), note }
            .Select(x => "\"" + x.Replace("\"", "\"\"").Replace("\r", " ").Replace("\n", " ") + "\""));
}
