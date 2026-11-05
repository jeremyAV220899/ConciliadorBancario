using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace ConciliadorBancario;

// CTR/SCTR: no recortar, separar ni crear PDFs individuales. Solo anotar el PDF completo.
internal static class CtrProcessor
{
    internal sealed record Outcome(int Pages, int Marked, int Unmatched);

    public static Outcome Generate(string inputPath, IReadOnlyList<Worker> workers, string root,
        IProgress<PreparationProgress> progress, CancellationToken ct, string documentType = "CTR")
    {
        if (documentType != "CTR" && documentType != "VIDA_LEY")
            throw new ArgumentOutOfRangeException(nameof(documentType));
        string displayName = documentType == "CTR" ? "CTR" : "VIDA LEY";
        if (!File.Exists(inputPath)) throw new FileNotFoundException($"No se encontró el PDF de {displayName}.", inputPath);
        if (!Path.GetExtension(inputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException($"{displayName} debe ser un archivo PDF.");

        var wanted = workers.Select(w => w.Dni).Where(d => Regex.IsMatch(d, @"^\d{8}$"))
            .ToHashSet(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var marks = new Dictionary<int, List<(double Top, double Bottom, string Dni)>>();
        int pageCount;
        using (var source = UglyToad.PdfPig.PdfDocument.Open(inputPath))
        {
            pageCount = source.NumberOfPages;
            for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++)
            {
                ct.ThrowIfCancellationRequested();
                var page = source.GetPage(pageNumber);
                var words = NearestNeighbourWordExtractor.Instance.GetWords(page.Letters).ToList();
                foreach (var word in words)
                {
                    string dni = word.Text.Trim();
                    if (!Regex.IsMatch(dni, @"^\d{8}$") || !wanted.Contains(dni)) continue;
                    // En constancias Mapfre la columna documento se encuentra
                    // a la izquierda del nombre; evita marcar DNI en pies o cabeceras.
                    var b = word.BoundingBox;
                    if (b.Left < page.Width * .24 || b.Left > page.Width * .48) continue;
                    double top = page.Height - b.Top;
                    double bottom = page.Height - b.Bottom;
                    if (top < 65 || bottom > page.Height - 28) continue;
                    if (!marks.TryGetValue(pageNumber, out var list)) marks[pageNumber] = list = new();
                    if (list.Any(x => x.Dni == dni && Math.Abs(x.Top - top) < 3)) continue;
                    list.Add((top, bottom, dni));
                    seen.Add(dni);
                }
                if (pageNumber % 5 == 0 || pageNumber == pageCount)
                    progress.Report(new(96, $"{displayName}: leyendo página {pageNumber}/{pageCount}."));
            }
        }
        var folder = Path.Combine(root, documentType == "CTR" ? "03_CTR_PDF_COMPLETO_ENMARCADO" : "04_VIDA_LEY_PDF_COMPLETO_ENMARCADO");
        Directory.CreateDirectory(folder);
        string outPath = Path.Combine(folder, Path.GetFileNameWithoutExtension(inputPath) + "_" + documentType + "_ENMARCADO.pdf");
        using (var doc = PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify))
        {
            foreach (var group in marks)
            {
                ct.ThrowIfCancellationRequested();
                var page = doc.Pages[group.Key - 1];
                using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                foreach (var mark in group.Value)
                {
                    // Recuadro completo por registro: número, tipo, documento y nombre.
                    // Coordenadas PdfPig se convierten a origen superior de PDFsharp.
                    double top = Math.Max(0, mark.Top - 1.2);
                    double bottom = Math.Min(page.Height.Point, mark.Bottom + 1.2);
                    // VIDA LEY incluye la columna Ocupación, que también debe quedar enmarcada.
                    double left = 78;
                    double right = documentType == "CTR" ? 468 : 529;
                    gfx.DrawRectangle(new XPen(XColors.Black, 1.1), left, top,
                        Math.Min(page.Width.Point - left - 12, right - left), Math.Max(1, bottom - top));
                }
            }
            doc.Save(outPath);
        }
        var unmatched = wanted.Except(seen).OrderBy(s => s).ToList();
        var report = new List<string> {
            displayName + " — Resultado del enmarcado (sin dividir el PDF)",
            $"PDF: {Path.GetFileName(inputPath)}", $"Páginas preservadas: {pageCount}",
            $"Documentos del Excel: {wanted.Count}", $"Documentos encontrados y enmarcados: {seen.Count}",
            $"Documentos no encontrados: {unmatched.Count}",
            "", "Documentos sin coincidencia (verificar, incluye personas ajenas a la constancia):"
        };
        report.AddRange(unmatched);
        File.WriteAllLines(Path.Combine(folder, documentType + "_REPORTE.txt"), report, Encoding.UTF8);
        progress.Report(new(98, $"{displayName}: {seen.Count} documentos marcados en PDF completo de {pageCount} páginas; {unmatched.Count} sin coincidencia."));
        return new Outcome(pageCount, seen.Count, unmatched.Count);
    }
}
