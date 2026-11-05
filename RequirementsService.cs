using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ConciliadorBancario;

internal sealed record Requirement(string Name, string Detail, bool Ready, bool Mandatory);
internal sealed record RequirementsSnapshot(IReadOnlyList<Requirement> Items)
{
    public bool AllReady => Items.All(x => x.Ready);
}

internal static class RequirementsService
{
    public static string? Find(string executable)
    {
        string root = AppContext.BaseDirectory;
        var paths = new List<string> {
            Path.Combine(root, executable),
            Path.Combine(root, "OCR", executable),
            Path.Combine(root, "OCR", "Tesseract-OCR", executable),
            Path.Combine(root, "OCR", "Poppler", "Library", "bin", executable),
            Path.Combine(root, "OCR", "Poppler", "bin", executable),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Tesseract-OCR", executable),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Tesseract-OCR", executable),
            Path.Combine(@"C:\poppler\Library\bin", executable),
            Path.Combine(@"C:\poppler\bin", executable)
        };
        foreach (string d in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(d)) continue;
            try { paths.Add(Path.Combine(d.Trim().Trim('"'), executable)); } catch { }
        }
        foreach (string file in paths)
            if (File.Exists(file)) return file;
        string packages = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WinGet", "Packages");
        if (Directory.Exists(packages))
        {
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(packages))
                {
                    if (!dir.Contains("Poppler", StringComparison.OrdinalIgnoreCase) && !dir.Contains("Tesseract", StringComparison.OrdinalIgnoreCase)) continue;
                    string? hit = Directory.EnumerateFiles(dir, executable, SearchOption.AllDirectories).FirstOrDefault();
                    if (hit != null) return hit;
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (IOException) { }
        }
        return null;
    }

    public static RequirementsSnapshot Check()
    {
        string? tess = Find("tesseract.exe");
        string? poppler = Find("pdftoppm.exe");
        bool languages = false;
        if (tess != null)
        {
            try
            {
                var psi = new ProcessStartInfo(tess, "--list-langs") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
                using var process = Process.Start(psi);
                if (process != null)
                {
                    string output = process.StandardOutput.ReadToEnd() + "\n" + process.StandardError.ReadToEnd();
                    if (process.WaitForExit(5000)) languages = output.Split('\n').Any(x => x.Trim() is "spa" or "eng");
                    else process.Kill();
                }
            }
            catch { }
        }
        bool windows = OperatingSystem.IsWindowsVersionAtLeast(10);
        return new(new List<Requirement> {
            new("Windows 10/11", "Sistema operativo compatible", windows, true),
            new(".NET 8", "Incluido cuando se publica como self-contained (Windows x64)", true, false),
            new("Tesseract OCR", tess ?? "No encontrado. Requerido para boletas escaneadas.", tess != null, false),
            new("Poppler / pdftoppm", poppler ?? "No encontrado. Convierte PDF escaneados en imágenes.", poppler != null, false),
            new("Idioma OCR (español o inglés)", languages ? "Disponible en Tesseract" : "No se encontró spa/eng o no se pudo comprobar.", languages, false)
        });
    }
}
