using PdfSharp.Fonts;

namespace ConciliadorBancario;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        // PDFsharp 6.2 Core: permitir resolver Arial desde fuentes instaladas en Windows.
        // Debe ejecutarse antes de crear cualquier XFont.
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}
