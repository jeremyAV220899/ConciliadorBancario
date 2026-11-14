using System.Diagnostics;
using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ConciliadorBancario;

// Modulo independiente de actualizaciones. No altera la conciliacion bancaria.
// Para abrirlo desde MainForm: UpdateCenter.ShowDialog(this);
internal static class UpdateCenter
{
    private const string Repository = "jeremyAV220899/ConciliadorBancario";
    private const string InstallerPrefix = "Setup_Conciliador_Bancario_";
    private static readonly Uri LatestRelease = new($"https://api.github.com/repos/{Repository}/releases/latest");

    public static void ShowDialog(IWin32Window? owner = null)
    {
        using var dialog = new UpdateCenterForm();
        if (owner is null) dialog.ShowDialog();
        else dialog.ShowDialog(owner);
    }

    private sealed record ReleaseInfo(Version Version, string Tag, string DownloadUrl, string Sha256, long Size, string FileName);

    private static Version InstalledVersion()
    {
        var v = typeof(UpdateCenter).Assembly.GetName().Version ?? new Version(0, 0, 0);
        return new Version(v.Major, v.Minor, Math.Max(0, v.Build));
    }

    private sealed class UpdateCenterForm : Form
    {
        private readonly Label status = new() { AutoSize = false, Height = 70, Dock = DockStyle.Top, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
        private readonly ProgressBar progress = new() { Dock = DockStyle.Top, Height = 26, Minimum = 0, Maximum = 100 };
        private readonly Label progressLabel = new() { Dock = DockStyle.Top, Height = 25, Text = "" };
        private readonly Button checkButton = new() { Text = "Buscar actualizaciones", Width = 165, Height = 33 };
        private readonly Button downloadButton = new() { Text = "Descargar e instalar", Width = 165, Height = 33, Enabled = false };
        private readonly Button cancelButton = new() { Text = "Cerrar", Width = 95, Height = 33 };
        private CancellationTokenSource? pending;
        private ReleaseInfo? release;
        private bool launching;

        public UpdateCenterForm()
        {
            Text = "Actualizaciones - Conciliador Bancario";
            Width = 530;
            Height = 255;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18), ColumnCount = 1, RowCount = 4 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(status, 0, 0);
            layout.Controls.Add(progress, 0, 1);
            layout.Controls.Add(progressLabel, 0, 2);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            buttons.Controls.Add(checkButton);
            buttons.Controls.Add(downloadButton);
            buttons.Controls.Add(cancelButton);
            layout.Controls.Add(buttons, 0, 3);
            Controls.Add(layout);
            status.Text = $"Versión instalada: {InstalledVersion()}\nComprueba si hay una nueva versión publicada en GitHub Releases.";
            checkButton.Click += async (_, _) => await CheckAsync();
            downloadButton.Click += async (_, _) => await DownloadAsync();
            cancelButton.Click += (_, _) => { if (pending is not null) pending.Cancel(); else Close(); };
            FormClosing += (_, e) => { if (pending is not null && !launching) { pending.Cancel(); e.Cancel = true; } };
            Shown += async (_, _) => await CheckAsync();
        }

        private void Busy(bool isBusy)
        {
            checkButton.Enabled = !isBusy;
            downloadButton.Enabled = !isBusy && release is not null;
            cancelButton.Text = isBusy ? "Cancelar" : "Cerrar";
        }

        private async Task CheckAsync()
        {
            if (pending is not null) return;
            release = null;
            downloadButton.Enabled = false;
            progress.Value = 0;
            progressLabel.Text = "Consultando GitHub...";
            pending = new CancellationTokenSource(TimeSpan.FromMinutes(2));
            Busy(true);
            try
            {
                using var client = CreateHttpClient();
                using var response = await client.GetAsync(LatestRelease, HttpCompletionOption.ResponseHeadersRead, pending.Token);
                if (response.StatusCode == HttpStatusCode.NotFound)
                    throw new InvalidOperationException("No hay ninguna versión publicada en GitHub Releases. Crea una Release antes de actualizar.");
                response.EnsureSuccessStatusCode();
                using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync(pending.Token));
                var root = json.RootElement;
                string tag = root.GetProperty("tag_name").GetString() ?? "";
                if (!TryParseVersion(tag, out var published))
                    throw new InvalidOperationException($"El tag '{tag}' no es una versión válida (ejemplo: v35.1.0).");
                if (published <= InstalledVersion())
                {
                    status.Text = $"Versión instalada: {InstalledVersion()}\nYa tienes la última versión publicada ({tag}).";
                    progressLabel.Text = "Sin actualizaciones pendientes.";
                    return;
                }
                var assets = root.GetProperty("assets").EnumerateArray();
                foreach (var asset in assets)
                {
                    string name = asset.GetProperty("name").GetString() ?? "";
                    if (!name.StartsWith(InstallerPrefix, StringComparison.OrdinalIgnoreCase) ||
                        !name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
                    string? url = asset.GetProperty("browser_download_url").GetString();
                    string? digest = asset.TryGetProperty("digest", out var digestJson) ? digestJson.GetString() : null;
                    if (!Uri.TryCreate(url, UriKind.Absolute, out var downloadUri) ||
                        downloadUri.Scheme != Uri.UriSchemeHttps ||
                        !downloadUri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("La dirección de descarga del instalador no es segura.");
                    // Solo se instala cuando GitHub publica el SHA256 del asset.
                    if (digest is null || !Regex.IsMatch(digest, "^sha256:[0-9a-fA-F]{64}$"))
                        throw new InvalidOperationException("La Release no proporciona SHA256 verificable para el instalador.");
                    long size = asset.GetProperty("size").GetInt64();
                    if (size <= 0 || size > 1_500_000_000)
                        throw new InvalidOperationException("El instalador excede el tamaño permitido.");
                    release = new ReleaseInfo(published, tag, url!, digest[7..].ToLowerInvariant(), size, name);
                    break;
                }
                if (release is null)
                    throw new InvalidOperationException("La Release no incluye un instalador Setup_Conciliador_Bancario_*.exe.");
                status.Text = $"Versión instalada: {InstalledVersion()}\nNueva versión disponible: {release.Tag}\nPuedes descargarla e instalarla.";
                progressLabel.Text = "Descarga lista para iniciar.";
            }
            catch (OperationCanceledException) { progressLabel.Text = "Consulta cancelada."; }
            catch (Exception ex) { status.Text = "No fue posible comprobar las actualizaciones."; progressLabel.Text = ex.Message; }
            finally { pending.Dispose(); pending = null; Busy(false); }
        }

        private async Task DownloadAsync()
        {
            if (pending is not null || release is null) return;
            var selected = release;
            var confirm = MessageBox.Show(this,
                "El Conciliador debe cerrarse para instalar la actualización. Guarda y termina cualquier conciliación en curso antes de continuar. ¿Descargar ahora?",
                "Confirmar actualización", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;
            string folder = Path.Combine(Path.GetTempPath(), "ConciliadorBancario_Updates");
            Directory.CreateDirectory(folder);
            string destination = Path.Combine(folder, "update_" + selected.Version.ToString().Replace('.', '_') + ".exe");
            string partial = destination + ".download";
            pending = new CancellationTokenSource();
            Busy(true);
            progress.Value = 0;
            try
            {
                using var client = CreateHttpClient();
                using var response = await client.GetAsync(selected.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, pending.Token);
                response.EnsureSuccessStatusCode();
                long total = response.Content.Headers.ContentLength ?? selected.Size;
                if (total <= 0 || total > 1_500_000_000) throw new InvalidOperationException("Tamaño de descarga no permitido.");
                await using (var input = await response.Content.ReadAsStreamAsync(pending.Token))
                await using (var output = new FileStream(partial, FileMode.Create, FileAccess.Write, FileShare.None, 65536, true))
                {
                    byte[] buffer = new byte[65536];
                    long readTotal = 0;
                    int count;
                    while ((count = await input.ReadAsync(buffer, pending.Token)) > 0)
                    {
                        readTotal += count;
                        if (readTotal > 1_500_000_000 || readTotal > selected.Size)
                            throw new InvalidOperationException("La descarga supera el tamaño esperado.");
                        await output.WriteAsync(buffer.AsMemory(0, count), pending.Token);
                        int pct = Math.Clamp((int)(100.0 * readTotal / total), 0, 100);
                        progress.Value = pct;
                        progressLabel.Text = $"Descargando... {pct}% ({readTotal / 1048576.0:F1} MB)";
                    }
                    if (readTotal != selected.Size) throw new InvalidOperationException("La descarga está incompleta.");
                }
                progressLabel.Text = "Verificando integridad SHA-256...";
                await using (var stream = File.OpenRead(partial))
                {
                    string sha = Convert.ToHexString(await SHA256.HashDataAsync(stream, pending.Token)).ToLowerInvariant();
                    if (!string.Equals(sha, selected.Sha256, StringComparison.Ordinal))
                        throw new InvalidOperationException("El archivo descargado no supera la comprobación SHA-256. No se ejecutará.");
                }
                File.Move(partial, destination, true);
                if (MessageBox.Show(this,
                    "Descarga y verificación correctas. Se cerrará el Conciliador y se abrirá el instalador. ¿Continuar?",
                    "Instalar actualización", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    progressLabel.Text = "Instalación pospuesta.";
                    return;
                }
                // Inno Setup muestra su propia barra de instalación. No ejecutar silenciosamente.
                Process.Start(new ProcessStartInfo
                {
                    FileName = destination,
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = folder
                });
                launching = true;
                Application.Exit();
            }
            catch (OperationCanceledException) { progressLabel.Text = "Descarga cancelada."; }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Actualización no completada", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally
            {
                try { if (File.Exists(partial)) File.Delete(partial); } catch { }
                pending.Dispose(); pending = null;
                if (!launching) Busy(false);
            }
        }

        private static bool TryParseVersion(string tag, out Version version)
        {
            var value = tag.Trim().TrimStart('v', 'V');
            var m = Regex.Match(value, @"^(\d+)\.(\d+)(?:\.(\d+))?$", RegexOptions.CultureInvariant);
            if (m.Success && Version.TryParse(value.Contains('.') && m.Groups[3].Success ? value : value + ".0", out var parsed))
            { version = new Version(parsed.Major, parsed.Minor, parsed.Build); return true; }
            version = new Version(0, 0, 0); return false;
        }

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = true, MaxAutomaticRedirections = 5 })
            {
                Timeout = TimeSpan.FromMinutes(30)
            };
            client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ConciliadorBancario", "35.0"));
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }
    }
}
