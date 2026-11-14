using System.Diagnostics;
using System.Net;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ConciliadorBancario;

// Actualizador aislado: no modifica el motor de conciliación ni los documentos del usuario.
internal static class UpdateService
{
    private const string ReleaseApi = "https://api.github.com/repos/jeremyAV220899/ConciliadorBancario/releases/latest";
    private const string InstallerName = "Setup_Conciliador_Bancario_V35.exe";
    private const long MaximumInstallerSize = 1_073_741_824; // 1 GiB
    internal sealed record AvailableRelease(Version Version, string Tag, Uri Installer, Uri Checksum);

    internal static Version CurrentVersion =>
        Assembly.GetExecutingAssembly().GetName().Version is { } v
            ? new Version(v.Major, v.Minor, Math.Max(0, v.Build))
            : new Version(0, 0, 0);

    private static HttpClient NewClient()
    {
        var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = true })
        {
            Timeout = TimeSpan.FromMinutes(30)
        };
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ConciliadorBancario", "35"));
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        return client;
    }

    internal static async Task<AvailableRelease?> LatestAsync(CancellationToken ct)
    {
        using var client = NewClient();
        using var response = await client.GetAsync(ReleaseApi, HttpCompletionOption.ResponseHeadersRead, ct);
        if (response.StatusCode == HttpStatusCode.NotFound)
            return null; // Aún no existe un Release publicado.
        response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        using var json = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
        var release = json.RootElement;
        if (!release.TryGetProperty("tag_name", out var tagProp))
            throw new InvalidDataException("La publicación de GitHub no contiene un número de versión.");
        string tag = tagProp.GetString() ?? "";
        string number = tag.TrimStart('v', 'V');
        if (!Version.TryParse(number, out var version))
            throw new InvalidDataException($"Versión desconocida en GitHub: {tag}");
        if (!release.TryGetProperty("assets", out var assets))
            throw new InvalidDataException("La publicación de GitHub no incluye archivos de instalación.");
        Uri? setup = null, checksum = null;
        foreach (var asset in assets.EnumerateArray())
        {
            string name = asset.GetProperty("name").GetString() ?? "";
            string url = asset.GetProperty("browser_download_url").GetString() ?? "";
            if (!Uri.TryCreate(url, UriKind.Absolute, out var parsed) ||
                parsed.Scheme != Uri.UriSchemeHttps ||
                !parsed.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
                continue;
            if (name.Equals(InstallerName, StringComparison.OrdinalIgnoreCase)) setup = parsed;
            if (name.Equals(InstallerName + ".sha256", StringComparison.OrdinalIgnoreCase)) checksum = parsed;
        }
        if (setup is null || checksum is null)
            throw new InvalidDataException("La versión publicada debe contener el Setup.exe y su archivo .sha256.");
        return new AvailableRelease(version, tag, setup, checksum);
    }

    internal static async Task<string> DownloadAsync(AvailableRelease release, IProgress<int> progress, CancellationToken ct)
    {
        using var client = NewClient();
        string hashText = await client.GetStringAsync(release.Checksum, ct);
        var match = Regex.Match(hashText, @"\b[a-fA-F0-9]{64}\b");
        if (!match.Success) throw new InvalidDataException("El SHA256 publicado no es válido.");
        string expectedHash = match.Value.ToUpperInvariant();

        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ConciliadorBancario", "Actualizaciones", release.Version.ToString());
        Directory.CreateDirectory(dir);
        string destination = Path.Combine(dir, InstallerName);
        string pending = destination + ".download";
        try
        {
            using var response = await client.GetAsync(release.Installer, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();
            long? size = response.Content.Headers.ContentLength;
            if (size is > MaximumInstallerSize)
                throw new InvalidDataException("El instalador supera el tamaño máximo admitido.");
            await using (var input = await response.Content.ReadAsStreamAsync(ct))
            await using (var output = new FileStream(pending, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
            {
                var buffer = new byte[81920];
                long total = 0;
                int count;
                while ((count = await input.ReadAsync(buffer, ct)) != 0)
                {
                    total += count;
                    if (total > MaximumInstallerSize)
                        throw new InvalidDataException("El instalador supera el tamaño máximo admitido.");
                    await output.WriteAsync(buffer.AsMemory(0, count), ct);
                    if (size is > 0)
                        progress.Report((int)Math.Clamp(total * 100 / size.Value, 0, 100));
                }
            }
            await using (var input = File.OpenRead(pending))
            {
                byte[] digest = await SHA256.HashDataAsync(input, ct);
                string actualHash = Convert.ToHexString(digest);
                if (!actualHash.Equals(expectedHash, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("La descarga no coincide con el SHA256 publicado. Se canceló la instalación.");
            }
            // Comprobación mínima del formato ejecutable de Windows.
            using (var input = File.OpenRead(pending))
                if (input.ReadByte() != 'M' || input.ReadByte() != 'Z')
                    throw new InvalidDataException("El archivo descargado no es un ejecutable de Windows válido.");
            File.Move(pending, destination, true);
            progress.Report(100);
            return destination;
        }
        finally
        {
            try { if (File.Exists(pending)) File.Delete(pending); } catch { }
        }
    }

    internal static void StartInstaller(string exePath)
    {
        // Inno Setup instala sobre la versión anterior gracias al mismo AppId.
        // No usar instalación silenciosa: el usuario debe confirmar UAC y el asistente.
        if (Process.Start(new ProcessStartInfo(exePath) { UseShellExecute = true, Verb = "runas" }) is null)
            throw new InvalidOperationException("Windows no pudo iniciar el instalador.");
    }
}

internal sealed class UpdateForm : Form
{
    private readonly Label _info = new() { Dock = DockStyle.Fill, AutoSize = false };
    private readonly Label _detail = new() { Dock = DockStyle.Fill, AutoSize = false };
    private readonly ProgressBar _bar = new() { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100 };
    private readonly Button _action = new() { Text = "Comprobar versión", Dock = DockStyle.Fill };
    private readonly Button _cancel = new() { Text = "Cerrar", Dock = DockStyle.Fill };
    private CancellationTokenSource? _downloadCancellation;
    private UpdateService.AvailableRelease? _available;
    private bool _busy;
    private bool _launched;

    internal UpdateForm()
    {
        Text = "Actualizaciones del Conciliador Bancario";
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(535, 260);
        MinimumSize = new Size(450, 260);
        Font = new Font("Segoe UI", 10);
        Padding = new Padding(20);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 12));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        _info.Text = "Versión instalada: " + UpdateService.CurrentVersion;
        _info.Font = new Font("Segoe UI Semibold", 12, FontStyle.Bold);
        _detail.Text = "Pulse 'Comprobar versión' para consultar GitHub Releases.\nNo se descargará nada sin su autorización.";
        _bar.Value = 0;
        layout.Controls.Add(_info, 0, 0);
        layout.Controls.Add(_detail, 0, 1);
        layout.Controls.Add(_bar, 0, 2);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
        _action.Size = new Size(188, 34);
        _cancel.Size = new Size(110, 34);
        buttons.Controls.Add(_action);
        buttons.Controls.Add(_cancel);
        layout.Controls.Add(buttons, 0, 4);
        Controls.Add(layout);
        _action.Click += async (_, _) => await OnActionAsync();
        _cancel.Click += (_, _) => { if (_busy) _downloadCancellation?.Cancel(); else Close(); };
        FormClosing += (_, e) => { if (_busy && !_launched) { e.Cancel = true; _downloadCancellation?.Cancel(); } };
        Shown += async (_, _) => await CheckAsync();
    }

    private async Task OnActionAsync()
    {
        if (_busy) return;
        if (_available is null) { await CheckAsync(); return; }
        if (MessageBox.Show(this,
            $"Se descargará la versión {_available.Version} y se iniciará su instalador.\n\n" +
            "Guarde sus archivos antes de continuar. ¿Desea actualizar ahora?",
            "Confirmar actualización", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _busy = true;
        _action.Enabled = false;
        _cancel.Text = "Cancelar";
        using var cts = new CancellationTokenSource();
        _downloadCancellation = cts;
        try
        {
            _detail.Text = "Descargando instalador...";
            var progress = new Progress<int>(p => { _bar.Value = Math.Clamp(p, 0, 100); _detail.Text = $"Descargando... {p}%"; });
            var path = await UpdateService.DownloadAsync(_available, progress, cts.Token);
            _detail.Text = "Descarga completa y SHA256 verificado. Abriendo instalador...";
            UpdateService.StartInstaller(path);
            _launched = true;
            MessageBox.Show(this, "Se inició el instalador. El Conciliador se cerrará para permitir la actualización.",
                "Instalador iniciado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Owner?.Close();
            Close();
        }
        catch (OperationCanceledException) { _detail.Text = "Descarga cancelada; no se modificó la instalación."; }
        catch (Exception ex) { ShowError(ex); }
        finally
        {
            _busy = false;
            _downloadCancellation = null;
            if (!_launched) { _action.Enabled = _available is not null; _cancel.Text = "Cerrar"; }
        }
    }

    private async Task CheckAsync()
    {
        if (_busy) return;
        _busy = true;
        _action.Enabled = false;
        _detail.Text = "Consultando última versión publicada...";
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(25));
            var release = await UpdateService.LatestAsync(cts.Token);
            if (release is null)
            {
                _detail.Text = "Todavía no hay versiones publicadas en GitHub Releases.";
                _available = null;
            }
            else if (release.Version <= UpdateService.CurrentVersion)
            {
                _detail.Text = $"Está actualizado. Última versión publicada: {release.Tag}.";
                _available = null;
            }
            else
            {
                _available = release;
                _detail.Text = $"Nueva versión disponible: {release.Tag}.\nPulse 'Descargar e instalar' para actualizar.";
                _action.Text = "Descargar e instalar";
                _action.Enabled = true;
            }
        }
        catch (Exception ex) { ShowError(ex); }
        finally
        {
            _busy = false;
            if (_available is null) { _action.Text = "Volver a comprobar"; _action.Enabled = true; }
        }
    }

    private void ShowError(Exception ex)
    {
        _detail.Text = "No se pudo completar la operación: " + ex.Message;
        MessageBox.Show(this, ex.Message, "Actualización no realizada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
