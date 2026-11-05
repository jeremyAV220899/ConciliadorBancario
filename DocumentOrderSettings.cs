using System.Text.Json;

namespace ConciliadorBancario;

/// <summary>Dos órdenes independientes; pagos siempre antes que adicionales.</summary>
internal static class DocumentOrderSettings
{
    internal static readonly string[] PaymentOrder =
    {
        "BOLETA", "PAGO 1", "PAGO 2", "PAGO 3", "PAGO CTS", "PAGO GRATIFICACIÓN",
        "VACACIONES", "LICENCIA", "SUBSIDIOS", "DESCANSO MEDICO", "PAYFLOW"
    };
    internal static readonly string[] AdditionalOrder = { "CTR", "VIDA LEY" };
    internal static readonly string[] DefaultOrder = PaymentOrder.Concat(AdditionalOrder).ToArray();

    private sealed class StoredOrders
    {
        public string[]? Pagos { get; set; }
        public string[]? Adicionales { get; set; }
    }

    private static string Folder => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConciliadorBancario");
    private static string SettingsPath => Path.Combine(Folder, "orden_documentos_bloques.json");
    private static string LegacyPath => Path.Combine(Folder, "orden_documentos.json");

    internal static (IReadOnlyList<string> Pagos, IReadOnlyList<string> Adicionales) LoadGroups()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var stored = JsonSerializer.Deserialize<StoredOrders>(File.ReadAllText(SettingsPath));
                if (stored != null)
                    return (NormalizeGroup(stored.Pagos ?? [], PaymentOrder),
                            NormalizeGroup(stored.Adicionales ?? [], AdditionalOrder));
            }
            // Migración transparente de V33/V34: conserva el orden relativo de cada grupo.
            if (File.Exists(LegacyPath))
            {
                var legacy = JsonSerializer.Deserialize<string[]>(File.ReadAllText(LegacyPath));
                if (legacy != null)
                    return (NormalizeGroup(legacy, PaymentOrder), NormalizeGroup(legacy, AdditionalOrder));
            }
        }
        catch (Exception) { /* Configuración inválida: valores predeterminados */ }
        return (PaymentOrder.ToArray(), AdditionalOrder.ToArray());
    }

    internal static IReadOnlyList<string> Load()
    {
        var groups = LoadGroups();
        return groups.Pagos.Concat(groups.Adicionales).ToArray();
    }

    internal static IReadOnlyList<string> Normalize(IEnumerable<string> values)
    {
        var valuesArray = values.ToArray();
        return NormalizeGroup(valuesArray, PaymentOrder).Concat(NormalizeGroup(valuesArray, AdditionalOrder)).ToArray();
    }

    internal static IReadOnlyList<string> NormalizeGroup(IEnumerable<string> values, IReadOnlyList<string> defaults)
    {
        var allowed = new HashSet<string>(defaults, StringComparer.Ordinal);
        var result = new List<string>();
        foreach (var value in values)
            if (value is not null && allowed.Contains(value) && !result.Contains(value, StringComparer.Ordinal))
                result.Add(value);
        foreach (var value in defaults)
            if (!result.Contains(value, StringComparer.Ordinal)) result.Add(value);
        return result;
    }

    internal static void SaveGroups(IEnumerable<string> pagos, IEnumerable<string> adicionales)
    {
        var data = new StoredOrders
        {
            Pagos = NormalizeGroup(pagos, PaymentOrder).ToArray(),
            Adicionales = NormalizeGroup(adicionales, AdditionalOrder).ToArray()
        };
        Directory.CreateDirectory(Folder);
        string tmp = SettingsPath + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true }));
        File.Move(tmp, SettingsPath, true);
    }

    internal static void Save(IEnumerable<string> values)
    {
        var valuesArray = values.ToArray();
        SaveGroups(valuesArray, valuesArray);
    }
}
