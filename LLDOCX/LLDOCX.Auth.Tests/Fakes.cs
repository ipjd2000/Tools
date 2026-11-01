using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

internal sealed class DictionaryPolicySource : IPolicySource
{
    public Dictionary<string, object> Values { get; } = new(StringComparer.OrdinalIgnoreCase);

    public string GetString(string name) => Values.TryGetValue(name, out var v) ? v as string : null;

    public int? GetInt(string name) => Values.TryGetValue(name, out var v) && v is int i ? i : null;
}
