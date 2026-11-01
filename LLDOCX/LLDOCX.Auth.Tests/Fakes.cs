using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

internal sealed class DictionaryPolicySource : IPolicySource
{
    public Dictionary<string, object> Values { get; } = new(StringComparer.OrdinalIgnoreCase);

    public string GetString(string name) => Values.TryGetValue(name, out var v) ? v as string : null;

    public int? GetInt(string name) => Values.TryGetValue(name, out var v) && v is int i ? i : null;
}

internal sealed class FakeClock : IClock
{
    public DateTime UtcNow { get; set; } = new DateTime(2026, 11, 1, 12, 0, 0, DateTimeKind.Utc);
}

internal sealed class MemoryStore : ILastVerifiedStore
{
    public VerifiedRecord Record { get; set; }
    public VerifiedRecord Load() => Record;
    public void Save(VerifiedRecord record) => Record = record;
    public void Clear() => Record = null;
}

internal sealed class FakeTokenSource : ITokenSource
{
    public Queue<Func<RenewalOutcome>> Silent { get; } = new();
    public Func<RenewalOutcome> Interactive { get; set; }
    public Func<string> GatewayToken { get; set; } = () => "token";
    public int SilentCalls { get; private set; }
    public int SignOutCalls { get; private set; }

    public Task<RenewalOutcome> RenewSilentAsync(CancellationToken ct)
    {
        SilentCalls++;
        return Task.FromResult(Silent.Dequeue()());
    }

    public Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentWindow, CancellationToken ct) => Task.FromResult(Interactive());

    public Task<string> GetGatewayTokenAsync(CancellationToken ct) => Task.FromResult(GatewayToken());

    public Task SignOutAsync()
    {
        SignOutCalls++;
        return Task.CompletedTask;
    }
}
