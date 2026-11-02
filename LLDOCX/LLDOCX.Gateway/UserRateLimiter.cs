namespace LLDOCX.Gateway;

/// <summary>Sliding one-hour window per user, held in memory per instance (acceptable at lozaip.com's scale).</summary>
public sealed class UserRateLimiter
{
    static readonly TimeSpan Window = TimeSpan.FromHours(1);
    readonly int _perHour;
    readonly Func<DateTime> _utcNow;
    readonly Dictionary<string, Queue<DateTime>> _hits = new();
    readonly object _gate = new();

    public UserRateLimiter(int perHour, Func<DateTime> utcNow)
    {
        _perHour = Math.Max(1, perHour);
        _utcNow = utcNow;
    }

    public bool TryAcquire(string userId, out TimeSpan retryAfter)
    {
        lock (_gate)
        {
            DateTime now = _utcNow();
            if (!_hits.TryGetValue(userId, out var queue)) _hits[userId] = queue = new Queue<DateTime>();
            while (queue.Count > 0 && queue.Peek() <= now - Window) queue.Dequeue();
            if (queue.Count >= _perHour)
            {
                retryAfter = queue.Peek() + Window - now;
                return false;
            }
            queue.Enqueue(now);
            retryAfter = TimeSpan.Zero;
            return true;
        }
    }
}
