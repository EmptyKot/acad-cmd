using System.Threading;

namespace AcadEventBridge;

public readonly struct StateSnapshot
{
    public StateSnapshot(long lastSeq, bool busy, int commandDepth, int lispDepth, string? activeDocId)
    {
        LastSeq = lastSeq;
        Busy = busy;
        CommandDepth = commandDepth;
        LispDepth = lispDepth;
        ActiveDocId = activeDocId;
    }

    public long LastSeq { get; }

    public bool Busy { get; }

    public int CommandDepth { get; }

    public int LispDepth { get; }

    public string? ActiveDocId { get; }
}

public sealed class StateTracker
{
    private long _seq;
    private int _busy;
    private int _commandDepth;
    private int _lispDepth;
    private readonly object _sync = new();
    private string? _activeDocId;

    public long LastSeq => Interlocked.Read(ref _seq);

    public bool Busy
    {
        get => Volatile.Read(ref _busy) != 0;
        set => Interlocked.Exchange(ref _busy, value ? 1 : 0);
    }

    public int CommandDepth
    {
        get => Volatile.Read(ref _commandDepth);
        set => Interlocked.Exchange(ref _commandDepth, value);
    }

    public int LispDepth
    {
        get => Volatile.Read(ref _lispDepth);
        set => Interlocked.Exchange(ref _lispDepth, value);
    }

    public string? ActiveDocId
    {
        get
        {
            lock (_sync)
            {
                return _activeDocId;
            }
        }
        set
        {
            lock (_sync)
            {
                _activeDocId = value;
            }
        }
    }

    public long NextSeq()
    {
        return Interlocked.Increment(ref _seq);
    }

    public StateSnapshot Snapshot()
    {
        string? activeDocId;
        lock (_sync)
        {
            activeDocId = _activeDocId;
        }

        return new StateSnapshot(
            lastSeq: LastSeq,
            busy: Busy,
            commandDepth: CommandDepth,
            lispDepth: LispDepth,
            activeDocId: activeDocId);
    }
}
