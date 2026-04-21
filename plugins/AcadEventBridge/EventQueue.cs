using System.Collections.Concurrent;
using System.Threading;

namespace AcadEventBridge;

public sealed class EventQueue
{
    private readonly ConcurrentQueue<string> _queue = new();
    private readonly int _capacity;
    private int _count;
    private long _droppedCount;

    public EventQueue(int capacity)
    {
        _capacity = capacity < 1 ? 1 : capacity;
    }

    public int Count => Volatile.Read(ref _count);

    public int Capacity => _capacity;

    public long DroppedCount => Interlocked.Read(ref _droppedCount);

    public void Enqueue(string message)
    {
        _queue.Enqueue(message);
        var count = Interlocked.Increment(ref _count);

        while (count > _capacity && _queue.TryDequeue(out _))
        {
            Interlocked.Decrement(ref _count);
            Interlocked.Increment(ref _droppedCount);
            count = Volatile.Read(ref _count);
        }
    }

    public bool TryDequeue(out string? message)
    {
        if (_queue.TryDequeue(out var m))
        {
            Interlocked.Decrement(ref _count);
            message = m;
            return true;
        }

        message = null;
        return false;
    }

    public void Clear()
    {
        while (_queue.TryDequeue(out _))
        {
            Interlocked.Decrement(ref _count);
        }
    }
}
