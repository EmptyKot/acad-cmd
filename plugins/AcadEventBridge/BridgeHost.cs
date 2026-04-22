using System.Diagnostics;

namespace AcadEventBridge;

public sealed class BridgeHost
{
    private readonly BridgeOptions _options;
    private readonly EventQueue _queue;
    private readonly StateTracker _state;
    private readonly PipeServer _pipeServer;
    private readonly EventRegistrar _registrar;
    private bool _started;

    public BridgeHost()
    {
        _options = BridgeOptions.FromEnvironment();
        _queue = new EventQueue(capacity: 1024);
        _state = new StateTracker();
        _registrar = new EventRegistrar(
            _queue,
            _state,
            objectEventsEnabled: _options.ObjectEventsEnabled);
        _pipeServer = new PipeServer(
            GetPipeName(),
            _queue,
            _state,
            objectEventsEnabledProvider: () => _registrar.ObjectEventsEnabled);
    }

    public string PipeName => _pipeServer.PipeName;

    public bool IsPipeServerRunning => _pipeServer.IsRunning;

    public long LastSeq => _state.LastSeq;

    public bool Busy => _state.Busy;

    public int CommandDepth => _state.CommandDepth;

    public int LispDepth => _state.LispDepth;

    public int QueueDepth => _queue.Count;

    public long DroppedCount => _queue.DroppedCount;

    public bool ObjectEventsEnabled => _registrar.ObjectEventsEnabled;

    public void SetObjectEventsEnabled(bool enabled)
    {
        _registrar.SetObjectEventsEnabled(enabled);
    }

    public void Start()
    {
        if (_started)
        {
            return;
        }

        _pipeServer.Start();
        _registrar.Start();
        _started = true;
    }

    public void Stop()
    {
        if (!_started)
        {
            return;
        }

        _registrar.Stop();
        _pipeServer.Stop();
        _started = false;
    }

    private static string GetPipeName()
    {
        var pid = Process.GetCurrentProcess().Id;
        return $"acad-event-bridge-{pid}";
    }
}
