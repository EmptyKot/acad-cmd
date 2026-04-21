using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;

namespace AcadEventBridge;

public sealed class PipeServer
{
    private static readonly Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    private static readonly object LogSync = new();
    private const int HeartbeatIntervalMs = 2000;

    private readonly object _sync = new();
    private readonly EventQueue _queue;
    private readonly StateTracker _state;
    private Thread? _listenerThread;
    private volatile bool _runRequested;

    public PipeServer(string pipeName, EventQueue queue, StateTracker state)
    {
        PipeName = pipeName;
        _queue = queue;
        _state = state;
    }

    public string PipeName { get; }

    public bool IsRunning { get; private set; }

    public void Start()
    {
        lock (_sync)
        {
            if (IsRunning)
            {
                return;
            }

            _runRequested = true;
            _listenerThread = new Thread(ListenLoop)
            {
                IsBackground = true,
                Name = "AcadEventBridge.PipeServer"
            };
            _listenerThread.Start();
            IsRunning = true;
        }
    }

    public void Stop()
    {
        lock (_sync)
        {
            if (!IsRunning)
            {
                return;
            }

            _runRequested = false;
            SignalStop();
        }

        var t = _listenerThread;
        if (t is not null)
        {
            t.Join(millisecondsTimeout: 2000);
        }

        lock (_sync)
        {
            _listenerThread = null;
            IsRunning = false;
        }
    }

    private void ListenLoop()
    {
        while (_runRequested)
        {
            try
            {
                using var server = new NamedPipeServerStream(
                    PipeName,
                    PipeDirection.Out,
                    maxNumberOfServerInstances: 1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.WriteThrough);

                server.WaitForConnection();
                if (!_runRequested)
                {
                    continue;
                }

                _queue.Clear();
                WriteHello(server);

                using var heartbeatStop = new ManualResetEventSlim(false);
                var heartbeatThread = new Thread(() => RunHeartbeatProducerLoop(heartbeatStop))
                {
                    IsBackground = true,
                    Name = "AcadEventBridge.HeartbeatProducer"
                };
                heartbeatThread.Start();

                try
                {
                    RunWriterLoop(server);
                }
                finally
                {
                    heartbeatStop.Set();
                    heartbeatThread.Join(millisecondsTimeout: 1500);
                }
            }
            catch (IOException)
            {
                Thread.Sleep(100);
            }
            catch (ObjectDisposedException)
            {
                break;
            }
            catch (Exception ex)
            {
                LogError(ex);
                Thread.Sleep(250);
            }
        }
    }

    private static void WriteHello(Stream stream)
    {
        var hello = new HelloMessage
        {
            Pid = Process.GetCurrentProcess().Id
        };
        WriteLine(stream, EventJson.SerializeHello(hello));
    }

    private void RunHeartbeatProducerLoop(ManualResetEventSlim stopSignal)
    {
        while (_runRequested && !stopSignal.IsSet)
        {
            try
            {
                EnqueueHeartbeat();
                stopSignal.Wait(millisecondsTimeout: HeartbeatIntervalMs);
            }
            catch (Exception ex)
            {
                LogError(ex);
                stopSignal.Wait(millisecondsTimeout: 250);
            }
        }
    }

    private void EnqueueHeartbeat()
    {
        var snapshot = _state.Snapshot();
        var hb = new HeartbeatMessage
        {
            Seq = _state.NextSeq(),
            Ts = EventJson.UtcNowIso(),
            Busy = snapshot.Busy,
            CommandDepth = snapshot.CommandDepth,
            LispDepth = snapshot.LispDepth,
            ActiveDocId = snapshot.ActiveDocId,
            QueueDepth = _queue.Count + 1,
            DroppedCount = _queue.DroppedCount
        };
        _queue.Enqueue(EventJson.SerializeHeartbeat(hb));
    }

    private void RunWriterLoop(Stream stream)
    {
        while (_runRequested)
        {
            if (_queue.TryDequeue(out var line) && !string.IsNullOrWhiteSpace(line))
            {
                var lineToWrite = line!;
                WriteLine(stream, lineToWrite);
                continue;
            }

            Thread.Sleep(10);
        }
    }

    private static void WriteLine(Stream stream, string line)
    {
        var payload = Utf8.GetBytes(line + "\n");
        stream.Write(payload, 0, payload.Length);
        stream.Flush();
    }

    private void SignalStop()
    {
        try
        {
            using var client = new NamedPipeClientStream(
                ".",
                PipeName,
                PipeDirection.In,
                PipeOptions.None);
            client.Connect(timeout: 250);
        }
        catch
        {
            // Listener may not be waiting yet; ignore.
        }
    }

    private static void LogError(Exception ex)
    {
        try
        {
            lock (LogSync)
            {
                var path = Path.Combine(Path.GetTempPath(), "AcadEventBridge.pipe.log");
                var line = $"[{DateTime.UtcNow:O}] {ex.GetType().Name}: {ex.Message}{Environment.NewLine}";
                File.AppendAllText(path, line, Utf8);
            }
        }
        catch
        {
            // Never throw from logging path.
        }
    }
}
