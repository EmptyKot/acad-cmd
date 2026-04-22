using System;
using System.Collections.Generic;
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
    private readonly Func<bool> _objectEventsEnabledProvider;
    private Thread? _listenerThread;
    private volatile bool _runRequested;

    public PipeServer(
        string pipeName,
        EventQueue queue,
        StateTracker state,
        Func<bool> objectEventsEnabledProvider)
    {
        PipeName = pipeName;
        _queue = queue;
        _state = state;
        _objectEventsEnabledProvider = objectEventsEnabledProvider;
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
                    PipeDirection.InOut,
                    maxNumberOfServerInstances: 1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous | PipeOptions.WriteThrough);

                server.WaitForConnection();
                if (!_runRequested)
                {
                    continue;
                }

                _queue.Clear();
                WriteHello(server, _objectEventsEnabledProvider());

                using var heartbeatStop = new ManualResetEventSlim(false);
                var heartbeatThread = new Thread(() => RunHeartbeatProducerLoop(heartbeatStop))
                {
                    IsBackground = true,
                    Name = "AcadEventBridge.HeartbeatProducer"
                };
                heartbeatThread.Start();

                using var requestStop = new ManualResetEventSlim(false);
                var requestThread = new Thread(() => RunRequestReaderLoop(server, requestStop))
                {
                    IsBackground = true,
                    Name = "AcadEventBridge.RequestReader"
                };
                requestThread.Start();

                try
                {
                    RunWriterLoop(server);
                }
                finally
                {
                    heartbeatStop.Set();
                    requestStop.Set();
                    heartbeatThread.Join(millisecondsTimeout: 1500);
                    requestThread.Join(millisecondsTimeout: 1500);
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

    private static void WriteHello(Stream stream, bool objectEventsEnabled)
    {
        var hello = new HelloMessage
        {
            Pid = Process.GetCurrentProcess().Id,
            ObjectEventsEnabled = objectEventsEnabled
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

    private void RunRequestReaderLoop(NamedPipeServerStream stream, ManualResetEventSlim stopSignal)
    {
        stream.ReadMode = PipeTransmissionMode.Byte;

        var buffer = string.Empty;
        var chunk = new byte[2048];

        while (_runRequested && !stopSignal.IsSet)
        {
            int read;
            try
            {
                read = stream.Read(chunk, 0, chunk.Length);
            }
            catch (IOException)
            {
                LogError(new IOException("Request reader: pipe was closed while reading request stream."));
                return;
            }
            catch (ObjectDisposedException)
            {
                return;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return;
            }

            if (read <= 0)
            {
                return;
            }

            buffer += Utf8.GetString(chunk, 0, read);
            while (true)
            {
                var nl = buffer.IndexOf('\n');
                if (nl < 0)
                {
                    break;
                }

                var line = buffer.Substring(0, nl).Trim();
                buffer = buffer.Substring(nl + 1);
                if (line.Length == 0)
                {
                    continue;
                }

                HandleRequestLine(line);
            }
        }
    }

    private void HandleRequestLine(string line)
    {
        if (!EventJson.TryDeserializeRequest(line, out var request, out var parseError))
        {
            var response = new ResponseMessage
            {
                Id = request?.Id ?? string.Empty,
                Ok = false,
                Ts = EventJson.UtcNowIso(),
                Error = parseError ?? "request_parse_failed"
            };
            _queue.Enqueue(EventJson.SerializeResponse(response));
            return;
        }

        var method = request!.Method.Trim().ToLowerInvariant();
        switch (method)
        {
            case "ping":
                _queue.Enqueue(EventJson.SerializeResponse(BuildPingResponse(request.Id)));
                return;
            case "status":
                _queue.Enqueue(EventJson.SerializeResponse(BuildStatusResponse(request.Id)));
                return;
            default:
                _queue.Enqueue(EventJson.SerializeResponse(new ResponseMessage
                {
                    Id = request.Id,
                    Ok = false,
                    Ts = EventJson.UtcNowIso(),
                    Error = "unsupported_method"
                }));
                return;
        }
    }

    private ResponseMessage BuildPingResponse(string requestId)
    {
        return new ResponseMessage
        {
            Id = requestId,
            Ok = true,
            Ts = EventJson.UtcNowIso(),
            Payload = new Dictionary<string, object?>
            {
                ["pong"] = true,
                ["plugin"] = "AcadEventBridge",
                ["version"] = "0.1.0",
                ["pid"] = Process.GetCurrentProcess().Id
            }
        };
    }

    private ResponseMessage BuildStatusResponse(string requestId)
    {
        var snapshot = _state.Snapshot();
        return new ResponseMessage
        {
            Id = requestId,
            Ok = true,
            Ts = EventJson.UtcNowIso(),
            Payload = new Dictionary<string, object?>
            {
                ["pipe_name"] = PipeName,
                ["pid"] = Process.GetCurrentProcess().Id,
                ["last_seq"] = snapshot.LastSeq,
                ["busy"] = snapshot.Busy,
                ["command_depth"] = snapshot.CommandDepth,
                ["lisp_depth"] = snapshot.LispDepth,
                ["active_doc_id"] = snapshot.ActiveDocId,
                ["queue_depth"] = _queue.Count,
                ["dropped_count"] = _queue.DroppedCount,
                ["object_events_enabled"] = _objectEventsEnabledProvider()
            }
        };
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
