using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace AcadEventBridge;

public sealed class DebugCommands
{
    [CommandMethod("AEB_STATUS")]
    public void AebStatus()
    {
        var host = BridgeRuntime.CurrentHost;
        if (host is null)
        {
            WriteLine("AcadEventBridge loaded, host is null.");
            return;
        }

        WriteLine(
            "AcadEventBridge loaded. " +
            $"pipe_name={host.PipeName}, " +
            $"pipe_running={host.IsPipeServerRunning}, " +
            $"last_seq={host.LastSeq}, " +
            $"busy={host.Busy}, " +
            $"command_depth={host.CommandDepth}, " +
            $"lisp_depth={host.LispDepth}, " +
            $"object_events_enabled={host.ObjectEventsEnabled}, " +
            $"queue_depth={host.QueueDepth}, " +
            $"dropped_count={host.DroppedCount}.");
    }

    [CommandMethod("AEB_PING")]
    public void AebPing()
    {
        WriteLine("AcadEventBridge ping: ok.");
    }

    [CommandMethod("AEB_DIAG")]
    public void AebDiag()
    {
        WriteLine("AcadEventBridge diagnostics: queue/state tracker active.");
    }

    [CommandMethod("AEB_OBJECT_EVENTS_ON")]
    public void AebObjectEventsOn()
    {
        var host = BridgeRuntime.CurrentHost;
        if (host is null)
        {
            WriteLine("AcadEventBridge host is null; cannot enable object events.");
            return;
        }

        host.SetObjectEventsEnabled(true);
        WriteLine("AcadEventBridge object events: enabled.");
    }

    [CommandMethod("AEB_OBJECT_EVENTS_OFF")]
    public void AebObjectEventsOff()
    {
        var host = BridgeRuntime.CurrentHost;
        if (host is null)
        {
            WriteLine("AcadEventBridge host is null; cannot disable object events.");
            return;
        }

        host.SetObjectEventsEnabled(false);
        WriteLine("AcadEventBridge object events: disabled.");
    }

    private static void WriteLine(string message)
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        Editor ed = doc.Editor;
        ed.WriteMessage($"\n{message}");
    }
}
