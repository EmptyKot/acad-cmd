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
