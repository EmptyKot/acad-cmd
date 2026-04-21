using Autodesk.AutoCAD.Runtime;

namespace AcadEventBridge;

public sealed class EntryPoint : IExtensionApplication
{
    private BridgeHost? _host;

    public void Initialize()
    {
        _host = new BridgeHost();
        _host.Start();
        BridgeRuntime.CurrentHost = _host;
    }

    public void Terminate()
    {
        if (_host is null)
        {
            return;
        }

        _host.Stop();
        BridgeRuntime.CurrentHost = null;
        _host = null;
    }
}
