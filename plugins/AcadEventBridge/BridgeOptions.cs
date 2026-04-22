using System;

namespace AcadEventBridge;

public sealed class BridgeOptions
{
    public const string ObjectEventsEnabledEnv = "AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED";

    public bool ObjectEventsEnabled { get; set; }

    public static BridgeOptions FromEnvironment()
    {
        return new BridgeOptions
        {
            ObjectEventsEnabled = ParseBoolEnv(ObjectEventsEnabledEnv, defaultValue: false)
        };
    }

    private static bool ParseBoolEnv(string name, bool defaultValue)
    {
        var raw = (Environment.GetEnvironmentVariable(name) ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(raw))
        {
            return defaultValue;
        }

        switch (raw.ToLowerInvariant())
        {
            case "1":
            case "true":
            case "yes":
            case "on":
                return true;
            case "0":
            case "false":
            case "no":
            case "off":
                return false;
            default:
                return defaultValue;
        }
    }
}
