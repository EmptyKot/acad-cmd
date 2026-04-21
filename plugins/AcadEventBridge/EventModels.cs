using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace AcadEventBridge;

[DataContract]
public sealed class HelloMessage
{
    [DataMember(Name = "type", Order = 1)]
    public string Type { get; set; } = "hello";

    [DataMember(Name = "protocol", Order = 2)]
    public int Protocol { get; set; } = 1;

    [DataMember(Name = "plugin", Order = 3)]
    public string Plugin { get; set; } = "AcadEventBridge";

    [DataMember(Name = "version", Order = 4)]
    public string Version { get; set; } = "0.1.0";

    [DataMember(Name = "pid", Order = 5)]
    public int Pid { get; set; }
}

[DataContract]
public sealed class EventMessage
{
    [DataMember(Name = "type", Order = 1)]
    public string Type { get; set; } = "event";

    [DataMember(Name = "seq", Order = 2)]
    public long Seq { get; set; }

    [DataMember(Name = "ts", Order = 3)]
    public string Ts { get; set; } = EventJson.UtcNowIso();

    [DataMember(Name = "source", Order = 4)]
    public string Source { get; set; } = "bridge";

    [DataMember(Name = "event", Order = 5)]
    public string Event { get; set; } = "bridge_ready";

    [DataMember(Name = "doc_id", Order = 6, EmitDefaultValue = false)]
    public string? DocId { get; set; }

    [DataMember(Name = "doc_name", Order = 7, EmitDefaultValue = false)]
    public string? DocName { get; set; }

    [DataMember(Name = "doc_path", Order = 8, EmitDefaultValue = false)]
    public string? DocPath { get; set; }

    [DataMember(Name = "payload", Order = 9)]
    public Dictionary<string, object?> Payload { get; set; } = new();
}

[DataContract]
public sealed class HeartbeatMessage
{
    [DataMember(Name = "type", Order = 1)]
    public string Type { get; set; } = "heartbeat";

    [DataMember(Name = "seq", Order = 2)]
    public long Seq { get; set; }

    [DataMember(Name = "ts", Order = 3)]
    public string Ts { get; set; } = EventJson.UtcNowIso();

    [DataMember(Name = "busy", Order = 4)]
    public bool Busy { get; set; }

    [DataMember(Name = "command_depth", Order = 5)]
    public int CommandDepth { get; set; }

    [DataMember(Name = "lisp_depth", Order = 6)]
    public int LispDepth { get; set; }

    [DataMember(Name = "active_doc_id", Order = 7, EmitDefaultValue = false)]
    public string? ActiveDocId { get; set; }

    [DataMember(Name = "queue_depth", Order = 8)]
    public int QueueDepth { get; set; }

    [DataMember(Name = "dropped_count", Order = 9)]
    public long DroppedCount { get; set; }
}

public static class EventJson
{
    private static readonly Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public static string UtcNowIso()
    {
        return DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
    }

    public static string SerializeHello(HelloMessage message)
    {
        return Serialize(message);
    }

    public static string SerializeEvent(EventMessage message)
    {
        return Serialize(message);
    }

    public static string SerializeHeartbeat(HeartbeatMessage message)
    {
        return Serialize(message);
    }

    private static string Serialize<T>(T message)
    {
        using var ms = new MemoryStream();
        var ser = new DataContractJsonSerializer(
            typeof(T),
            new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
        ser.WriteObject(ms, message);
        return Utf8.GetString(ms.ToArray());
    }
}
