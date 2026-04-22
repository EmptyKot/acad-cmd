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

    [DataMember(Name = "object_events_enabled", Order = 6)]
    public bool ObjectEventsEnabled { get; set; }
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

[DataContract]
public sealed class RequestMessage
{
    [DataMember(Name = "type", Order = 1)]
    public string Type { get; set; } = "request";

    [DataMember(Name = "id", Order = 2)]
    public string Id { get; set; } = string.Empty;

    [DataMember(Name = "method", Order = 3)]
    public string Method { get; set; } = string.Empty;

    [DataMember(Name = "payload", Order = 4)]
    public Dictionary<string, object?> Payload { get; set; } = new();
}

[DataContract]
public sealed class ResponseMessage
{
    [DataMember(Name = "type", Order = 1)]
    public string Type { get; set; } = "response";

    [DataMember(Name = "id", Order = 2)]
    public string Id { get; set; } = string.Empty;

    [DataMember(Name = "ok", Order = 3)]
    public bool Ok { get; set; } = true;

    [DataMember(Name = "ts", Order = 4)]
    public string Ts { get; set; } = EventJson.UtcNowIso();

    [DataMember(Name = "payload", Order = 5, EmitDefaultValue = false)]
    public Dictionary<string, object?>? Payload { get; set; }

    [DataMember(Name = "error", Order = 6, EmitDefaultValue = false)]
    public string? Error { get; set; }
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

    public static string SerializeResponse(ResponseMessage message)
    {
        return Serialize(message);
    }

    public static bool TryDeserializeRequest(string json, out RequestMessage? request, out string? error)
    {
        request = null;
        error = null;
        if (string.IsNullOrWhiteSpace(json))
        {
            error = "empty_request";
            return false;
        }

        try
        {
            var parsed = Deserialize<RequestMessage>(json);
            if (parsed is null)
            {
                error = "request_parse_failed";
                return false;
            }

            if (!string.Equals(parsed.Type, "request", StringComparison.OrdinalIgnoreCase))
            {
                error = "invalid_request_type";
                return false;
            }

            parsed.Id = (parsed.Id ?? string.Empty).Trim();
            parsed.Method = (parsed.Method ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(parsed.Id))
            {
                error = "missing_request_id";
                return false;
            }
            if (string.IsNullOrEmpty(parsed.Method))
            {
                error = "missing_request_method";
                return false;
            }

            parsed.Payload ??= new Dictionary<string, object?>();
            request = parsed;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.GetType().Name;
            return false;
        }
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

    private static T? Deserialize<T>(string json)
    {
        var bytes = Utf8.GetBytes(json);
        using var ms = new MemoryStream(bytes);
        var ser = new DataContractJsonSerializer(
            typeof(T),
            new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
        return (T?)ser.ReadObject(ms);
    }
}
