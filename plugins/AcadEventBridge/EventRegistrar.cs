using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AcadEventBridge;

public sealed class EventRegistrar
{
    private readonly EventQueue _queue;
    private readonly StateTracker _state;
    private readonly object _sync = new();
    private readonly Dictionary<IntPtr, string> _docIds = new();
    private DocumentCollection? _docs;
    private bool _started;

    public EventRegistrar(EventQueue queue, StateTracker state)
    {
        _queue = queue;
        _state = state;
    }

    public bool IsStarted => _started;

    public int QueueCapacity => _queue.Capacity;

    public void Start()
    {
        if (_started)
        {
            return;
        }

        _docs = AcApplication.DocumentManager;
        Subscribe(_docs);
        RegisterExistingDocuments(_docs);

        _state.Busy = false;
        _state.CommandDepth = 0;
        _state.LispDepth = 0;
        _state.ActiveDocId = TryGetDocId(_docs.MdiActiveDocument);
        _started = true;
    }

    public void Stop()
    {
        if (!_started)
        {
            return;
        }

        if (_docs is not null)
        {
            Unsubscribe(_docs);
            _docs = null;
        }

        lock (_sync)
        {
            _docIds.Clear();
        }
        _state.ActiveDocId = null;
        _started = false;
    }

    private void Subscribe(DocumentCollection docs)
    {
        docs.DocumentCreated += OnDocumentCreated;
        docs.DocumentBecameCurrent += OnDocumentBecameCurrent;
        docs.DocumentToBeDestroyed += OnDocumentToBeDestroyed;
    }

    private void Unsubscribe(DocumentCollection docs)
    {
        docs.DocumentCreated -= OnDocumentCreated;
        docs.DocumentBecameCurrent -= OnDocumentBecameCurrent;
        docs.DocumentToBeDestroyed -= OnDocumentToBeDestroyed;
    }

    private void RegisterExistingDocuments(DocumentCollection docs)
    {
        foreach (Document doc in docs)
        {
            _ = EnsureDocId(doc);
        }
    }

    private void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = e.Document;
        var docId = EnsureDocId(doc);
        EnqueueDocumentEvent("document_created", doc, docId);
    }

    private void OnDocumentBecameCurrent(object sender, DocumentCollectionEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = e.Document;
        var docId = EnsureDocId(doc);
        _state.ActiveDocId = docId;
        EnqueueDocumentEvent("document_activated", doc, docId);
    }

    private void OnDocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = e.Document;
        var docId = EnsureDocId(doc);
        EnqueueDocumentEvent("document_destroyed", doc, docId);
        RemoveDoc(doc);

        if (string.Equals(_state.ActiveDocId, docId, StringComparison.Ordinal))
        {
            _state.ActiveDocId = null;
        }
    }

    private string EnsureDocId(Document? doc)
    {
        if (doc is null)
        {
            return "unknown";
        }

        var key = doc.UnmanagedObject;
        lock (_sync)
        {
            if (_docIds.TryGetValue(key, out var existing))
            {
                return existing;
            }

            var docId = Guid.NewGuid().ToString();
            _docIds[key] = docId;
            return docId;
        }
    }

    private string? TryGetDocId(Document? doc)
    {
        if (doc is null)
        {
            return null;
        }

        return EnsureDocId(doc);
    }

    private void RemoveDoc(Document? doc)
    {
        if (doc is null)
        {
            return;
        }

        lock (_sync)
        {
            _docIds.Remove(doc.UnmanagedObject);
        }
    }

    private void EnqueueDocumentEvent(string eventName, Document? doc, string docId)
    {
        var (docName, docPath) = ResolveDocNamePath(doc);
        var ev = new EventMessage
        {
            Seq = _state.NextSeq(),
            Ts = EventJson.UtcNowIso(),
            Source = "document_collection",
            Event = eventName,
            DocId = docId,
            DocName = docName,
            DocPath = docPath,
            Payload = new Dictionary<string, object?>
            {
                ["is_active"] = doc?.IsActive ?? false,
                ["is_named_drawing"] = doc?.IsNamedDrawing ?? false
            }
        };
        _queue.Enqueue(EventJson.SerializeEvent(ev));
    }

    private static (string? docName, string? docPath) ResolveDocNamePath(Document? doc)
    {
        var rawName = doc?.Name;
        if (string.IsNullOrWhiteSpace(rawName))
        {
            return (null, null);
        }

        if (Path.IsPathRooted(rawName))
        {
            return (Path.GetFileName(rawName), rawName);
        }

        return (rawName, null);
    }
}
