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
    private readonly HashSet<IntPtr> _subscribedDocs = new();
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
            foreach (Document doc in _docs)
            {
                UnsubscribeDocumentEvents(doc);
            }
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
            SubscribeDocumentEvents(doc);
        }
    }

    private void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = e.Document;
        SubscribeDocumentEvents(doc);
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
        UnsubscribeDocumentEvents(doc);
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

        var key = doc.UnmanagedObject;
        lock (_sync)
        {
            _docIds.Remove(key);
            _subscribedDocs.Remove(key);
        }
    }

    private void SubscribeDocumentEvents(Document? doc)
    {
        if (doc is null)
        {
            return;
        }

        var key = doc.UnmanagedObject;
        lock (_sync)
        {
            if (_subscribedDocs.Contains(key))
            {
                return;
            }

            doc.CommandWillStart += OnCommandWillStart;
            doc.CommandEnded += OnCommandEnded;
            doc.CommandCancelled += OnCommandCancelled;
            doc.CommandFailed += OnCommandFailed;
            doc.LispWillStart += OnLispWillStart;
            doc.LispEnded += OnLispEnded;
            doc.LispCancelled += OnLispCancelled;
            doc.UnknownCommand += OnUnknownCommand;
            doc.ImpliedSelectionChanged += OnImpliedSelectionChanged;
            _subscribedDocs.Add(key);
        }
    }

    private void UnsubscribeDocumentEvents(Document? doc)
    {
        if (doc is null)
        {
            return;
        }

        var key = doc.UnmanagedObject;
        lock (_sync)
        {
            if (!_subscribedDocs.Contains(key))
            {
                return;
            }

            doc.CommandWillStart -= OnCommandWillStart;
            doc.CommandEnded -= OnCommandEnded;
            doc.CommandCancelled -= OnCommandCancelled;
            doc.CommandFailed -= OnCommandFailed;
            doc.LispWillStart -= OnLispWillStart;
            doc.LispEnded -= OnLispEnded;
            doc.LispCancelled -= OnLispCancelled;
            doc.UnknownCommand -= OnUnknownCommand;
            doc.ImpliedSelectionChanged -= OnImpliedSelectionChanged;
            _subscribedDocs.Remove(key);
        }
    }

    private void OnCommandWillStart(object sender, CommandEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = sender as Document;
        var docId = EnsureDocId(doc);
        if (doc?.IsActive ?? false)
        {
            _state.ActiveDocId = docId;
        }

        var nextDepth = _state.CommandDepth + 1;
        _state.CommandDepth = nextDepth;
        _state.Busy = nextDepth > 0 || _state.LispDepth > 0;
        EnqueueCommandEvent("command_will_start", doc, docId, e.GlobalCommandName);
    }

    private void OnCommandEnded(object sender, CommandEventArgs e)
    {
        HandleCommandCompletion("command_ended", sender as Document, e.GlobalCommandName);
    }

    private void OnCommandCancelled(object sender, CommandEventArgs e)
    {
        HandleCommandCompletion("command_cancelled", sender as Document, e.GlobalCommandName);
    }

    private void OnCommandFailed(object sender, CommandEventArgs e)
    {
        HandleCommandCompletion("command_failed", sender as Document, e.GlobalCommandName);
    }

    private void HandleCommandCompletion(string eventName, Document? doc, string? commandName)
    {
        if (!_started)
        {
            return;
        }

        var docId = EnsureDocId(doc);
        var nextDepth = _state.CommandDepth;
        if (nextDepth > 0)
        {
            nextDepth -= 1;
        }

        _state.CommandDepth = nextDepth;
        _state.Busy = nextDepth > 0 || _state.LispDepth > 0;
        EnqueueCommandEvent(eventName, doc, docId, commandName);
    }

    private void OnLispWillStart(object sender, LispWillStartEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = sender as Document;
        var docId = EnsureDocId(doc);
        if (doc?.IsActive ?? false)
        {
            _state.ActiveDocId = docId;
        }

        var nextDepth = _state.LispDepth + 1;
        _state.LispDepth = nextDepth;
        _state.Busy = _state.CommandDepth > 0 || nextDepth > 0;
        EnqueueLispEvent("lisp_will_start", doc, docId, e.FirstLine);
    }

    private void OnLispEnded(object sender, EventArgs e)
    {
        HandleLispCompletion("lisp_ended", sender as Document);
    }

    private void OnLispCancelled(object sender, EventArgs e)
    {
        HandleLispCompletion("lisp_cancelled", sender as Document);
    }

    private void HandleLispCompletion(string eventName, Document? doc)
    {
        if (!_started)
        {
            return;
        }

        var docId = EnsureDocId(doc);
        var nextDepth = _state.LispDepth;
        if (nextDepth > 0)
        {
            nextDepth -= 1;
        }

        _state.LispDepth = nextDepth;
        _state.Busy = _state.CommandDepth > 0 || nextDepth > 0;
        EnqueueLispEvent(eventName, doc, docId, firstLine: null);
    }

    private void OnUnknownCommand(object sender, UnknownCommandEventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = sender as Document;
        var docId = EnsureDocId(doc);
        EnqueueAuxEvent(
            eventName: "unknown_command",
            doc: doc,
            docId: docId,
            payload: new Dictionary<string, object?>
            {
                ["name"] = e.GlobalCommandName,
                ["command_depth"] = _state.CommandDepth,
                ["lisp_depth"] = _state.LispDepth,
                ["busy"] = _state.Busy
            });
    }

    private void OnImpliedSelectionChanged(object? sender, EventArgs e)
    {
        if (!_started)
        {
            return;
        }

        var doc = sender as Document;
        var docId = EnsureDocId(doc);
        if (doc?.IsActive ?? false)
        {
            _state.ActiveDocId = docId;
        }

        EnqueueAuxEvent(
            eventName: "implied_selection_changed",
            doc: doc,
            docId: docId,
            payload: new Dictionary<string, object?>
            {
                ["command_depth"] = _state.CommandDepth,
                ["lisp_depth"] = _state.LispDepth,
                ["busy"] = _state.Busy
            });
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

    private void EnqueueCommandEvent(string eventName, Document? doc, string docId, string? commandName)
    {
        var (docName, docPath) = ResolveDocNamePath(doc);
        var ev = new EventMessage
        {
            Seq = _state.NextSeq(),
            Ts = EventJson.UtcNowIso(),
            Source = "document",
            Event = eventName,
            DocId = docId,
            DocName = docName,
            DocPath = docPath,
            Payload = new Dictionary<string, object?>
            {
                ["name"] = commandName,
                ["command_depth"] = _state.CommandDepth,
                ["lisp_depth"] = _state.LispDepth,
                ["busy"] = _state.Busy
            }
        };
        _queue.Enqueue(EventJson.SerializeEvent(ev));
    }

    private void EnqueueLispEvent(string eventName, Document? doc, string docId, string? firstLine)
    {
        var payload = new Dictionary<string, object?>
        {
            ["first_line"] = firstLine,
            ["command_depth"] = _state.CommandDepth,
            ["lisp_depth"] = _state.LispDepth,
            ["busy"] = _state.Busy
        };
        EnqueueAuxEvent(eventName, doc, docId, payload);
    }

    private void EnqueueAuxEvent(
        string eventName,
        Document? doc,
        string docId,
        Dictionary<string, object?> payload)
    {
        var (docName, docPath) = ResolveDocNamePath(doc);
        var ev = new EventMessage
        {
            Seq = _state.NextSeq(),
            Ts = EventJson.UtcNowIso(),
            Source = "document",
            Event = eventName,
            DocId = docId,
            DocName = docName,
            DocPath = docPath,
            Payload = payload
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
