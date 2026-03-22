using System;
using System.Collections.Generic;
using VerticalScenarioEditor.Models;

namespace VerticalScenarioEditor.History;

public sealed class DocumentHistoryService
{
    private readonly int _maxSnapshots;
    private readonly TimeSpan _inputMergeWindow;
    private readonly Func<DateTime> _utcNow;
    private readonly Stack<DocumentState> _undoStack = new();
    private readonly Stack<DocumentState> _redoStack = new();
    private InputMergeContext? _lastInputMerge;

    public DocumentHistoryService(
        int maxSnapshots = 200,
        TimeSpan? inputMergeWindow = null,
        Func<DateTime>? utcNow = null)
    {
        if (maxSnapshots <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxSnapshots));
        }

        _maxSnapshots = maxSnapshots;
        _inputMergeWindow = inputMergeWindow ?? TimeSpan.FromMilliseconds(800);
        _utcNow = utcNow ?? (() => DateTime.UtcNow);
    }

    public bool TryUndo(DocumentState currentState, out DocumentState previousState)
    {
        if (_undoStack.Count == 0)
        {
            previousState = currentState;
            return false;
        }

        _redoStack.Push(DocumentStateCloner.Clone(currentState));
        TrimToLimit(_redoStack);
        previousState = _undoStack.Pop();
        ResetInputMerge();
        return true;
    }

    public bool TryRedo(DocumentState currentState, out DocumentState nextState)
    {
        if (_redoStack.Count == 0)
        {
            nextState = currentState;
            return false;
        }

        _undoStack.Push(DocumentStateCloner.Clone(currentState));
        TrimToLimit(_undoStack);
        nextState = _redoStack.Pop();
        ResetInputMerge();
        return true;
    }

    public void PushSnapshot(DocumentState state)
    {
        PushSnapshotCore(state);
        ResetInputMerge();
    }

    public void PushInputSnapshot(DocumentState state, int recordIndex, string field)
    {
        if (string.IsNullOrWhiteSpace(field))
        {
            PushSnapshot(state);
            return;
        }

        var now = _utcNow();
        if (_lastInputMerge.HasValue
            && _lastInputMerge.Value.RecordIndex == recordIndex
            && string.Equals(_lastInputMerge.Value.Field, field, StringComparison.Ordinal)
            && now - _lastInputMerge.Value.UpdatedAtUtc <= _inputMergeWindow)
        {
            _lastInputMerge = _lastInputMerge.Value with { UpdatedAtUtc = now };
            return;
        }

        PushSnapshotCore(state);
        _lastInputMerge = new InputMergeContext(recordIndex, field, now);
    }

    public void Clear()
    {
        _undoStack.Clear();
        _redoStack.Clear();
        ResetInputMerge();
    }

    private void PushSnapshotCore(DocumentState state)
    {
        _undoStack.Push(DocumentStateCloner.Clone(state));
        TrimToLimit(_undoStack);
        _redoStack.Clear();
    }

    private void TrimToLimit(Stack<DocumentState> stack)
    {
        if (stack.Count <= _maxSnapshots)
        {
            return;
        }

        var snapshots = stack.ToArray();
        var keepCount = Math.Min(_maxSnapshots, snapshots.Length);
        stack.Clear();
        for (var index = keepCount - 1; index >= 0; index -= 1)
        {
            stack.Push(snapshots[index]);
        }
    }

    private void ResetInputMerge()
    {
        _lastInputMerge = null;
    }

    private readonly record struct InputMergeContext(int RecordIndex, string Field, DateTime UpdatedAtUtc);
}
