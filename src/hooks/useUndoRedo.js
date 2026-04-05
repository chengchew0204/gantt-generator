import { useRef, useCallback } from 'react';

const MAX_HISTORY = 50;

/**
 * Provides undo/redo capability for an externally-managed state value.
 *
 * Returns { pushState, undo, redo, canUndo, canRedo, beginBatch, endBatch }:
 *  - pushState(snapshot) records a snapshot before a mutation
 *  - undo(currentState) / redo(currentState) return restored state or null
 *  - canUndo() / canRedo() return booleans
 *  - beginBatch(snapshot) starts a batch; only this first snapshot is recorded
 *  - endBatch() ends the batch; subsequent pushState calls resume normally
 *
 * During a batch (e.g. a drag operation), pushState is a no-op so the
 * undo stack sees the whole drag as a single action.
 */
export default function useUndoRedo() {
  const undoStack = useRef([]);
  const redoStack = useRef([]);
  const batching = useRef(false);

  const pushState = useCallback((snapshot) => {
    if (batching.current) return;
    undoStack.current = undoStack.current.slice(-(MAX_HISTORY - 1));
    undoStack.current.push(snapshot);
    redoStack.current = [];
  }, []);

  const beginBatch = useCallback((snapshot) => {
    if (batching.current) return;
    batching.current = true;
    undoStack.current = undoStack.current.slice(-(MAX_HISTORY - 1));
    undoStack.current.push(snapshot);
    redoStack.current = [];
  }, []);

  const endBatch = useCallback(() => {
    batching.current = false;
  }, []);

  const undo = useCallback((currentState) => {
    if (undoStack.current.length === 0) return null;
    const prev = undoStack.current.pop();
    redoStack.current.push(currentState);
    return prev;
  }, []);

  const redo = useCallback((currentState) => {
    if (redoStack.current.length === 0) return null;
    const next = redoStack.current.pop();
    undoStack.current.push(currentState);
    return next;
  }, []);

  const canUndo = useCallback(() => undoStack.current.length > 0, []);
  const canRedo = useCallback(() => redoStack.current.length > 0, []);

  return { pushState, undo, redo, canUndo, canRedo, beginBatch, endBatch };
}
