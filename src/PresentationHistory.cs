using System;
using System.Collections.Generic;

namespace PptxViewer
{
    internal sealed class PresentationHistory
    {
        private readonly Stack<PresentationDocument> undo =
            new Stack<PresentationDocument>();
        private readonly Stack<PresentationDocument> redo =
            new Stack<PresentationDocument>();
        private readonly int capacity;

        public bool CanUndo
        {
            get { return undo.Count > 0; }
        }

        public bool CanRedo
        {
            get { return redo.Count > 0; }
        }

        public PresentationHistory(int maxSnapshots)
        {
            capacity = Math.Max(5, maxSnapshots);
        }

        public void Capture(PresentationDocument document)
        {
            if (document == null)
                return;

            undo.Push(document.Clone());
            redo.Clear();
            Trim(undo);
        }

        public PresentationDocument Undo(PresentationDocument current)
        {
            if (!CanUndo || current == null)
                return null;

            redo.Push(current.Clone());
            Trim(redo);
            return undo.Pop();
        }

        public PresentationDocument Redo(PresentationDocument current)
        {
            if (!CanRedo || current == null)
                return null;

            undo.Push(current.Clone());
            Trim(undo);
            return redo.Pop();
        }

        public void Clear()
        {
            undo.Clear();
            redo.Clear();
        }

        private void Trim(Stack<PresentationDocument> stack)
        {
            if (stack.Count <= capacity)
                return;

            PresentationDocument[] values = stack.ToArray();
            stack.Clear();

            int count = Math.Min(capacity, values.Length);
            for (int i = count - 1; i >= 0; i--)
                stack.Push(values[i]);
        }
    }
}
