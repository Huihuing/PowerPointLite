using System;
using System.IO;

namespace PptxViewer
{
    internal sealed class PresentationEditSession
    {
        public PresentationDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }
        public DateTime LastSavedUtc { get; private set; }

        private PresentationEditSession(PresentationDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            FilePath = string.Empty;
            IsDirty = true;
            LastSavedUtc = DateTime.MinValue;
        }

        public static PresentationEditSession CreateNew(string title)
        {
            return new PresentationEditSession(
                PresentationDocument.CreateNew(title));
        }

        public PresentationSlide AddSlide(string title)
        {
            PresentationSlide slide = Document.AddSlide(title);
            MarkDirty();
            return slide;
        }

        public bool RemoveSlide(int index)
        {
            bool removed = Document.RemoveSlide(index);

            if (removed)
                MarkDirty();

            return removed;
        }

        public bool MoveSlide(int fromIndex, int toIndex)
        {
            bool moved = Document.MoveSlide(fromIndex, toIndex);

            if (moved)
                MarkDirty();

            return moved;
        }

        public bool SetText(int slideIndex, int textBoxIndex, string text)
        {
            PresentationTextBox box = GetTextBox(slideIndex, textBoxIndex);

            if (box == null)
                return false;

            string newValue = text ?? string.Empty;

            if (string.Equals(box.Text, newValue, StringComparison.Ordinal))
                return true;

            box.Text = newValue;
            MarkDirty();
            return true;
        }

        public bool SetFont(
            int slideIndex,
            int textBoxIndex,
            string fontFamily,
            float fontSizePoints,
            bool bold,
            bool italic)
        {
            PresentationTextBox box = GetTextBox(slideIndex, textBoxIndex);

            if (box == null)
                return false;

            box.FontFamily = string.IsNullOrEmpty(fontFamily)
                ? box.FontFamily
                : fontFamily;
            box.FontSizePoints = Math.Max(1f, Math.Min(400f, fontSizePoints));
            box.Bold = bold;
            box.Italic = italic;
            MarkDirty();
            return true;
        }

        public void MarkDirty()
        {
            IsDirty = true;
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("Save As is required for a new presentation.");

            PptxWriter.Save(Document, FilePath);
            IsDirty = false;
            LastSavedUtc = DateTime.UtcNow;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");

            string extension = Path.GetExtension(path);

            if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase))
                path += ".pptx";

            PptxWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
            LastSavedUtc = DateTime.UtcNow;
        }

        private PresentationTextBox GetTextBox(int slideIndex, int textBoxIndex)
        {
            if (slideIndex < 0 || slideIndex >= Document.Slides.Count)
                return null;

            PresentationSlide slide = Document.Slides[slideIndex];

            if (slide == null ||
                textBoxIndex < 0 ||
                textBoxIndex >= slide.TextBoxes.Count)
            {
                return null;
            }

            return slide.TextBoxes[textBoxIndex];
        }
    }
}
