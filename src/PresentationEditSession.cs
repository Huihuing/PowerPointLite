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

        public static PresentationEditSession OpenEditable(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("PPTX file was not found.", path);

            PptxEditableLoadResult loaded = PptxEditableReader.Read(path);

            if (loaded == null || loaded.Document == null)
                throw new InvalidDataException("The presentation could not be loaded into the editable model.");

            if (!loaded.CanRoundTripSafely)
            {
                throw new InvalidOperationException(
                    string.IsNullOrEmpty(loaded.Warning)
                        ? "This presentation cannot yet be edited without risking unsupported-content loss."
                        : loaded.Warning);
            }

            // The conservative base reader establishes whether the package is
            // safe for editing. Rich text is layered on only after that safety
            // decision so unsupported external files are never promoted into
            // editable mode merely because their text can be parsed.
            PptxRichTextPackage.ReadIntoDocument(path, loaded.Document);

            PresentationEditSession session =
                new PresentationEditSession(loaded.Document);

            session.FilePath = Path.GetFullPath(path);
            session.IsDirty = false;
            session.LastSavedUtc = File.GetLastWriteTimeUtc(path);
            return session;
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

        public PresentationTextBox AddTextBox(int slideIndex, string text)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            if (slide == null)
                return null;

            PresentationTextBox box = slide.AddTextBox(text);
            MarkDirty();
            return box;
        }

        public PresentationShape AddShape(
            int slideIndex,
            PresentationShapeKind kind)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            if (slide == null)
                return null;

            PresentationShape shape = slide.AddShape(kind);
            MarkDirty();
            return shape;
        }

        public PresentationImage AddImage(
            int slideIndex,
            byte[] data,
            string extension,
            string contentType)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            if (slide == null)
                return null;

            PresentationImage image = slide.AddImage(
                data,
                extension,
                contentType);
            MarkDirty();
            return image;
        }

        public PresentationTable AddTable(
            int slideIndex,
            int rows,
            int columns)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            if (slide == null)
                return null;

            PresentationTable table = slide.AddTable(rows, columns);
            MarkDirty();
            return table;
        }

        public bool RemoveTextBox(int slideIndex, int index)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            bool removed = slide != null && slide.RemoveTextBox(index);
            if (removed)
                MarkDirty();
            return removed;
        }

        public bool RemoveShape(int slideIndex, int index)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            bool removed = slide != null && slide.RemoveShape(index);
            if (removed)
                MarkDirty();
            return removed;
        }

        public bool RemoveImage(int slideIndex, int index)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            bool removed = slide != null && slide.RemoveImage(index);
            if (removed)
                MarkDirty();
            return removed;
        }

        public bool RemoveTable(int slideIndex, int index)
        {
            PresentationSlide slide = GetSlide(slideIndex);
            bool removed = slide != null && slide.RemoveTable(index);
            if (removed)
                MarkDirty();
            return removed;
        }

        public bool SetText(int slideIndex, int textBoxIndex, string text)
        {
            PresentationTextBox box = GetTextBox(slideIndex, textBoxIndex);
            if (box == null)
                return false;

            string newValue = text ?? string.Empty;
            if (string.Equals(box.Text, newValue, StringComparison.Ordinal))
                return true;

            // Plain text editing intentionally collapses run-level formatting
            // rather than silently applying stale rich runs to changed text.
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

            // The existing property panel edits the whole text box. When it is
            // used on rich text, apply those four properties to every run but
            // keep underline, baseline, colors, bullets and paragraph spacing.
            if (box.HasRichText)
            {
                for (int p = 0; p < box.RichParagraphs.Count; p++)
                {
                    PresentationTextParagraph paragraph = box.RichParagraphs[p];
                    if (paragraph == null)
                        continue;

                    for (int r = 0; r < paragraph.Runs.Count; r++)
                    {
                        PresentationTextRun run = paragraph.Runs[r];
                        if (run == null)
                            continue;

                        run.FontFamily = box.FontFamily;
                        run.FontSizePoints = box.FontSizePoints;
                        run.Bold = bold;
                        run.Italic = italic;
                    }
                }
            }

            MarkDirty();
            return true;
        }

        public void ReplaceDocument(PresentationDocument document, bool dirty)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            IsDirty = dirty;
        }

        public void MarkDirty()
        {
            IsDirty = true;
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("Save As is required for a new presentation.");

            PresentationPackageWriter.Save(Document, FilePath);
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

            PresentationPackageWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
            LastSavedUtc = DateTime.UtcNow;
        }

        private PresentationSlide GetSlide(int slideIndex)
        {
            if (slideIndex < 0 || slideIndex >= Document.Slides.Count)
                return null;

            return Document.Slides[slideIndex];
        }

        private PresentationTextBox GetTextBox(int slideIndex, int textBoxIndex)
        {
            PresentationSlide slide = GetSlide(slideIndex);

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
