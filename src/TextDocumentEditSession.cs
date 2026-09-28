using System;
using System.IO;

namespace PptxViewer
{
    internal sealed class TextDocumentEditSession
    {
        public TextDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }

        private TextDocumentEditSession(TextDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            FilePath = string.Empty;
            IsDirty = true;
        }

        public static TextDocumentEditSession CreateNew(string title)
        {
            TextDocument document = TextDocument.CreateNew(title);
            document.AddParagraph(string.Empty);
            return new TextDocumentEditSession(document);
        }

        public static TextDocumentEditSession Open(string path)
        {
            TextDocument document = DocxReader.Read(path);
            TextDocumentEditSession session =
                new TextDocumentEditSession(document);
            session.FilePath = Path.GetFullPath(path);
            session.IsDirty = false;
            return session;
        }

        public DocumentParagraph AddParagraph(string text)
        {
            DocumentParagraph paragraph = Document.AddParagraph(text);
            IsDirty = true;
            return paragraph;
        }

        public bool RemoveParagraph(int index)
        {
            if (index < 0 || index >= Document.Paragraphs.Count)
                return false;

            if (Document.Paragraphs.Count <= 1)
                return false;

            Document.Paragraphs.RemoveAt(index);
            IsDirty = true;
            return true;
        }

        public bool MoveParagraph(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= Document.Paragraphs.Count ||
                toIndex < 0 || toIndex >= Document.Paragraphs.Count ||
                fromIndex == toIndex)
            {
                return false;
            }

            DocumentParagraph paragraph = Document.Paragraphs[fromIndex];
            Document.Paragraphs.RemoveAt(fromIndex);
            Document.Paragraphs.Insert(toIndex, paragraph);
            IsDirty = true;
            return true;
        }

        public void MarkDirty()
        {
            IsDirty = true;
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("Save As is required for a new document.");

            DocxWriter.Save(Document, FilePath);
            IsDirty = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");

            if (!string.Equals(
                    Path.GetExtension(path),
                    ".docx",
                    StringComparison.OrdinalIgnoreCase))
            {
                path += ".docx";
            }

            DocxWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
        }
    }
}
