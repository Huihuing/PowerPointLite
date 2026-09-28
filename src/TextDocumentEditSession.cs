using System;
using System.IO;

namespace PptxViewer
{
    internal enum TextDocumentFileFormat
    {
        Docx,
        Hwpx,
        Odt
    }

    internal sealed class TextDocumentEditSession
    {
        public TextDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }
        public TextDocumentFileFormat Format { get; private set; }

        public string FormatDisplayName
        {
            get
            {
                if (Format == TextDocumentFileFormat.Hwpx)
                    return "HWPX";
                if (Format == TextDocumentFileFormat.Odt)
                    return "ODT";
                return "DOCX";
            }
        }

        public string DefaultExtension
        {
            get
            {
                if (Format == TextDocumentFileFormat.Hwpx)
                    return "hwpx";
                if (Format == TextDocumentFileFormat.Odt)
                    return "odt";
                return "docx";
            }
        }

        public string SaveDialogFilter
        {
            get
            {
                if (Format == TextDocumentFileFormat.Hwpx)
                    return "HWPX Document (*.hwpx)|*.hwpx";
                if (Format == TextDocumentFileFormat.Odt)
                    return "OpenDocument Text (*.odt)|*.odt";
                return "Word Open XML Document (*.docx)|*.docx";
            }
        }

        private TextDocumentEditSession(
            TextDocument document,
            TextDocumentFileFormat format)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            Format = format;
            FilePath = string.Empty;
            IsDirty = true;
        }

        public static TextDocumentEditSession CreateNew(string title)
        {
            return CreateNew(title, TextDocumentFileFormat.Docx);
        }

        public static TextDocumentEditSession CreateNewHwpx(string title)
        {
            return CreateNew(title, TextDocumentFileFormat.Hwpx);
        }

        public static TextDocumentEditSession CreateNewOdt(string title)
        {
            return CreateNew(title, TextDocumentFileFormat.Odt);
        }

        public static TextDocumentEditSession CreateNew(
            string title,
            TextDocumentFileFormat format)
        {
            TextDocument document = TextDocument.CreateNew(title);
            document.AddParagraph(string.Empty);
            return new TextDocumentEditSession(document, format);
        }

        public static TextDocumentEditSession Open(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("Document file was not found.", path);

            string extension = Path.GetExtension(path).ToLowerInvariant();

            if (extension == ".hwpx")
            {
                HwpxEditSafetyResult safety = HwpxEditSafety.Analyze(path);
                if (safety == null || !safety.CanEditSafely)
                {
                    throw new InvalidOperationException(
                        safety == null || string.IsNullOrEmpty(safety.Warning)
                            ? "This HWPX cannot yet be edited without risking unsupported-content loss."
                            : safety.Warning);
                }

                TextDocumentEditSession hwpx =
                    new TextDocumentEditSession(
                        HwpxReader.Read(path),
                        TextDocumentFileFormat.Hwpx);
                hwpx.FilePath = Path.GetFullPath(path);
                hwpx.IsDirty = false;
                return hwpx;
            }

            if (extension == ".odt")
            {
                OdtEditSafetyResult safety = OdtEditSafety.Analyze(path);
                if (safety == null || !safety.CanEditSafely)
                {
                    throw new InvalidOperationException(
                        safety == null || string.IsNullOrEmpty(safety.Warning)
                            ? "This ODT cannot yet be edited without risking unsupported-content loss."
                            : safety.Warning);
                }

                TextDocumentEditSession odt =
                    new TextDocumentEditSession(
                        OdtReader.Read(path),
                        TextDocumentFileFormat.Odt);
                odt.FilePath = Path.GetFullPath(path);
                odt.IsDirty = false;
                return odt;
            }

            if (extension != ".docx")
                throw new NotSupportedException("Only DOCX, HWPX and ODT are supported by the text document editor.");

            DocxEditSafetyResult docxSafety =
                DocxEditSafety.Analyze(path);

            if (docxSafety == null || !docxSafety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    docxSafety == null || string.IsNullOrEmpty(docxSafety.Warning)
                        ? "This DOCX cannot yet be edited without risking unsupported-content loss."
                        : docxSafety.Warning);
            }

            TextDocumentEditSession session =
                new TextDocumentEditSession(
                    DocxReader.Read(path),
                    TextDocumentFileFormat.Docx);
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

            SaveTo(FilePath);
            IsDirty = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");

            string expected = "." + DefaultExtension;
            if (!string.Equals(
                    Path.GetExtension(path),
                    expected,
                    StringComparison.OrdinalIgnoreCase))
            {
                path = Path.ChangeExtension(path, DefaultExtension);
            }

            SaveTo(path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
        }

        private void SaveTo(string path)
        {
            if (Format == TextDocumentFileFormat.Hwpx)
                HwpxWriter.Save(Document, path);
            else if (Format == TextDocumentFileFormat.Odt)
                OdtWriter.Save(Document, path);
            else
                DocxWriter.Save(Document, path);
        }
    }
}
