using System;
using System.Collections.Generic;

namespace PptxViewer
{
    internal enum DocumentParagraphAlignment
    {
        Left,
        Center,
        Right,
        Justify
    }

    internal sealed class DocumentTextRun
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string ColorHex { get; set; }

        public DocumentTextRun()
        {
            Text = string.Empty;
            FontFamily = "Arial";
            FontSizePoints = 11f;
            ColorHex = "20242A";
        }

        public DocumentTextRun Clone()
        {
            DocumentTextRun copy = new DocumentTextRun();
            copy.Text = Text;
            copy.FontFamily = FontFamily;
            copy.FontSizePoints = FontSizePoints;
            copy.Bold = Bold;
            copy.Italic = Italic;
            copy.Underline = Underline;
            copy.ColorHex = ColorHex;
            return copy;
        }
    }

    internal sealed class DocumentParagraph
    {
        private readonly List<DocumentTextRun> runs =
            new List<DocumentTextRun>();

        public DocumentParagraphAlignment Alignment { get; set; }

        public IList<DocumentTextRun> Runs
        {
            get { return runs; }
        }

        public DocumentParagraph()
        {
            Alignment = DocumentParagraphAlignment.Left;
        }

        public DocumentTextRun AddRun(string text)
        {
            DocumentTextRun run = new DocumentTextRun();
            run.Text = text ?? string.Empty;
            runs.Add(run);
            return run;
        }

        public string PlainText
        {
            get
            {
                string result = string.Empty;
                for (int i = 0; i < runs.Count; i++)
                    result += runs[i] == null ? string.Empty : runs[i].Text;
                return result;
            }
        }

        public DocumentParagraph Clone()
        {
            DocumentParagraph copy = new DocumentParagraph();
            copy.Alignment = Alignment;

            for (int i = 0; i < runs.Count; i++)
            {
                if (runs[i] != null)
                    copy.runs.Add(runs[i].Clone());
            }

            return copy;
        }
    }

    internal sealed class TextDocument
    {
        private readonly List<DocumentParagraph> paragraphs =
            new List<DocumentParagraph>();

        public string Title { get; set; }

        public IList<DocumentParagraph> Paragraphs
        {
            get { return paragraphs; }
        }

        public TextDocument()
        {
            Title = "New Document";
        }

        public static TextDocument CreateNew(string title)
        {
            TextDocument document = new TextDocument();
            if (!string.IsNullOrEmpty(title))
                document.Title = title;
            return document;
        }

        public DocumentParagraph AddParagraph(string text)
        {
            DocumentParagraph paragraph = new DocumentParagraph();
            paragraph.AddRun(text ?? string.Empty);
            paragraphs.Add(paragraph);
            return paragraph;
        }

        public TextDocument Clone()
        {
            TextDocument copy = new TextDocument();
            copy.Title = Title;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i] != null)
                    copy.paragraphs.Add(paragraphs[i].Clone());
            }

            return copy;
        }
    }
}
