using System;
using System.Collections.Generic;
using System.Text;

namespace PptxViewer
{
    internal enum PresentationTextAlignment
    {
        Left,
        Center,
        Right
    }

    internal enum PresentationShapeKind
    {
        Rectangle,
        RoundedRectangle,
        Ellipse,
        Triangle,
        Diamond
    }

    internal enum PresentationLayerKind
    {
        TextBox,
        Shape,
        Image,
        Table
    }

    internal sealed class PresentationLayerEntry
    {
        public PresentationLayerKind Kind { get; set; }
        public int Index { get; set; }

        public PresentationLayerEntry Clone()
        {
            PresentationLayerEntry copy =
                new PresentationLayerEntry();
            copy.Kind = Kind;
            copy.Index = Index;
            return copy;
        }
    }

    internal sealed class PresentationTextRun
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string ColorHex { get; set; }
        public int BaselinePercent { get; set; }

        public PresentationTextRun()
        {
            Text = string.Empty;
            FontFamily = "Arial";
            FontSizePoints = 20f;
            Bold = false;
            Italic = false;
            Underline = false;
            ColorHex = "20242A";
            BaselinePercent = 0;
        }

        public PresentationTextRun Clone()
        {
            PresentationTextRun copy = new PresentationTextRun();
            copy.Text = Text;
            copy.FontFamily = FontFamily;
            copy.FontSizePoints = FontSizePoints;
            copy.Bold = Bold;
            copy.Italic = Italic;
            copy.Underline = Underline;
            copy.ColorHex = ColorHex;
            copy.BaselinePercent = BaselinePercent;
            return copy;
        }
    }

    internal sealed class PresentationTextParagraph
    {
        private readonly List<PresentationTextRun> runs =
            new List<PresentationTextRun>();

        public PresentationTextAlignment Alignment { get; set; }
        public int Level { get; set; }
        public string BulletText { get; set; }
        public float SpaceBeforePoints { get; set; }
        public float SpaceAfterPoints { get; set; }

        public IList<PresentationTextRun> Runs
        {
            get { return runs; }
        }

        public PresentationTextParagraph()
        {
            Alignment = PresentationTextAlignment.Left;
            Level = 0;
            BulletText = string.Empty;
            SpaceBeforePoints = 0f;
            SpaceAfterPoints = 0f;
        }

        public PresentationTextParagraph Clone()
        {
            PresentationTextParagraph copy = new PresentationTextParagraph();
            copy.Alignment = Alignment;
            copy.Level = Level;
            copy.BulletText = BulletText;
            copy.SpaceBeforePoints = SpaceBeforePoints;
            copy.SpaceAfterPoints = SpaceAfterPoints;

            for (int i = 0; i < runs.Count; i++)
                copy.runs.Add(runs[i].Clone());

            return copy;
        }

        public string GetPlainText()
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < runs.Count; i++)
            {
                if (runs[i] != null && !string.IsNullOrEmpty(runs[i].Text))
                    builder.Append(runs[i].Text);
            }
            return builder.ToString();
        }
    }

    internal sealed class PresentationTextBox
    {
        private string text;
        private readonly List<PresentationTextParagraph> richParagraphs =
            new List<PresentationTextParagraph>();

        public string Name { get; set; }

        public string Text
        {
            get { return text ?? string.Empty; }
            set
            {
                text = value ?? string.Empty;
                richParagraphs.Clear();
            }
        }

        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string ColorHex { get; set; }
        public PresentationTextAlignment Alignment { get; set; }
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }

        public IList<PresentationTextParagraph> RichParagraphs
        {
            get { return richParagraphs; }
        }

        public bool HasRichText
        {
            get { return richParagraphs.Count > 0; }
        }

        public PresentationTextBox()
        {
            Name = "Text Box";
            text = string.Empty;
            FontFamily = "Arial";
            FontSizePoints = 20f;
            Bold = false;
            Italic = false;
            ColorHex = "20242A";
            Alignment = PresentationTextAlignment.Left;
            X = 914400;
            Y = 914400;
            Width = 9144000;
            Height = 914400;
        }

        public void SetRichParagraphs(
            IEnumerable<PresentationTextParagraph> paragraphs)
        {
            richParagraphs.Clear();

            if (paragraphs != null)
            {
                foreach (PresentationTextParagraph paragraph in paragraphs)
                {
                    if (paragraph != null)
                        richParagraphs.Add(paragraph.Clone());
                }
            }

            text = BuildPlainText(richParagraphs);
        }

        public void ClearRichText()
        {
            richParagraphs.Clear();
        }

        public void RefreshTextFromRichParagraphs()
        {
            text = BuildPlainText(
                richParagraphs);
        }

        private static string BuildPlainText(
            IList<PresentationTextParagraph> paragraphs)
        {
            if (paragraphs == null || paragraphs.Count == 0)
                return string.Empty;

            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (i > 0)
                    builder.AppendLine();

                PresentationTextParagraph paragraph = paragraphs[i];
                if (paragraph != null)
                    builder.Append(paragraph.GetPlainText());
            }
            return builder.ToString();
        }

        public PresentationTextBox Clone()
        {
            PresentationTextBox copy = new PresentationTextBox();
            copy.Name = Name;
            copy.text = text;
            copy.FontFamily = FontFamily;
            copy.FontSizePoints = FontSizePoints;
            copy.Bold = Bold;
            copy.Italic = Italic;
            copy.ColorHex = ColorHex;
            copy.Alignment = Alignment;
            copy.X = X;
            copy.Y = Y;
            copy.Width = Width;
            copy.Height = Height;

            for (int i = 0; i < richParagraphs.Count; i++)
                copy.richParagraphs.Add(richParagraphs[i].Clone());

            return copy;
        }
    }

    internal sealed class PresentationShape
    {
        public string Name { get; set; }
        public PresentationShapeKind Kind { get; set; }
        public string FillColorHex { get; set; }
        public string LineColorHex { get; set; }
        public float LineWidthPoints { get; set; }
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }

        public PresentationShape()
        {
            Name = "Shape";
            Kind = PresentationShapeKind.Rectangle;
            FillColorHex = "5B8CFF";
            LineColorHex = "356AE6";
            LineWidthPoints = 1.25f;
            X = 1371600;
            Y = 1371600;
            Width = 2743200;
            Height = 1828800;
        }

        public PresentationShape Clone()
        {
            PresentationShape copy = new PresentationShape();
            copy.Name = Name;
            copy.Kind = Kind;
            copy.FillColorHex = FillColorHex;
            copy.LineColorHex = LineColorHex;
            copy.LineWidthPoints = LineWidthPoints;
            copy.X = X;
            copy.Y = Y;
            copy.Width = Width;
            copy.Height = Height;
            return copy;
        }
    }

    internal sealed class PresentationImage
    {
        public string Name { get; set; }
        public string Extension { get; set; }
        public string ContentType { get; set; }
        public byte[] Data { get; set; }
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }

        public PresentationImage()
        {
            Name = "Image";
            Extension = "png";
            ContentType = "image/png";
            Data = new byte[0];
            X = 1371600;
            Y = 1371600;
            Width = 3657600;
            Height = 2743200;
        }

        public PresentationImage Clone()
        {
            PresentationImage copy = new PresentationImage();
            copy.Name = Name;
            copy.Extension = Extension;
            copy.ContentType = ContentType;
            copy.Data = Data == null ? new byte[0] : (byte[])Data.Clone();
            copy.X = X;
            copy.Y = Y;
            copy.Width = Width;
            copy.Height = Height;
            return copy;
        }
    }

    internal sealed class PresentationTableCell
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool Bold { get; set; }
        public string TextColorHex { get; set; }
        public string FillColorHex { get; set; }
        public PresentationTextAlignment Alignment { get; set; }

        public PresentationTableCell()
        {
            Text = string.Empty;
            FontFamily = "Arial";
            FontSizePoints = 16f;
            Bold = false;
            TextColorHex = "20242A";
            FillColorHex = "FFFFFF";
            Alignment = PresentationTextAlignment.Left;
        }

        public PresentationTableCell Clone()
        {
            PresentationTableCell copy = new PresentationTableCell();
            copy.Text = Text;
            copy.FontFamily = FontFamily;
            copy.FontSizePoints = FontSizePoints;
            copy.Bold = Bold;
            copy.TextColorHex = TextColorHex;
            copy.FillColorHex = FillColorHex;
            copy.Alignment = Alignment;
            return copy;
        }
    }

    internal sealed class PresentationTable
    {
        private readonly List<PresentationTableCell> cells =
            new List<PresentationTableCell>();

        public string Name { get; set; }
        public int Rows { get; private set; }
        public int Columns { get; private set; }
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }

        public PresentationTable(int rows, int columns)
        {
            Rows = Math.Max(1, Math.Min(50, rows));
            Columns = Math.Max(1, Math.Min(50, columns));
            Name = "Table";
            X = 1371600;
            Y = 1828800;
            Width = 9448800;
            Height = 2743200;

            int count = Rows * Columns;
            for (int i = 0; i < count; i++)
                cells.Add(new PresentationTableCell());
        }

        public PresentationTableCell GetCell(int row, int column)
        {
            if (row < 0 || row >= Rows || column < 0 || column >= Columns)
                return null;

            return cells[row * Columns + column];
        }

        public PresentationTable Clone()
        {
            PresentationTable copy = new PresentationTable(Rows, Columns);
            copy.Name = Name;
            copy.X = X;
            copy.Y = Y;
            copy.Width = Width;
            copy.Height = Height;

            for (int i = 0; i < cells.Count; i++)
                copy.cells[i] = cells[i].Clone();

            return copy;
        }
    }

    internal sealed class PresentationSlide
    {
        private readonly List<PresentationTextBox> textBoxes =
            new List<PresentationTextBox>();
        private readonly List<PresentationShape> shapes =
            new List<PresentationShape>();
        private readonly List<PresentationImage> images =
            new List<PresentationImage>();
        private readonly List<PresentationTable> tables =
            new List<PresentationTable>();
        private readonly List<PresentationLayerEntry> objectOrder =
            new List<PresentationLayerEntry>();

        public string Name { get; set; }

        public IList<PresentationTextBox> TextBoxes
        {
            get { return textBoxes; }
        }

        public IList<PresentationShape> Shapes
        {
            get { return shapes; }
        }

        public IList<PresentationImage> Images
        {
            get { return images; }
        }

        public IList<PresentationTable> Tables
        {
            get { return tables; }
        }

        public IList<PresentationLayerEntry> ObjectOrder
        {
            get { return objectOrder; }
        }

        public PresentationSlide()
        {
            Name = "Slide";
        }

        public PresentationTextBox AddTextBox(string text)
        {
            PresentationTextBox box = new PresentationTextBox();
            box.Text = text ?? string.Empty;
            textBoxes.Add(box);
            RegisterObjectOrder(
                PresentationLayerKind.TextBox,
                textBoxes.Count - 1);
            return box;
        }

        public PresentationTextBox AddTitle(string text)
        {
            PresentationTextBox box = AddTextBox(text);
            box.Name = "Title";
            box.X = 914400;
            box.Y = 1828800;
            box.Width = 10363200;
            box.Height = 1371600;
            box.FontSizePoints = 28f;
            box.Bold = true;
            box.Alignment = PresentationTextAlignment.Center;
            return box;
        }

        public PresentationShape AddShape(PresentationShapeKind kind)
        {
            PresentationShape shape = new PresentationShape();
            shape.Kind = kind;
            shape.Name = kind.ToString();
            shapes.Add(shape);
            RegisterObjectOrder(
                PresentationLayerKind.Shape,
                shapes.Count - 1);
            return shape;
        }

        public PresentationImage AddImage(
            byte[] data,
            string extension,
            string contentType)
        {
            PresentationImage image = new PresentationImage();
            image.Data = data == null ? new byte[0] : (byte[])data.Clone();
            image.Extension = string.IsNullOrEmpty(extension)
                ? "png"
                : extension.Trim().TrimStart('.').ToLowerInvariant();
            image.ContentType = string.IsNullOrEmpty(contentType)
                ? "image/png"
                : contentType;
            images.Add(image);
            RegisterObjectOrder(
                PresentationLayerKind.Image,
                images.Count - 1);
            return image;
        }

        public PresentationTable AddTable(int rows, int columns)
        {
            PresentationTable table = new PresentationTable(rows, columns);
            table.Name = "Table " + (tables.Count + 1).ToString();
            tables.Add(table);
            RegisterObjectOrder(
                PresentationLayerKind.Table,
                tables.Count - 1);
            return table;
        }

        public bool RemoveTextBox(int index)
        {
            if (index < 0 || index >= textBoxes.Count)
                return false;
            textBoxes.RemoveAt(index);
            RemoveObjectOrder(
                PresentationLayerKind.TextBox,
                index);
            return true;
        }

        public bool RemoveShape(int index)
        {
            if (index < 0 || index >= shapes.Count)
                return false;
            shapes.RemoveAt(index);
            RemoveObjectOrder(
                PresentationLayerKind.Shape,
                index);
            return true;
        }

        public bool RemoveImage(int index)
        {
            if (index < 0 || index >= images.Count)
                return false;
            images.RemoveAt(index);
            RemoveObjectOrder(
                PresentationLayerKind.Image,
                index);
            return true;
        }

        public bool RemoveTable(int index)
        {
            if (index < 0 || index >= tables.Count)
                return false;
            tables.RemoveAt(index);
            RemoveObjectOrder(
                PresentationLayerKind.Table,
                index);
            return true;
        }


        public void RegisterObjectOrder(
            PresentationLayerKind kind,
            int index)
        {
            if (!IsValidObjectIndex(
                    kind,
                    index))
            {
                return;
            }

            for (int i = 0;
                 i < objectOrder.Count;
                 i++)
            {
                PresentationLayerEntry existing =
                    objectOrder[i];

                if (existing != null &&
                    existing.Kind == kind &&
                    existing.Index == index)
                {
                    return;
                }
            }

            PresentationLayerEntry entry =
                new PresentationLayerEntry();
            entry.Kind = kind;
            entry.Index = index;
            objectOrder.Add(entry);
        }

        public void SynchronizeObjectOrder()
        {
            List<PresentationLayerEntry> normalized =
                new List<PresentationLayerEntry>();

            bool[] textSeen =
                new bool[textBoxes.Count];
            bool[] shapeSeen =
                new bool[shapes.Count];
            bool[] imageSeen =
                new bool[images.Count];
            bool[] tableSeen =
                new bool[tables.Count];

            for (int i = 0;
                 i < objectOrder.Count;
                 i++)
            {
                PresentationLayerEntry entry =
                    objectOrder[i];

                if (entry == null ||
                    !IsValidObjectIndex(
                        entry.Kind,
                        entry.Index))
                {
                    continue;
                }

                bool alreadySeen = false;

                if (entry.Kind ==
                    PresentationLayerKind.TextBox)
                {
                    alreadySeen =
                        textSeen[entry.Index];
                    textSeen[entry.Index] = true;
                }
                else if (entry.Kind ==
                         PresentationLayerKind.Shape)
                {
                    alreadySeen =
                        shapeSeen[entry.Index];
                    shapeSeen[entry.Index] = true;
                }
                else if (entry.Kind ==
                         PresentationLayerKind.Image)
                {
                    alreadySeen =
                        imageSeen[entry.Index];
                    imageSeen[entry.Index] = true;
                }
                else if (entry.Kind ==
                         PresentationLayerKind.Table)
                {
                    alreadySeen =
                        tableSeen[entry.Index];
                    tableSeen[entry.Index] = true;
                }

                if (!alreadySeen)
                    normalized.Add(
                        entry.Clone());
            }

            for (int i = 0;
                 i < shapes.Count;
                 i++)
            {
                if (!shapeSeen[i])
                    AddNormalizedLayerEntry(
                        normalized,
                        PresentationLayerKind.Shape,
                        i);
            }

            for (int i = 0;
                 i < images.Count;
                 i++)
            {
                if (!imageSeen[i])
                    AddNormalizedLayerEntry(
                        normalized,
                        PresentationLayerKind.Image,
                        i);
            }

            for (int i = 0;
                 i < textBoxes.Count;
                 i++)
            {
                if (!textSeen[i])
                    AddNormalizedLayerEntry(
                        normalized,
                        PresentationLayerKind.TextBox,
                        i);
            }

            // Tables historically rendered after the other editable
            // object kinds. Keep that fallback for legacy model instances
            // that predate the shared layer-order list.
            for (int i = 0;
                 i < tables.Count;
                 i++)
            {
                if (!tableSeen[i])
                    AddNormalizedLayerEntry(
                        normalized,
                        PresentationLayerKind.Table,
                        i);
            }

            objectOrder.Clear();

            for (int i = 0;
                 i < normalized.Count;
                 i++)
            {
                objectOrder.Add(
                    normalized[i]);
            }
        }

        public bool MoveObjectToFront(
            PresentationLayerKind kind,
            int index)
        {
            SynchronizeObjectOrder();

            int position =
                FindObjectOrderPosition(
                    kind,
                    index);

            if (position < 0 ||
                position ==
                    objectOrder.Count - 1)
            {
                return false;
            }

            PresentationLayerEntry entry =
                objectOrder[position];

            objectOrder.RemoveAt(position);
            objectOrder.Add(entry);
            return true;
        }

        public bool MoveObjectToBack(
            PresentationLayerKind kind,
            int index)
        {
            SynchronizeObjectOrder();

            int position =
                FindObjectOrderPosition(
                    kind,
                    index);

            if (position <= 0)
                return false;

            PresentationLayerEntry entry =
                objectOrder[position];

            objectOrder.RemoveAt(position);
            objectOrder.Insert(
                0,
                entry);
            return true;
        }

        public bool CanMoveObjectToFront(
            PresentationLayerKind kind,
            int index)
        {
            SynchronizeObjectOrder();

            int position =
                FindObjectOrderPosition(
                    kind,
                    index);

            return position >= 0 &&
                position <
                    objectOrder.Count - 1;
        }

        public bool CanMoveObjectToBack(
            PresentationLayerKind kind,
            int index)
        {
            SynchronizeObjectOrder();

            return FindObjectOrderPosition(
                kind,
                index) > 0;
        }

        private static void AddNormalizedLayerEntry(
            List<PresentationLayerEntry> target,
            PresentationLayerKind kind,
            int index)
        {
            PresentationLayerEntry entry =
                new PresentationLayerEntry();
            entry.Kind = kind;
            entry.Index = index;
            target.Add(entry);
        }

        private int FindObjectOrderPosition(
            PresentationLayerKind kind,
            int index)
        {
            for (int i = 0;
                 i < objectOrder.Count;
                 i++)
            {
                PresentationLayerEntry entry =
                    objectOrder[i];

                if (entry != null &&
                    entry.Kind == kind &&
                    entry.Index == index)
                {
                    return i;
                }
            }

            return -1;
        }

        private bool IsValidObjectIndex(
            PresentationLayerKind kind,
            int index)
        {
            if (index < 0)
                return false;

            if (kind ==
                PresentationLayerKind.TextBox)
            {
                return index <
                    textBoxes.Count;
            }

            if (kind ==
                PresentationLayerKind.Shape)
            {
                return index <
                    shapes.Count;
            }

            if (kind ==
                PresentationLayerKind.Image)
            {
                return index <
                    images.Count;
            }

            if (kind ==
                PresentationLayerKind.Table)
            {
                return index <
                    tables.Count;
            }

            return false;
        }

        private void RemoveObjectOrder(
            PresentationLayerKind kind,
            int removedIndex)
        {
            for (int i =
                     objectOrder.Count - 1;
                 i >= 0;
                 i--)
            {
                PresentationLayerEntry entry =
                    objectOrder[i];

                if (entry == null)
                {
                    objectOrder.RemoveAt(i);
                    continue;
                }

                if (entry.Kind != kind)
                    continue;

                if (entry.Index ==
                    removedIndex)
                {
                    objectOrder.RemoveAt(i);
                }
                else if (entry.Index >
                         removedIndex)
                {
                    entry.Index--;
                }
            }
        }

        public PresentationSlide Clone()
        {
            PresentationSlide copy = new PresentationSlide();
            copy.Name = Name;

            for (int i = 0; i < textBoxes.Count; i++)
                copy.textBoxes.Add(textBoxes[i].Clone());

            for (int i = 0; i < shapes.Count; i++)
                copy.shapes.Add(shapes[i].Clone());

            for (int i = 0; i < images.Count; i++)
                copy.images.Add(images[i].Clone());

            for (int i = 0; i < tables.Count; i++)
                copy.tables.Add(tables[i].Clone());

            SynchronizeObjectOrder();

            for (int i = 0;
                 i < objectOrder.Count;
                 i++)
            {
                copy.objectOrder.Add(
                    objectOrder[i].Clone());
            }

            return copy;
        }
    }

    internal sealed class PresentationDocument
    {
        public const long DefaultWidthEmu = 12192000;
        public const long DefaultHeightEmu = 6858000;

        private readonly List<PresentationSlide> slides =
            new List<PresentationSlide>();

        public string Title { get; set; }
        public long WidthEmu { get; set; }
        public long HeightEmu { get; set; }

        public IList<PresentationSlide> Slides
        {
            get { return slides; }
        }

        public PresentationDocument()
        {
            Title = "New Presentation";
            WidthEmu = DefaultWidthEmu;
            HeightEmu = DefaultHeightEmu;
        }

        public static PresentationDocument CreateNew(string title)
        {
            PresentationDocument document = new PresentationDocument();

            if (!string.IsNullOrEmpty(title))
                document.Title = title;

            PresentationSlide first = document.AddSlide();
            first.Name = "Title Slide";
            first.AddTitle(document.Title);
            return document;
        }

        public PresentationSlide AddSlide()
        {
            PresentationSlide slide = new PresentationSlide();
            slide.Name = "Slide " + (slides.Count + 1).ToString();
            slides.Add(slide);
            return slide;
        }

        public PresentationSlide AddSlide(string title)
        {
            PresentationSlide slide = AddSlide();

            if (!string.IsNullOrEmpty(title))
                slide.AddTitle(title);

            return slide;
        }

        public bool RemoveSlide(int index)
        {
            if (index < 0 || index >= slides.Count)
                return false;

            if (slides.Count <= 1)
                return false;

            slides.RemoveAt(index);
            return true;
        }

        public bool MoveSlide(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= slides.Count ||
                toIndex < 0 || toIndex >= slides.Count ||
                fromIndex == toIndex)
            {
                return false;
            }

            PresentationSlide slide = slides[fromIndex];
            slides.RemoveAt(fromIndex);
            slides.Insert(toIndex, slide);
            return true;
        }

        public PresentationDocument Clone()
        {
            PresentationDocument copy = new PresentationDocument();
            copy.Title = Title;
            copy.WidthEmu = WidthEmu;
            copy.HeightEmu = HeightEmu;

            for (int i = 0; i < slides.Count; i++)
                copy.slides.Add(slides[i].Clone());

            return copy;
        }
    }
}
