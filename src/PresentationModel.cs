using System;
using System.Collections.Generic;

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

    internal sealed class PresentationTextBox
    {
        public string Name { get; set; }
        public string Text { get; set; }
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

        public PresentationTextBox()
        {
            Name = "Text Box";
            Text = string.Empty;
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

        public PresentationTextBox Clone()
        {
            PresentationTextBox copy = new PresentationTextBox();
            copy.Name = Name;
            copy.Text = Text;
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

        public PresentationSlide()
        {
            Name = "Slide";
        }

        public PresentationTextBox AddTextBox(string text)
        {
            PresentationTextBox box = new PresentationTextBox();
            box.Text = text ?? string.Empty;
            textBoxes.Add(box);
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
            return image;
        }

        public PresentationTable AddTable(int rows, int columns)
        {
            PresentationTable table = new PresentationTable(rows, columns);
            table.Name = "Table " + (tables.Count + 1).ToString();
            tables.Add(table);
            return table;
        }

        public bool RemoveTextBox(int index)
        {
            if (index < 0 || index >= textBoxes.Count)
                return false;
            textBoxes.RemoveAt(index);
            return true;
        }

        public bool RemoveShape(int index)
        {
            if (index < 0 || index >= shapes.Count)
                return false;
            shapes.RemoveAt(index);
            return true;
        }

        public bool RemoveImage(int index)
        {
            if (index < 0 || index >= images.Count)
                return false;
            images.RemoveAt(index);
            return true;
        }

        public bool RemoveTable(int index)
        {
            if (index < 0 || index >= tables.Count)
                return false;
            tables.RemoveAt(index);
            return true;
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
