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

    internal sealed class PresentationSlide
    {
        private readonly List<PresentationTextBox> textBoxes =
            new List<PresentationTextBox>();

        public string Name { get; set; }

        public IList<PresentationTextBox> TextBoxes
        {
            get { return textBoxes; }
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

        public PresentationSlide Clone()
        {
            PresentationSlide copy = new PresentationSlide();
            copy.Name = Name;

            for (int i = 0; i < textBoxes.Count; i++)
                copy.textBoxes.Add(textBoxes[i].Clone());

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

            // Keep one slide so writer/editor state never becomes ambiguous.
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
