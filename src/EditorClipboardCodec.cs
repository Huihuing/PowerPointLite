using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PptxViewer
{
    internal sealed class EditorClipboardObject
    {
        public EditorObjectKind Kind { get; set; }
        public PresentationTextBox TextBox { get; set; }
        public PresentationShape Shape { get; set; }
        public PresentationImage Image { get; set; }

        public EditorClipboardObject Clone()
        {
            EditorClipboardObject copy =
                new EditorClipboardObject();
            copy.Kind = Kind;
            copy.TextBox =
                TextBox == null
                    ? null
                    : TextBox.Clone();
            copy.Shape =
                Shape == null
                    ? null
                    : Shape.Clone();
            copy.Image =
                Image == null
                    ? null
                    : Image.Clone();
            return copy;
        }

        public bool TryGetGeometry(
            out long x,
            out long y,
            out long width,
            out long height)
        {
            x = 0L;
            y = 0L;
            width = 0L;
            height = 0L;

            if (Kind == EditorObjectKind.TextBox &&
                TextBox != null)
            {
                x = TextBox.X;
                y = TextBox.Y;
                width = TextBox.Width;
                height = TextBox.Height;
                return true;
            }

            if (Kind == EditorObjectKind.Shape &&
                Shape != null)
            {
                x = Shape.X;
                y = Shape.Y;
                width = Shape.Width;
                height = Shape.Height;
                return true;
            }

            if (Kind == EditorObjectKind.Image &&
                Image != null)
            {
                x = Image.X;
                y = Image.Y;
                width = Image.Width;
                height = Image.Height;
                return true;
            }

            return false;
        }

        public void Offset(
            long dx,
            long dy)
        {
            if (Kind == EditorObjectKind.TextBox &&
                TextBox != null)
            {
                TextBox.X += dx;
                TextBox.Y += dy;
            }
            else if (Kind == EditorObjectKind.Shape &&
                     Shape != null)
            {
                Shape.X += dx;
                Shape.Y += dy;
            }
            else if (Kind == EditorObjectKind.Image &&
                     Image != null)
            {
                Image.X += dx;
                Image.Y += dy;
            }
        }
    }

    internal sealed class EditorClipboardPackage
    {
        private readonly List<EditorClipboardObject> objects =
            new List<EditorClipboardObject>();

        public IList<EditorClipboardObject> Objects
        {
            get { return objects; }
        }

        public bool HasObjects
        {
            get { return objects.Count > 0; }
        }

        public EditorClipboardPackage Clone()
        {
            EditorClipboardPackage copy =
                new EditorClipboardPackage();

            for (int i = 0;
                 i < objects.Count;
                 i++)
            {
                if (objects[i] != null)
                    copy.objects.Add(
                        objects[i].Clone());
            }

            return copy;
        }
    }

    internal static class EditorClipboardCodec
    {
        public const string ClipboardFormat =
            "PowerPointLite.ObjectSelection.v1";

        private const string Magic =
            "PPLT_OBJECTS";
        private const int Version = 1;
        private const int MaxObjects = 512;
        private const int MaxParagraphs = 4096;
        private const int MaxRuns = 16384;
        private const int MaxImageBytes =
            48 * 1024 * 1024;
        private const int MaxEncodedLength =
            96 * 1024 * 1024;

        public static EditorClipboardPackage CreatePackage(
            PresentationSlide slide,
            IList<EditorSelectionEntry> selection)
        {
            EditorClipboardPackage package =
                new EditorClipboardPackage();

            if (slide == null ||
                selection == null ||
                selection.Count == 0)
            {
                return package;
            }

            slide.SynchronizeObjectOrder();

            for (int orderIndex = 0;
                 orderIndex < slide.ObjectOrder.Count;
                 orderIndex++)
            {
                PresentationLayerEntry layer =
                    slide.ObjectOrder[orderIndex];

                if (layer == null)
                    continue;

                EditorObjectKind kind =
                    LayerToEditorKind(
                        layer.Kind);

                if (kind == EditorObjectKind.None ||
                    !ContainsSelection(
                        selection,
                        kind,
                        layer.Index))
                {
                    continue;
                }

                EditorClipboardObject item =
                    CloneObject(
                        slide,
                        kind,
                        layer.Index);

                if (item != null)
                    package.Objects.Add(item);
            }

            return package;
        }

        public static string Encode(
            EditorClipboardPackage package)
        {
            if (package == null ||
                !package.HasObjects)
            {
                return string.Empty;
            }

            using (MemoryStream stream =
                new MemoryStream())
            using (BinaryWriter writer =
                new BinaryWriter(
                    stream,
                    Encoding.UTF8))
            {
                writer.Write(Magic);
                writer.Write(Version);
                writer.Write(package.Objects.Count);

                for (int i = 0;
                     i < package.Objects.Count;
                     i++)
                {
                    WriteObject(
                        writer,
                        package.Objects[i]);
                }

                writer.Flush();
                return Convert.ToBase64String(
                    stream.ToArray());
            }
        }

        public static bool TryDecode(
            string encoded,
            out EditorClipboardPackage package)
        {
            package = null;

            if (string.IsNullOrEmpty(encoded) ||
                encoded.Length > MaxEncodedLength)
            {
                return false;
            }

            try
            {
                byte[] bytes =
                    Convert.FromBase64String(
                        encoded);

                using (MemoryStream stream =
                    new MemoryStream(
                        bytes,
                        false))
                using (BinaryReader reader =
                    new BinaryReader(
                        stream,
                        Encoding.UTF8))
                {
                    if (!string.Equals(
                            reader.ReadString(),
                            Magic,
                            StringComparison.Ordinal))
                    {
                        return false;
                    }

                    int version =
                        reader.ReadInt32();

                    if (version != Version)
                        return false;

                    int count =
                        ReadCount(
                            reader,
                            MaxObjects);

                    if (count <= 0)
                        return false;

                    EditorClipboardPackage result =
                        new EditorClipboardPackage();

                    for (int i = 0;
                         i < count;
                         i++)
                    {
                        EditorClipboardObject item =
                            ReadObject(reader);

                        if (item == null)
                            return false;

                        result.Objects.Add(item);
                    }

                    package = result;
                    return package.HasObjects;
                }
            }
            catch
            {
                package = null;
                return false;
            }
        }

        private static EditorClipboardObject CloneObject(
            PresentationSlide slide,
            EditorObjectKind kind,
            int index)
        {
            if (kind == EditorObjectKind.TextBox &&
                index >= 0 &&
                index < slide.TextBoxes.Count)
            {
                EditorClipboardObject item =
                    new EditorClipboardObject();
                item.Kind = kind;
                item.TextBox =
                    slide.TextBoxes[index].Clone();
                return item;
            }

            if (kind == EditorObjectKind.Shape &&
                index >= 0 &&
                index < slide.Shapes.Count)
            {
                EditorClipboardObject item =
                    new EditorClipboardObject();
                item.Kind = kind;
                item.Shape =
                    slide.Shapes[index].Clone();
                return item;
            }

            if (kind == EditorObjectKind.Image &&
                index >= 0 &&
                index < slide.Images.Count)
            {
                EditorClipboardObject item =
                    new EditorClipboardObject();
                item.Kind = kind;
                item.Image =
                    slide.Images[index].Clone();
                return item;
            }

            return null;
        }

        private static bool ContainsSelection(
            IList<EditorSelectionEntry> selection,
            EditorObjectKind kind,
            int index)
        {
            for (int i = 0;
                 i < selection.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selection[i];

                if (entry != null &&
                    entry.Kind == kind &&
                    entry.Index == index)
                {
                    return true;
                }
            }

            return false;
        }

        private static EditorObjectKind LayerToEditorKind(
            PresentationLayerKind kind)
        {
            if (kind == PresentationLayerKind.TextBox)
                return EditorObjectKind.TextBox;

            if (kind == PresentationLayerKind.Shape)
                return EditorObjectKind.Shape;

            if (kind == PresentationLayerKind.Image)
                return EditorObjectKind.Image;

            return EditorObjectKind.None;
        }

        private static void WriteObject(
            BinaryWriter writer,
            EditorClipboardObject item)
        {
            if (item == null)
                throw new InvalidDataException(
                    "Clipboard object is missing.");

            writer.Write((int)item.Kind);

            if (item.Kind == EditorObjectKind.TextBox)
                WriteTextBox(writer, item.TextBox);
            else if (item.Kind == EditorObjectKind.Shape)
                WriteShape(writer, item.Shape);
            else if (item.Kind == EditorObjectKind.Image)
                WriteImage(writer, item.Image);
            else
                throw new InvalidDataException(
                    "Unsupported clipboard object kind.");
        }

        private static EditorClipboardObject ReadObject(
            BinaryReader reader)
        {
            EditorObjectKind kind =
                (EditorObjectKind)reader.ReadInt32();
            EditorClipboardObject item =
                new EditorClipboardObject();
            item.Kind = kind;

            if (kind == EditorObjectKind.TextBox)
                item.TextBox = ReadTextBox(reader);
            else if (kind == EditorObjectKind.Shape)
                item.Shape = ReadShape(reader);
            else if (kind == EditorObjectKind.Image)
                item.Image = ReadImage(reader);
            else
                return null;

            return item;
        }

        private static void WriteTextBox(
            BinaryWriter writer,
            PresentationTextBox box)
        {
            if (box == null)
                throw new InvalidDataException(
                    "Clipboard text box is missing.");

            WriteString(writer, box.Name);
            WriteString(writer, box.Text);
            WriteString(writer, box.FontFamily);
            writer.Write(box.FontSizePoints);
            writer.Write(box.Bold);
            writer.Write(box.Italic);
            WriteString(writer, box.ColorHex);
            writer.Write((int)box.Alignment);
            writer.Write(box.X);
            writer.Write(box.Y);
            writer.Write(box.Width);
            writer.Write(box.Height);
            writer.Write(box.RichParagraphs.Count);

            for (int i = 0;
                 i < box.RichParagraphs.Count;
                 i++)
            {
                PresentationTextParagraph paragraph =
                    box.RichParagraphs[i];

                writer.Write((int)paragraph.Alignment);
                writer.Write(paragraph.Level);
                WriteString(
                    writer,
                    paragraph.BulletText);
                writer.Write(
                    paragraph.SpaceBeforePoints);
                writer.Write(
                    paragraph.SpaceAfterPoints);
                writer.Write(paragraph.Runs.Count);

                for (int r = 0;
                     r < paragraph.Runs.Count;
                     r++)
                {
                    PresentationTextRun run =
                        paragraph.Runs[r];
                    WriteString(writer, run.Text);
                    WriteString(
                        writer,
                        run.FontFamily);
                    writer.Write(
                        run.FontSizePoints);
                    writer.Write(run.Bold);
                    writer.Write(run.Italic);
                    writer.Write(run.Underline);
                    WriteString(
                        writer,
                        run.ColorHex);
                    writer.Write(
                        run.BaselinePercent);
                }
            }
        }

        private static PresentationTextBox ReadTextBox(
            BinaryReader reader)
        {
            PresentationTextBox box =
                new PresentationTextBox();
            box.Name = reader.ReadString();
            string plainText =
                reader.ReadString();
            box.Text = plainText;
            box.FontFamily =
                reader.ReadString();
            box.FontSizePoints =
                reader.ReadSingle();
            box.Bold = reader.ReadBoolean();
            box.Italic = reader.ReadBoolean();
            box.ColorHex =
                reader.ReadString();
            box.Alignment =
                (PresentationTextAlignment)
                reader.ReadInt32();
            box.X = reader.ReadInt64();
            box.Y = reader.ReadInt64();
            box.Width = reader.ReadInt64();
            box.Height = reader.ReadInt64();

            int paragraphCount =
                ReadCount(
                    reader,
                    MaxParagraphs);

            if (paragraphCount > 0)
            {
                List<PresentationTextParagraph> paragraphs =
                    new List<PresentationTextParagraph>();

                for (int i = 0;
                     i < paragraphCount;
                     i++)
                {
                    PresentationTextParagraph paragraph =
                        new PresentationTextParagraph();
                    paragraph.Alignment =
                        (PresentationTextAlignment)
                        reader.ReadInt32();
                    paragraph.Level =
                        reader.ReadInt32();
                    paragraph.BulletText =
                        reader.ReadString();
                    paragraph.SpaceBeforePoints =
                        reader.ReadSingle();
                    paragraph.SpaceAfterPoints =
                        reader.ReadSingle();

                    int runCount =
                        ReadCount(
                            reader,
                            MaxRuns);

                    for (int r = 0;
                         r < runCount;
                         r++)
                    {
                        PresentationTextRun run =
                            new PresentationTextRun();
                        run.Text =
                            reader.ReadString();
                        run.FontFamily =
                            reader.ReadString();
                        run.FontSizePoints =
                            reader.ReadSingle();
                        run.Bold =
                            reader.ReadBoolean();
                        run.Italic =
                            reader.ReadBoolean();
                        run.Underline =
                            reader.ReadBoolean();
                        run.ColorHex =
                            reader.ReadString();
                        run.BaselinePercent =
                            reader.ReadInt32();
                        paragraph.Runs.Add(run);
                    }

                    paragraphs.Add(paragraph);
                }

                box.SetRichParagraphs(
                    paragraphs);
            }

            return box;
        }

        private static void WriteShape(
            BinaryWriter writer,
            PresentationShape shape)
        {
            if (shape == null)
                throw new InvalidDataException(
                    "Clipboard shape is missing.");

            WriteString(writer, shape.Name);
            writer.Write((int)shape.Kind);
            WriteString(
                writer,
                shape.FillColorHex);
            WriteString(
                writer,
                shape.LineColorHex);
            writer.Write(
                shape.LineWidthPoints);
            writer.Write(shape.X);
            writer.Write(shape.Y);
            writer.Write(shape.Width);
            writer.Write(shape.Height);
        }

        private static PresentationShape ReadShape(
            BinaryReader reader)
        {
            PresentationShape shape =
                new PresentationShape();
            shape.Name = reader.ReadString();
            shape.Kind =
                (PresentationShapeKind)
                reader.ReadInt32();
            shape.FillColorHex =
                reader.ReadString();
            shape.LineColorHex =
                reader.ReadString();
            shape.LineWidthPoints =
                reader.ReadSingle();
            shape.X = reader.ReadInt64();
            shape.Y = reader.ReadInt64();
            shape.Width = reader.ReadInt64();
            shape.Height = reader.ReadInt64();
            return shape;
        }

        private static void WriteImage(
            BinaryWriter writer,
            PresentationImage image)
        {
            if (image == null)
                throw new InvalidDataException(
                    "Clipboard image is missing.");

            WriteString(writer, image.Name);
            WriteString(
                writer,
                image.Extension);
            WriteString(
                writer,
                image.ContentType);
            writer.Write(image.X);
            writer.Write(image.Y);
            writer.Write(image.Width);
            writer.Write(image.Height);

            byte[] data =
                image.Data ??
                new byte[0];

            if (data.Length > MaxImageBytes)
                throw new InvalidDataException(
                    "Clipboard image is too large.");

            writer.Write(data.Length);
            writer.Write(data);
        }

        private static PresentationImage ReadImage(
            BinaryReader reader)
        {
            PresentationImage image =
                new PresentationImage();
            image.Name = reader.ReadString();
            image.Extension =
                reader.ReadString();
            image.ContentType =
                reader.ReadString();
            image.X = reader.ReadInt64();
            image.Y = reader.ReadInt64();
            image.Width = reader.ReadInt64();
            image.Height = reader.ReadInt64();

            int length =
                reader.ReadInt32();

            if (length < 0 ||
                length > MaxImageBytes)
            {
                throw new InvalidDataException(
                    "Clipboard image length is invalid.");
            }

            byte[] data =
                reader.ReadBytes(length);

            if (data.Length != length)
                throw new EndOfStreamException();

            image.Data = data;
            return image;
        }

        private static void WriteString(
            BinaryWriter writer,
            string value)
        {
            writer.Write(
                value ??
                string.Empty);
        }

        private static int ReadCount(
            BinaryReader reader,
            int maximum)
        {
            int count =
                reader.ReadInt32();

            if (count < 0 ||
                count > maximum)
            {
                throw new InvalidDataException(
                    "Clipboard collection count is invalid.");
            }

            return count;
        }
    }

    internal static class EditorClipboardCodecDiagnostics
    {
        public static void Validate()
        {
            PresentationDocument document =
                PresentationDocument.CreateNew(
                    "Clipboard Self Test");
            PresentationSlide slide =
                document.Slides[0];

            PresentationShape shape =
                slide.AddShape(
                    PresentationShapeKind.Diamond);
            shape.Name = "Clipboard diamond";
            shape.X = 1828800;
            shape.Y = 2057400;
            shape.FillColorHex = "A05BC7";

            PresentationImage image =
                slide.AddImage(
                    new byte[]
                    {
                        1, 2, 3, 4, 5, 6
                    },
                    "png",
                    "image/png");
            image.Name = "Clipboard image";
            image.X = 4114800;
            image.Y = 2286000;

            PresentationTextBox text =
                slide.AddTextBox(
                    "Clipboard text");
            text.Name = "Clipboard text box";
            text.X = 6400800;
            text.Y = 2514600;
            text.Bold = true;
            text.ColorHex = "2255AA";

            List<EditorSelectionEntry> selection =
                new List<EditorSelectionEntry>();
            selection.Add(
                Selection(
                    EditorObjectKind.Shape,
                    slide.Shapes.Count - 1));
            selection.Add(
                Selection(
                    EditorObjectKind.Image,
                    slide.Images.Count - 1));
            selection.Add(
                Selection(
                    EditorObjectKind.TextBox,
                    slide.TextBoxes.Count - 1));

            EditorClipboardPackage package =
                EditorClipboardCodec.CreatePackage(
                    slide,
                    selection);

            if (package == null ||
                package.Objects.Count != 3)
            {
                throw new InvalidOperationException(
                    "Clipboard package did not preserve the selected object count.");
            }

            string encoded =
                EditorClipboardCodec.Encode(
                    package);
            EditorClipboardPackage decoded;

            if (!EditorClipboardCodec.TryDecode(
                    encoded,
                    out decoded) ||
                decoded == null ||
                decoded.Objects.Count != 3)
            {
                throw new InvalidOperationException(
                    "Clipboard package codec round-trip failed.");
            }

            if (decoded.Objects[0].Kind !=
                    EditorObjectKind.Shape ||
                decoded.Objects[1].Kind !=
                    EditorObjectKind.Image ||
                decoded.Objects[2].Kind !=
                    EditorObjectKind.TextBox)
            {
                throw new InvalidOperationException(
                    "Clipboard package did not preserve relative z-order.");
            }

            if (decoded.Objects[0].Shape == null ||
                decoded.Objects[0].Shape.Name !=
                    "Clipboard diamond" ||
                decoded.Objects[1].Image == null ||
                decoded.Objects[1].Image.Data == null ||
                decoded.Objects[1].Image.Data.Length != 6 ||
                decoded.Objects[2].TextBox == null ||
                decoded.Objects[2].TextBox.Text !=
                    "Clipboard text" ||
                !decoded.Objects[2].TextBox.Bold)
            {
                throw new InvalidOperationException(
                    "Clipboard package object properties were not preserved.");
            }
        }

        private static EditorSelectionEntry Selection(
            EditorObjectKind kind,
            int index)
        {
            EditorSelectionEntry entry =
                new EditorSelectionEntry();
            entry.Kind = kind;
            entry.Index = index;
            return entry;
        }
    }
}
