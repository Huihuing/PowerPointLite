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
        public PresentationTable Table { get; set; }

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
            copy.Table =
                Table == null
                    ? null
                    : Table.Clone();
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

            if (Kind == EditorObjectKind.Table &&
                Table != null)
            {
                x = Table.X;
                y = Table.Y;
                width = Table.Width;
                height = Table.Height;
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
            else if (Kind == EditorObjectKind.Table &&
                     Table != null)
            {
                Table.X += dx;
                Table.Y += dy;
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

            if (kind == EditorObjectKind.Table &&
                index >= 0 &&
                index < slide.Tables.Count)
            {
                EditorClipboardObject item =
                    new EditorClipboardObject();
                item.Kind = kind;
                item.Table =
                    slide.Tables[index].Clone();
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

            if (kind == PresentationLayerKind.Table)
                return EditorObjectKind.Table;

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
            else if (item.Kind == EditorObjectKind.Table)
                WriteTable(writer, item.Table);
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
            else if (kind == EditorObjectKind.Table)
                item.Table = ReadTable(reader);
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

        private static void WriteTable(
            BinaryWriter writer,
            PresentationTable table)
        {
            if (table == null)
                throw new InvalidDataException(
                    "Clipboard table is missing.");

            WriteString(writer, table.Name);
            writer.Write(table.Rows);
            writer.Write(table.Columns);
            writer.Write(table.X);
            writer.Write(table.Y);
            writer.Write(table.Width);
            writer.Write(table.Height);

            for (int row = 0;
                 row < table.Rows;
                 row++)
            {
                for (int column = 0;
                     column < table.Columns;
                     column++)
                {
                    PresentationTableCell cell =
                        table.GetCell(
                            row,
                            column) ??
                        new PresentationTableCell();
                    WriteString(writer, cell.Text);
                    WriteString(
                        writer,
                        cell.FontFamily);
                    writer.Write(
                        cell.FontSizePoints);
                    writer.Write(cell.Bold);
                    WriteString(
                        writer,
                        cell.TextColorHex);
                    WriteString(
                        writer,
                        cell.FillColorHex);
                    writer.Write(
                        (int)cell.Alignment);
                }
            }
        }

        private static PresentationTable ReadTable(
            BinaryReader reader)
        {
            string name = reader.ReadString();
            int rows = reader.ReadInt32();
            int columns = reader.ReadInt32();

            if (rows < 1 || rows > 50 ||
                columns < 1 || columns > 50)
            {
                throw new InvalidDataException(
                    "Clipboard table dimensions are invalid.");
            }

            PresentationTable table =
                new PresentationTable(
                    rows,
                    columns);
            table.Name = name;
            table.X = reader.ReadInt64();
            table.Y = reader.ReadInt64();
            table.Width = reader.ReadInt64();
            table.Height = reader.ReadInt64();

            for (int row = 0;
                 row < rows;
                 row++)
            {
                for (int column = 0;
                     column < columns;
                     column++)
                {
                    PresentationTableCell cell =
                        table.GetCell(
                            row,
                            column);
                    cell.Text =
                        reader.ReadString();
                    cell.FontFamily =
                        reader.ReadString();
                    cell.FontSizePoints =
                        reader.ReadSingle();
                    cell.Bold =
                        reader.ReadBoolean();
                    cell.TextColorHex =
                        reader.ReadString();
                    cell.FillColorHex =
                        reader.ReadString();
                    cell.Alignment =
                        (PresentationTextAlignment)
                        reader.ReadInt32();
                }
            }

            return table;
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

            PresentationTable table =
                slide.AddTable(2, 2);
            table.Name = "Clipboard table";
            table.X = 2743200;
            table.Y = 3657600;
            table.GetCell(0, 0).Text =
                "Table cell";
            table.GetCell(0, 0).Bold = true;
            table.GetCell(0, 0).FillColorHex =
                "DDE8FF";

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
            selection.Add(
                Selection(
                    EditorObjectKind.Table,
                    slide.Tables.Count - 1));

            EditorClipboardPackage package =
                EditorClipboardCodec.CreatePackage(
                    slide,
                    selection);

            if (package == null ||
                package.Objects.Count != 4)
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
                decoded.Objects.Count != 4)
            {
                throw new InvalidOperationException(
                    "Clipboard package codec round-trip failed.");
            }

            if (decoded.Objects[0].Kind !=
                    EditorObjectKind.Shape ||
                decoded.Objects[1].Kind !=
                    EditorObjectKind.Image ||
                decoded.Objects[2].Kind !=
                    EditorObjectKind.TextBox ||
                decoded.Objects[3].Kind !=
                    EditorObjectKind.Table)
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
                !decoded.Objects[2].TextBox.Bold ||
                decoded.Objects[3].Table == null ||
                decoded.Objects[3].Table.Rows != 2 ||
                decoded.Objects[3].Table.Columns != 2 ||
                decoded.Objects[3].Table.GetCell(
                    0,
                    0).Text !=
                    "Table cell" ||
                !decoded.Objects[3].Table.GetCell(
                    0,
                    0).Bold)
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
