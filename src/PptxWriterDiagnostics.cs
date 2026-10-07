using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxWriterDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            EditorClipboardCodecDiagnostics.Validate();

            PresentationDocument document =
                PresentationDocument.CreateNew("Writer Self Test");

            PresentationSlide second = document.AddSlide("Second Slide");
            PresentationTextBox body = second.AddTextBox(
                "This presentation was generated entirely by the PowerPointLite writer.\nNo Microsoft Office binary or template is bundled.");
            body.Name = "Body";
            body.X = 1371600;
            body.Y = 2743200;
            body.Width = 9448800;
            body.Height = 1371600;
            body.FontSizePoints = 18f;
            body.Alignment = PresentationTextAlignment.Center;
            body.ColorHex = "44546A";

            PresentationShape shape = second.AddShape(
                PresentationShapeKind.RoundedRectangle);
            shape.Name = "Generated Shape";
            shape.X = 914400;
            shape.Y = 4572000;
            shape.Width = 2743200;
            shape.Height = 1143000;
            shape.FillColorHex = "31B6A1";
            shape.LineColorHex = "247F72";

            PresentationImage generatedImage = second.AddImage(
                CreateGeneratedPng(),
                "png",
                "image/png");
            generatedImage.Name = "Generated test image";
            generatedImage.X = 8229600;
            generatedImage.Y = 4343400;
            generatedImage.Width = 2286000;
            generatedImage.Height = 1371600;

            if (!second.MoveObjectToFront(
                    PresentationLayerKind.Shape,
                    0))
            {
                throw new InvalidOperationException(
                    "Writer self-test could not move the shape in front of mixed object types.");
            }

            if (!second.MoveObjectToBack(
                    PresentationLayerKind.Image,
                    0))
            {
                throw new InvalidOperationException(
                    "Writer self-test could not move the image behind mixed object types.");
            }

            PresentationTable table = second.AddTable(2, 3);
            table.Name = "Generated Table";
            table.X = 1371600;
            table.Y = 960120;
            table.Width = 9448800;
            table.Height = 1143000;

            for (int row = 0; row < table.Rows; row++)
            {
                for (int column = 0; column < table.Columns; column++)
                {
                    PresentationTableCell cell = table.GetCell(row, column);
                    cell.Text = "R" + (row + 1).ToString() + " C" + (column + 1).ToString();
                    cell.FontSizePoints = 13f;
                    cell.Alignment = PresentationTextAlignment.Center;
                    cell.FillColorHex = row == 0 ? "EAF0FF" : "FFFFFF";
                    cell.Bold = row == 0;
                }
            }

            PresentationSlide third = document.AddSlide("Formatting Test");
            PresentationTextBox sample = third.AddTextBox("Bold / italic writer test");
            sample.Name = "Formatting sample";
            sample.X = 1371600;
            sample.Y = 2514600;
            sample.Width = 9448800;
            sample.Height = 1371600;
            sample.FontSizePoints = 24f;
            sample.Bold = true;
            sample.Italic = true;
            sample.Alignment = PresentationTextAlignment.Center;
            sample.ColorHex = "5B8CFF";

            PresentationPackageWriter.Save(document, outputPath);
            ValidatePackage(outputPath, document.Slides.Count, true);
            ValidateEditableRoundTrip(outputPath);
            ValidateAlternateContentPreparation(outputPath);
        }

        private static byte[] CreateGeneratedPng()
        {
            using (Bitmap bitmap = new Bitmap(96, 56))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (MemoryStream stream = new MemoryStream())
            {
                graphics.Clear(Color.FromArgb(245, 246, 248));

                using (Brush accent = new SolidBrush(Color.FromArgb(91, 140, 255)))
                    graphics.FillRectangle(accent, 8, 8, 80, 40);

                using (Pen line = new Pen(Color.FromArgb(32, 36, 42), 2f))
                    graphics.DrawRectangle(line, 8, 8, 80, 40);

                bitmap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }

        private static void ValidateEditableRoundTrip(string path)
        {
            PptxEditableLoadResult loaded = PptxEditableReader.Read(path);

            if (loaded == null || loaded.Document == null)
                throw new InvalidOperationException("Editable reader returned no document.");

            if (!loaded.CanRoundTripSafely)
                throw new InvalidOperationException("Writer output was not considered safe for editable round-trip: " + loaded.Warning);

            if (loaded.Document.Slides.Count != 3)
                throw new InvalidOperationException("Editable reader did not preserve the expected slide count.");

            PresentationSlide second = loaded.Document.Slides[1];

            if (second.TextBoxes.Count < 2)
                throw new InvalidOperationException("Editable reader did not preserve the expected text boxes.");

            if (second.Shapes.Count != 1)
                throw new InvalidOperationException("Editable reader did not preserve the generated shape.");

            if (second.Images.Count != 1 ||
                second.Images[0].Data == null ||
                second.Images[0].Data.Length == 0)
            {
                throw new InvalidOperationException("Editable reader did not preserve the generated image.");
            }

            second.SynchronizeObjectOrder();

            if (second.ObjectOrder.Count < 4 ||
                second.ObjectOrder[0].Kind !=
                    PresentationLayerKind.Image ||
                second.ObjectOrder[second.ObjectOrder.Count - 1].Kind !=
                    PresentationLayerKind.Shape)
            {
                throw new InvalidOperationException(
                    "Editable reader did not preserve mixed object z-order.");
            }

            if (second.Tables.Count != 1 ||
                second.Tables[0].Rows != 2 ||
                second.Tables[0].Columns != 3 ||
                second.Tables[0].GetCell(0, 1).Text != "R1 C2")
            {
                throw new InvalidOperationException("Editable reader did not preserve the generated table.");
            }

            second.TextBoxes[1].Text =
                "Round-trip edit completed successfully.";
            second.Shapes[0].FillColorHex = "F4A261";
            second.Tables[0].GetCell(1, 2).Text = "Table round-trip OK";

            string roundTripPath = path + ".roundtrip.pptx";

            try
            {
                if (File.Exists(roundTripPath))
                    File.Delete(roundTripPath);

                PresentationPackageWriter.Save(loaded.Document, roundTripPath);
                ValidatePackage(roundTripPath, loaded.Document.Slides.Count, true);

                PptxEditableLoadResult secondRead =
                    PptxEditableReader.Read(roundTripPath);

                if (!secondRead.CanRoundTripSafely ||
                    secondRead.Document.Slides.Count != 3)
                {
                    throw new InvalidOperationException("Second editable read failed after Writer save.");
                }

                PresentationSlide verify = secondRead.Document.Slides[1];

                if (verify.TextBoxes.Count < 2 ||
                    verify.TextBoxes[1].Text !=
                        "Round-trip edit completed successfully.")
                {
                    throw new InvalidOperationException("Read-edit-write text round-trip verification failed.");
                }

                if (verify.Shapes.Count != 1 ||
                    !string.Equals(
                        verify.Shapes[0].FillColorHex,
                        "F4A261",
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Shape round-trip verification failed.");
                }

                if (verify.Images.Count != 1 ||
                    verify.Images[0].Data == null ||
                    verify.Images[0].Data.Length == 0)
                {
                    throw new InvalidOperationException("Image round-trip verification failed.");
                }

                verify.SynchronizeObjectOrder();

                if (verify.ObjectOrder.Count < 4 ||
                    verify.ObjectOrder[0].Kind !=
                        PresentationLayerKind.Image ||
                    verify.ObjectOrder[verify.ObjectOrder.Count - 1].Kind !=
                        PresentationLayerKind.Shape)
                {
                    throw new InvalidOperationException(
                        "Mixed object z-order was not preserved after read-edit-write.");
                }

                if (verify.Tables.Count != 1 ||
                    verify.Tables[0].GetCell(1, 2).Text != "Table round-trip OK")
                {
                    throw new InvalidOperationException("Table round-trip verification failed.");
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(roundTripPath))
                        File.Delete(roundTripPath);
                }
                catch { }
            }
        }

        private static void ValidateAlternateContentPreparation(string sourcePath)
        {
            string testPath = sourcePath + ".alternate-content.pptx";
            string prepared = null;

            try
            {
                if (File.Exists(testPath))
                    File.Delete(testPath);

                File.Copy(sourcePath, testPath);
                InjectAlternateContentMarker(testPath);

                if (!PptxRenderPreprocessor.NeedsAlternateContentFallback(testPath))
                {
                    throw new InvalidOperationException(
                        "AlternateContent render-preparation test was not detected.");
                }

                prepared = PptxRenderPreprocessor.PrepareForRendering(testPath);

                if (string.IsNullOrEmpty(prepared) ||
                    !File.Exists(prepared) ||
                    string.Equals(prepared, testPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "AlternateContent render preparation did not create a temporary package.");
                }

                using (ZipArchive archive = ZipFile.OpenRead(prepared))
                {
                    ZipArchiveEntry slide = archive.GetEntry("ppt/slides/slide3.xml");
                    if (slide == null)
                        throw new InvalidOperationException("Prepared slide3.xml is missing.");

                    string xml;
                    using (Stream stream = slide.Open())
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
                        xml = reader.ReadToEnd();

                    if (xml.IndexOf("AlternateContent", StringComparison.Ordinal) >= 0)
                    {
                        throw new InvalidOperationException(
                            "Prepared package still contains AlternateContent.");
                    }

                    if (xml.IndexOf("FALLBACK_MARKER", StringComparison.Ordinal) < 0 ||
                        xml.IndexOf("CHOICE_MARKER", StringComparison.Ordinal) >= 0)
                    {
                        throw new InvalidOperationException(
                            "AlternateContent fallback branch was not selected correctly.");
                    }
                }
            }
            finally
            {
                PptxRenderPreprocessor.DeletePreparedFile(prepared, testPath);

                try
                {
                    if (File.Exists(testPath))
                        File.Delete(testPath);
                }
                catch { }
            }
        }

        private static void InjectAlternateContentMarker(string path)
        {
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                ZipArchiveEntry entry = archive.GetEntry("ppt/slides/slide3.xml");
                if (entry == null)
                    throw new InvalidOperationException("slide3.xml is missing for AlternateContent test.");

                XmlDocument document = new XmlDocument();
                using (Stream stream = entry.Open())
                    document.Load(stream);

                XmlNode shapeTree = FindFirstByLocalName(document.DocumentElement, "spTree");
                if (shapeTree == null)
                    throw new InvalidOperationException("slide3 shape tree was not found.");

                const string mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                const string p = "http://schemas.openxmlformats.org/presentationml/2006/main";

                XmlElement alternate = document.CreateElement("mc", "AlternateContent", mc);
                XmlElement choice = document.CreateElement("mc", "Choice", mc);
                choice.SetAttribute("Requires", "p14");
                XmlElement choiceMarker = document.CreateElement("p", "compatMarker", p);
                choiceMarker.SetAttribute("value", "CHOICE_MARKER");
                choice.AppendChild(choiceMarker);

                XmlElement fallback = document.CreateElement("mc", "Fallback", mc);
                XmlElement fallbackMarker = document.CreateElement("p", "compatMarker", p);
                fallbackMarker.SetAttribute("value", "FALLBACK_MARKER");
                fallback.AppendChild(fallbackMarker);

                alternate.AppendChild(choice);
                alternate.AppendChild(fallback);
                shapeTree.AppendChild(alternate);

                entry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(
                    "ppt/slides/slide3.xml",
                    CompressionLevel.Optimal);

                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                settings.Indent = false;

                using (Stream stream = replacement.Open())
                using (XmlWriter writer = XmlWriter.Create(stream, settings))
                    document.Save(writer);
            }
        }

        private static XmlNode FindFirstByLocalName(XmlNode node, string localName)
        {
            if (node == null)
                return null;

            if (node.LocalName == localName)
                return node;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirstByLocalName(node.ChildNodes[i], localName);
                if (found != null)
                    return found;
            }

            return null;
        }

        private static void ValidatePackage(
            string path,
            int expectedSlides,
            bool expectGeneratedImage)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new InvalidOperationException("Writer output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                RequireEntry(archive, "[Content_Types].xml");
                RequireEntry(archive, "_rels/.rels");
                RequireEntry(archive, "ppt/presentation.xml");
                RequireEntry(archive, "ppt/_rels/presentation.xml.rels");
                RequireEntry(archive, "ppt/slideMasters/slideMaster1.xml");
                RequireEntry(archive, "ppt/slideLayouts/slideLayout1.xml");
                RequireEntry(archive, "ppt/theme/theme1.xml");

                for (int i = 1; i <= expectedSlides; i++)
                {
                    RequireEntry(
                        archive,
                        "ppt/slides/slide" + i.ToString() + ".xml");
                    RequireEntry(
                        archive,
                        "ppt/slides/_rels/slide" + i.ToString() + ".xml.rels");
                }

                if (expectGeneratedImage)
                    RequireEntry(archive, "ppt/media/slide2_image1.png");

                XmlDocument presentation =
                    OpcPackageUtility.ReadXmlPart(archive, "ppt/presentation.xml");

                if (presentation == null)
                    throw new InvalidOperationException("presentation.xml could not be parsed.");

                XmlNodeList slideIds = presentation.GetElementsByTagName(
                    "p:sldId");

                if (slideIds == null || slideIds.Count != expectedSlides)
                {
                    throw new InvalidOperationException(
                        "Expected " + expectedSlides.ToString() +
                        " slides but presentation.xml contains " +
                        (slideIds == null ? "0" : slideIds.Count.ToString()) + ".");
                }

                XmlDocument slide2 =
                    OpcPackageUtility.ReadXmlPart(archive, "ppt/slides/slide2.xml");

                if (slide2 == null ||
                    slide2.GetElementsByTagName("a:tbl").Count != 1)
                {
                    throw new InvalidOperationException("Generated table XML was not found in slide2.xml.");
                }
            }
        }

        private static void RequireEntry(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
                throw new InvalidOperationException("Required PPTX part is missing: " + path);
        }
    }
}
