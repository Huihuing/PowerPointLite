using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxWriterDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            PresentationDocument document =
                PresentationDocument.CreateNew("Writer Self Test");

            PresentationSlide second = document.AddSlide("Second Slide");
            PresentationTextBox body = second.AddTextBox(
                "This presentation was generated entirely by the PowerPointLite writer.\nNo Microsoft Office binary or template is bundled.");
            body.Name = "Body";
            body.X = 1371600;
            body.Y = 2743200;
            body.Width = 9448800;
            body.Height = 1828800;
            body.FontSizePoints = 18f;
            body.Alignment = PresentationTextAlignment.Center;
            body.ColorHex = "44546A";

            PresentationShape shape = second.AddShape(
                PresentationShapeKind.RoundedRectangle);
            shape.Name = "Generated Shape";
            shape.X = 914400;
            shape.Y = 4800600;
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
            generatedImage.Y = 4572000;
            generatedImage.Width = 2286000;
            generatedImage.Height = 1371600;

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

            PptxWriter.Save(document, outputPath);
            ValidatePackage(outputPath, document.Slides.Count, true);
            ValidateEditableRoundTrip(outputPath);
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

            second.TextBoxes[1].Text =
                "Round-trip edit completed successfully.";
            second.Shapes[0].FillColorHex = "F4A261";

            string roundTripPath = path + ".roundtrip.pptx";

            try
            {
                if (File.Exists(roundTripPath))
                    File.Delete(roundTripPath);

                PptxWriter.Save(loaded.Document, roundTripPath);
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
            }
        }

        private static void RequireEntry(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
                throw new InvalidOperationException("Required PPTX part is missing: " + path);
        }
    }
}
