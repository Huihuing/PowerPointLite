using System;
using System.IO;
using System.IO.Compression;

namespace PptxViewer
{
    internal static class OdpDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            PresentationDocument document =
                PresentationDocument.CreateNew("ODP Writer Self Test");

            PresentationSlide second = document.AddSlide("Second ODP Slide");
            PresentationTextBox text = second.AddTextBox(
                "Independent ODP reader/writer round-trip test");
            text.X = 1371600;
            text.Y = 2286000;
            text.Width = 9448800;
            text.Height = 1371600;
            text.FontSizePoints = 20f;
            text.Alignment = PresentationTextAlignment.Center;

            PresentationShape shape = second.AddShape(PresentationShapeKind.Ellipse);
            shape.X = 4572000;
            shape.Y = 4114800;
            shape.Width = 3048000;
            shape.Height = 1371600;
            shape.FillColorHex = "5B8CFF";

            OdpWriter.Save(document, outputPath);
            ValidatePackage(outputPath);

            OdpEditSafetyResult safety = OdpEditSafety.Analyze(outputPath);
            if (safety == null || !safety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    safety == null
                        ? "ODP safety analysis failed."
                        : safety.Warning);
            }

            PresentationDocument reopened = OdpReader.Read(outputPath);
            if (reopened == null || reopened.Slides.Count < 2)
                throw new InvalidOperationException("ODP slide round-trip failed.");

            if (reopened.Slides[1].TextBoxes.Count == 0)
                throw new InvalidOperationException("ODP text object was not recovered.");

            reopened.Slides[1].TextBoxes[0].Text =
                "Edited after ODP round-trip";
            OdpWriter.Save(reopened, outputPath);
            ValidatePackage(outputPath);

            PresentationDocument verify = OdpReader.Read(outputPath);
            if (verify.Slides.Count < 2 ||
                verify.Slides[1].TextBoxes.Count == 0 ||
                !string.Equals(
                    verify.Slides[1].TextBoxes[0].Text,
                    "Edited after ODP round-trip",
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "ODP edit/save/read round-trip verification failed.");
            }
        }

        private static void ValidatePackage(string path)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("ODP writer output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                Require(archive, "mimetype");
                Require(archive, "content.xml");
                Require(archive, "styles.xml");
                Require(archive, "meta.xml");
                Require(archive, "META-INF/manifest.xml");

                string mime = OdfPackageUtility
                    .ReadTextEntry(archive, "mimetype")
                    .Trim();

                if (!string.Equals(
                        mime,
                        OdfPackageUtility.PresentationMimeType,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("ODP mimetype is invalid.");
                }
            }
        }

        private static void Require(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
            {
                throw new InvalidOperationException(
                    "Required ODP part is missing: " + path);
            }
        }
    }
}
