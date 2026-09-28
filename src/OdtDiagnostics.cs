using System;
using System.IO;
using System.IO.Compression;

namespace PptxViewer
{
    internal static class OdtDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            TextDocument document = TextDocument.CreateNew("ODT Writer Self Test");

            DocumentParagraph first = document.AddParagraph("PowerPointLite ODT independent writer self-test");
            first.Alignment = DocumentParagraphAlignment.Center;
            first.Runs[0].FontFamily = "Arial";
            first.Runs[0].FontSizePoints = 18f;
            first.Runs[0].Bold = true;
            first.Runs[0].ColorHex = "356AE6";

            DocumentParagraph second = document.AddParagraph("Styled text");
            DocumentTextRun extra = second.AddRun(" / italic underline");
            extra.FontFamily = "Arial";
            extra.FontSizePoints = 11f;
            extra.Italic = true;
            extra.Underline = true;
            extra.ColorHex = "5B8CFF";

            OdtWriter.Save(document, outputPath);
            ValidatePackage(outputPath);

            OdtEditSafetyResult safety = OdtEditSafety.Analyze(outputPath);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "ODT safety analysis failed." : safety.Warning);

            TextDocument reopened = OdtReader.Read(outputPath);
            if (reopened.Paragraphs.Count < 2 || reopened.Paragraphs[0].Runs.Count == 0)
                throw new InvalidOperationException("ODT reader did not recover the generated document.");

            reopened.Paragraphs[0].Runs[0].Text = "Edited after ODT round-trip";
            reopened.Paragraphs[0].Runs[0].Bold = true;
            OdtWriter.Save(reopened, outputPath);
            ValidatePackage(outputPath);

            TextDocument secondRead = OdtReader.Read(outputPath);
            if (secondRead.Paragraphs.Count == 0 ||
                secondRead.Paragraphs[0].Runs.Count == 0 ||
                !string.Equals(
                    secondRead.Paragraphs[0].Runs[0].Text,
                    "Edited after ODT round-trip",
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException("ODT edit/save/read round-trip verification failed.");
            }
        }

        private static void ValidatePackage(string path)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("ODT writer output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                RequireEntry(archive, "mimetype");
                RequireEntry(archive, "content.xml");
                RequireEntry(archive, "styles.xml");
                RequireEntry(archive, "meta.xml");
                RequireEntry(archive, "META-INF/manifest.xml");

                string mime = OdfPackageUtility.ReadTextEntry(archive, "mimetype").Trim();
                if (!string.Equals(mime, OdfPackageUtility.TextMimeType, StringComparison.Ordinal))
                    throw new InvalidOperationException("ODT mimetype is invalid.");
            }
        }

        private static void RequireEntry(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
                throw new InvalidOperationException("Required ODT part is missing: " + path);
        }
    }
}
