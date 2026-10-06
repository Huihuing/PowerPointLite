using System;
using System.IO;
using System.IO.Compression;

namespace PptxViewer
{
    internal static class HwpxDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            TextDocument document = TextDocument.CreateNew("HWPX Writer Self Test");

            DocumentParagraph first = document.AddParagraph("PowerPointLite HWPX clean-room writer self-test");
            first.Alignment = DocumentParagraphAlignment.Center;
            DocumentTextRun firstRun = first.Runs[0];
            firstRun.FontFamily = "Arial";
            firstRun.FontSizePoints = 18f;
            firstRun.Bold = true;
            firstRun.ColorHex = "356AE6";

            DocumentParagraph second =
                document.AddParagraph(string.Empty);
            DocumentTextRun korean =
                second.AddRun(
                    "HWPX package create / read / edit / save test");
            korean.FontFamily = "Arial";
            korean.FontSizePoints = 11f;

            DocumentTextRun emphasis = second.AddRun(" - styled run");
            emphasis.FontFamily = "Arial";
            emphasis.FontSizePoints = 11f;
            emphasis.Italic = true;
            emphasis.Underline = true;
            emphasis.ColorHex = "5B8CFF";

            HwpxWriter.Save(document, outputPath);
            ValidatePackage(outputPath);

            HwpxEditSafetyResult safety = HwpxEditSafety.Analyze(outputPath);
            if (safety == null || !safety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    safety == null
                        ? "HWPX edit safety analysis returned no result."
                        : safety.Warning);
            }

            TextDocument reopened = HwpxReader.Read(outputPath);
            if (reopened.Paragraphs.Count < 2)
                throw new InvalidOperationException("HWPX reader did not recover the generated paragraphs.");

            DocumentParagraph edited = reopened.Paragraphs[0];
            if (edited.Runs.Count == 0)
                edited.AddRun(string.Empty);
            edited.Runs[0].Text = "Edited after HWPX round-trip";
            edited.Runs[0].Bold = true;

            HwpxWriter.Save(reopened, outputPath);
            ValidatePackage(outputPath);

            TextDocument secondRead = HwpxReader.Read(outputPath);
            if (secondRead.Paragraphs.Count == 0 ||
                secondRead.Paragraphs[0].Runs.Count == 0 ||
                !string.Equals(
                    secondRead.Paragraphs[0].Runs[0].Text,
                    "Edited after HWPX round-trip",
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException("HWPX edit/save/read round-trip verification failed.");
            }
        }

        private static void ValidatePackage(string path)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("HWPX writer output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                RequireEntry(archive, "mimetype");
                RequireEntry(archive, "version.xml");
                RequireEntry(archive, "META-INF/container.xml");
                RequireEntry(archive, "META-INF/manifest.xml");
                RequireEntry(archive, "Contents/content.hpf");
                RequireEntry(archive, "Contents/header.xml");
                RequireEntry(archive, "Contents/section0.xml");
                RequireEntry(archive, "Preview/PrvText.txt");

                ZipArchiveEntry mimetype = archive.GetEntry("mimetype");
                using (Stream stream = mimetype.Open())
                using (StreamReader reader = new StreamReader(stream))
                {
                    string value = reader.ReadToEnd().Trim();
                    if (!string.Equals(value, "application/hwp+zip", StringComparison.Ordinal))
                        throw new InvalidOperationException("HWPX mimetype is invalid.");
                }
            }
        }

        private static void RequireEntry(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
                throw new InvalidOperationException("Required HWPX part is missing: " + path);
        }
    }
}
