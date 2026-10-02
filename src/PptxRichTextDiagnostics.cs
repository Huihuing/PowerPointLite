using System;
using System.IO;

namespace PptxViewer
{
    internal static class PptxRichTextDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            outputPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
            catch { }

            PresentationDocument document =
                PresentationDocument.CreateNew("Rich Text Self Test");
            PresentationTextBox box = document.Slides[0].TextBoxes[0];
            box.Name = "Rich text sample";
            box.FontFamily = "Arial";
            box.FontSizePoints = 22f;
            box.ColorHex = "20242A";

            PresentationTextParagraph first =
                new PresentationTextParagraph();
            first.Alignment = PresentationTextAlignment.Center;
            first.BulletText = "•";
            first.SpaceBeforePoints = 6f;
            first.SpaceAfterPoints = 9f;

            PresentationTextRun run1 = new PresentationTextRun();
            run1.Text = "PowerPointLite ";
            run1.FontFamily = "Arial";
            run1.FontSizePoints = 22f;
            run1.Bold = true;
            run1.ColorHex = "5B8CFF";
            first.Runs.Add(run1);

            PresentationTextRun run2 = new PresentationTextRun();
            run2.Text = "Rich Text";
            run2.FontFamily = "Arial";
            run2.FontSizePoints = 18f;
            run2.Italic = true;
            run2.Underline = true;
            run2.ColorHex = "31B6A1";
            first.Runs.Add(run2);

            PresentationTextParagraph second =
                new PresentationTextParagraph();
            second.Alignment = PresentationTextAlignment.Left;
            second.Level = 1;

            PresentationTextRun run3 = new PresentationTextRun();
            run3.Text = "Baseline";
            run3.FontFamily = "Arial";
            run3.FontSizePoints = 16f;
            run3.BaselinePercent = 12;
            run3.ColorHex = "F4A261";
            second.Runs.Add(run3);

            box.SetRichParagraphs(
                new PresentationTextParagraph[] { first, second });

            PresentationPackageWriter.Save(document, outputPath);

            PptxEditableLoadResult loaded =
                PptxEditableReader.Read(outputPath);
            if (loaded == null || loaded.Document == null)
                throw new InvalidOperationException("Rich text package could not be reopened.");
            if (!loaded.CanRoundTripSafely)
                throw new InvalidOperationException(
                    "Generated rich text package was not considered safe for round-trip: " +
                    (loaded.Warning ?? string.Empty));

            PptxRichTextPackage.ReadIntoDocument(
                outputPath,
                loaded.Document);

            PresentationTextBox verify =
                loaded.Document.Slides[0].TextBoxes[0];
            if (verify == null || !verify.HasRichText ||
                verify.RichParagraphs.Count != 2)
            {
                throw new InvalidOperationException(
                    "Rich text paragraphs were not restored after round-trip.");
            }

            PresentationTextParagraph verifyFirst =
                verify.RichParagraphs[0];
            PresentationTextParagraph verifySecond =
                verify.RichParagraphs[1];

            if (verifyFirst.Runs.Count != 2 ||
                verifyFirst.Alignment != PresentationTextAlignment.Center ||
                verifyFirst.BulletText != "•" ||
                Math.Abs(verifyFirst.SpaceBeforePoints - 6f) > 0.01f ||
                Math.Abs(verifyFirst.SpaceAfterPoints - 9f) > 0.01f)
            {
                throw new InvalidOperationException(
                    "Rich paragraph formatting was not preserved.");
            }

            PresentationTextRun verifyRun1 = verifyFirst.Runs[0];
            PresentationTextRun verifyRun2 = verifyFirst.Runs[1];

            if (verifyRun1.Text != "PowerPointLite " ||
                !verifyRun1.Bold ||
                !string.Equals(verifyRun1.ColorHex, "5B8CFF", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The first rich text run was not preserved.");
            }

            if (verifyRun2.Text != "Rich Text" ||
                !verifyRun2.Italic ||
                !verifyRun2.Underline ||
                Math.Abs(verifyRun2.FontSizePoints - 18f) > 0.01f ||
                !string.Equals(verifyRun2.ColorHex, "31B6A1", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The second rich text run was not preserved.");
            }

            if (verifySecond.Runs.Count != 1 ||
                verifySecond.Level != 1 ||
                verifySecond.Runs[0].Text != "Baseline" ||
                verifySecond.Runs[0].BaselinePercent != 12)
            {
                throw new InvalidOperationException(
                    "Rich text level/baseline formatting was not preserved.");
            }
        }
    }
}
