using System;
using System.IO;

namespace PptxViewer
{
    internal static class DocumentConversionDiagnostics
    {
        public static void CreateAndValidate(string outputDirectory)
        {
            if (string.IsNullOrEmpty(outputDirectory))
                outputDirectory = AppDomain.CurrentDomain.BaseDirectory;

            Directory.CreateDirectory(outputDirectory);

            ValidateTextConversions(outputDirectory);
            ValidateSpreadsheetConversions(outputDirectory);
            ValidatePresentationConversions(outputDirectory);
        }

        private static void ValidateTextConversions(string directory)
        {
            TextDocument model = TextDocument.CreateNew("Conversion Text Test");
            DocumentParagraph paragraph = model.AddParagraph(
                "One internal text model should be writable as DOCX, HWPX and ODT.");
            if (paragraph.Runs.Count > 0)
            {
                paragraph.Runs[0].Bold = true;
                paragraph.Runs[0].FontSizePoints = 13f;
            }

            string docx = Path.Combine(directory, "CONVERT_SOURCE.docx");
            string hwpx = Path.Combine(directory, "CONVERT_TEXT.hwpx");
            string odt = Path.Combine(directory, "CONVERT_TEXT.odt");

            DocxWriter.Save(model, docx);

            ConvertibleDocument source = SafeDocumentConversion.Load(docx);
            if (source.Kind != ConvertibleDocumentKind.Text || source.TextDocument == null)
                throw new InvalidOperationException("DOCX conversion model was not loaded as text.");

            SafeDocumentConversion.SaveAs(source, hwpx);
            SafeDocumentConversion.SaveAs(source, odt);

            TextDocument hwpxRead = HwpxReader.Read(hwpx);
            TextDocument odtRead = OdtReader.Read(odt);

            if (hwpxRead.Paragraphs.Count == 0 || odtRead.Paragraphs.Count == 0)
                throw new InvalidOperationException("Text conversion output could not be read back.");
        }

        private static void ValidateSpreadsheetConversions(string directory)
        {
            SpreadsheetDocument model = SpreadsheetDocument.CreateNew("Conversion Spreadsheet Test");
            SpreadsheetSheet sheet = model.Sheets[0];
            sheet.SetInput(1, 1, "Label");
            sheet.SetInput(1, 2, "Value");
            sheet.SetInput(2, 1, "Alpha");
            sheet.SetInput(2, 2, "42.5");
            sheet.SetInput(3, 2, "=B2*2");

            string xlsx = Path.Combine(directory, "CONVERT_SOURCE.xlsx");
            string ods = Path.Combine(directory, "CONVERT_SHEET.ods");

            XlsxWriter.Save(model, xlsx);

            ConvertibleDocument source = SafeDocumentConversion.Load(xlsx);
            if (source.Kind != ConvertibleDocumentKind.Spreadsheet ||
                source.SpreadsheetDocument == null)
            {
                throw new InvalidOperationException("XLSX conversion model was not loaded as a spreadsheet.");
            }

            SafeDocumentConversion.SaveAs(source, ods);
            SpreadsheetDocument reopened = OdsReader.Read(ods);

            SpreadsheetCell value = reopened.Sheets[0].GetCell(2, 2, false);
            if (value == null || Math.Abs(value.NumberValue - 42.5) > 0.00001)
                throw new InvalidOperationException("XLSX → ODS numeric round-trip failed.");
        }

        private static void ValidatePresentationConversions(string directory)
        {
            PresentationDocument model = PresentationDocument.CreateNew("Conversion Presentation Test");
            PresentationSlide slide = model.AddSlide("Second slide");
            PresentationTextBox text = slide.AddTextBox("PPTX and ODP conversion test");
            text.Alignment = PresentationTextAlignment.Center;

            PresentationShape shape = slide.AddShape(PresentationShapeKind.Ellipse);
            shape.FillColorHex = "5B8CFF";
            shape.LineColorHex = "356AE6";

            string pptx = Path.Combine(directory, "CONVERT_SOURCE.pptx");
            string odp = Path.Combine(directory, "CONVERT_PRESENTATION.odp");

            PptxWriter.Save(model, pptx);

            ConvertibleDocument source = SafeDocumentConversion.Load(pptx);
            if (source.Kind != ConvertibleDocumentKind.Presentation ||
                source.PresentationDocument == null)
            {
                throw new InvalidOperationException("PPTX conversion model was not loaded as a presentation.");
            }

            SafeDocumentConversion.SaveAs(source, odp);
            PresentationDocument reopened = OdpReader.Read(odp);

            if (reopened.Slides.Count < 2)
                throw new InvalidOperationException("PPTX → ODP slide round-trip failed.");
        }
    }
}
