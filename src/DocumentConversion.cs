using System;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal enum ConvertibleDocumentKind
    {
        Text,
        Spreadsheet,
        Presentation
    }

    internal sealed class ConvertibleDocument
    {
        public ConvertibleDocumentKind Kind { get; set; }
        public TextDocument TextDocument { get; set; }
        public SpreadsheetDocument SpreadsheetDocument { get; set; }
        public PresentationDocument PresentationDocument { get; set; }
        public bool IsLossySource { get; set; }
        public string Warning { get; set; }

        public ConvertibleDocument()
        {
            Warning = string.Empty;
        }
    }

    internal static class SafeDocumentConversion
    {
        public static ConvertibleDocument Load(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("Source document was not found.", path);

            string extension = Path.GetExtension(path).ToLowerInvariant();

            if (extension == ".docx")
            {
                EnsureDocxSafe(path);
                return Text(DocxReader.Read(path));
            }

            if (extension == ".hwpx")
            {
                EnsureHwpxSafe(path);
                return Text(HwpxReader.Read(path));
            }

            if (extension == ".odt")
            {
                EnsureOdtSafe(path);
                return Text(OdtReader.Read(path));
            }

            if (extension == ".hwp")
            {
                HwpReadResult result = HwpReader.Read(path);
                ConvertibleDocument document = Text(result.Document);
                document.IsLossySource = true;
                document.Warning =
                    "HWP conversion currently uses the read-only plain-text extraction path. " +
                    "Formatting, tables, images and controls are not preserved in the converted output.";
                return document;
            }

            if (extension == ".xlsx")
            {
                EnsureXlsxSafe(path);
                return Spreadsheet(XlsxReader.Read(path));
            }

            if (extension == ".ods")
            {
                EnsureOdsSafe(path);
                return Spreadsheet(OdsReader.Read(path));
            }

            if (extension == ".pptx")
            {
                PptxEditableLoadResult load = PptxEditableReader.Read(path);
                if (load == null || !load.CanRoundTripSafely || load.Document == null)
                {
                    throw new InvalidOperationException(
                        load == null || string.IsNullOrEmpty(load.Warning)
                            ? "This PPTX cannot be converted through the editable model without known content loss. Open it in Viewer mode instead."
                            : load.Warning);
                }

                return Presentation(load.Document);
            }

            if (extension == ".odp")
            {
                EnsureOdpSafe(path);
                return Presentation(OdpReader.Read(path));
            }

            throw new NotSupportedException(
                "Cross-format conversion is not available for this file type.");
        }

        public static void SaveAs(
            ConvertibleDocument document,
            string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Destination path is required.", "outputPath");

            string extension = Path.GetExtension(outputPath).ToLowerInvariant();

            if (document.Kind == ConvertibleDocumentKind.Text)
            {
                if (extension == ".docx")
                    DocxWriter.Save(document.TextDocument, outputPath);
                else if (extension == ".hwpx")
                    HwpxWriter.Save(document.TextDocument, outputPath);
                else if (extension == ".odt")
                    OdtWriter.Save(document.TextDocument, outputPath);
                else
                    throw new NotSupportedException("Text documents can currently be converted to DOCX, HWPX or ODT.");
                return;
            }

            if (document.Kind == ConvertibleDocumentKind.Spreadsheet)
            {
                if (extension == ".xlsx")
                    XlsxWriter.Save(document.SpreadsheetDocument, outputPath);
                else if (extension == ".ods")
                    OdsWriter.Save(document.SpreadsheetDocument, outputPath);
                else
                    throw new NotSupportedException("Spreadsheets can currently be converted to XLSX or ODS.");
                return;
            }

            if (document.Kind == ConvertibleDocumentKind.Presentation)
            {
                if (extension == ".pptx")
                    PptxWriter.Save(document.PresentationDocument, outputPath);
                else if (extension == ".odp")
                {
                    ValidateOdpShapeCompatibility(document.PresentationDocument);
                    OdpWriter.Save(document.PresentationDocument, outputPath);
                }
                else
                    throw new NotSupportedException("Presentations can currently be converted to PPTX or ODP.");
                return;
            }

            throw new NotSupportedException("Unknown conversion model.");
        }

        private static ConvertibleDocument Text(TextDocument model)
        {
            ConvertibleDocument result = new ConvertibleDocument();
            result.Kind = ConvertibleDocumentKind.Text;
            result.TextDocument = model;
            return result;
        }

        private static ConvertibleDocument Spreadsheet(SpreadsheetDocument model)
        {
            ConvertibleDocument result = new ConvertibleDocument();
            result.Kind = ConvertibleDocumentKind.Spreadsheet;
            result.SpreadsheetDocument = model;
            return result;
        }

        private static ConvertibleDocument Presentation(PresentationDocument model)
        {
            ConvertibleDocument result = new ConvertibleDocument();
            result.Kind = ConvertibleDocumentKind.Presentation;
            result.PresentationDocument = model;
            return result;
        }

        private static void EnsureDocxSafe(string path)
        {
            DocxEditSafetyResult safety = DocxEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "DOCX safety analysis failed." : safety.Warning);
        }

        private static void EnsureHwpxSafe(string path)
        {
            HwpxEditSafetyResult safety = HwpxEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "HWPX safety analysis failed." : safety.Warning);
        }

        private static void EnsureOdtSafe(string path)
        {
            OdtEditSafetyResult safety = OdtEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "ODT safety analysis failed." : safety.Warning);
        }

        private static void EnsureXlsxSafe(string path)
        {
            XlsxEditSafetyResult safety = XlsxEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "XLSX safety analysis failed." : safety.Warning);
        }

        private static void EnsureOdsSafe(string path)
        {
            OdsEditSafetyResult safety = OdsEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "ODS safety analysis failed." : safety.Warning);
        }

        private static void EnsureOdpSafe(string path)
        {
            OdpEditSafetyResult safety = OdpEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "ODP safety analysis failed." : safety.Warning);
        }

        private static void ValidateOdpShapeCompatibility(PresentationDocument document)
        {
            if (document == null)
                return;

            for (int s = 0; s < document.Slides.Count; s++)
            {
                PresentationSlide slide = document.Slides[s];
                if (slide == null)
                    continue;

                for (int i = 0; i < slide.Shapes.Count; i++)
                {
                    PresentationShape shape = slide.Shapes[i];
                    if (shape == null)
                        continue;

                    if (shape.Kind != PresentationShapeKind.Rectangle &&
                        shape.Kind != PresentationShapeKind.Ellipse)
                    {
                        throw new InvalidOperationException(
                            "PPTX → ODP conversion stopped because slide " +
                            (s + 1).ToString() +
                            " contains a shape not yet mapped exactly to ODF: " +
                            shape.Kind.ToString() + ".");
                    }
                }
            }
        }
    }

    public sealed partial class MainForm : Form
    {
        private void ConvertSupportedDocument()
        {
            using (OpenFileDialog open = new OpenFileDialog())
            {
                open.Filter =
                    "Convertible documents (*.pptx;*.odp;*.docx;*.hwpx;*.hwp;*.odt;*.xlsx;*.ods)|*.pptx;*.odp;*.docx;*.hwpx;*.hwp;*.odt;*.xlsx;*.ods|All files (*.*)|*.*";
                open.CheckFileExists = true;

                if (open.ShowDialog(this) != DialogResult.OK)
                    return;

                ConvertibleDocument document;

                try
                {
                    document = SafeDocumentConversion.Load(open.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        this,
                        ex.Message + "\r\n\r\nThe source file was not modified.",
                        "Conversion is not safe yet",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (document.IsLossySource)
                {
                    DialogResult confirmation = MessageBox.Show(
                        this,
                        document.Warning +
                        "\r\n\r\nContinue and create a new file from the extracted content? The original will remain unchanged.",
                        "Limited conversion",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);

                    if (confirmation != DialogResult.OK)
                        return;
                }

                using (SaveFileDialog save = CreateConversionSaveDialog(document, open.FileName))
                {
                    if (save.ShowDialog(this) != DialogResult.OK)
                        return;

                    string outputPath = ApplySelectedConversionExtension(
                        save.FileName,
                        document.Kind,
                        save.FilterIndex);

                    try
                    {
                        SafeDocumentConversion.SaveAs(document, outputPath);
                        status.Text =
                            "Converted: " + Path.GetFileName(open.FileName) +
                            " → " + Path.GetFileName(outputPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            this,
                            ex.Message + "\r\n\r\nThe source file was not modified.",
                            "Conversion failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static SaveFileDialog CreateConversionSaveDialog(
            ConvertibleDocument document,
            string sourcePath)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.AddExtension = true;

            string baseName = Path.GetFileNameWithoutExtension(sourcePath) + "_converted";

            if (document.Kind == ConvertibleDocumentKind.Text)
            {
                dialog.Filter =
                    "Word Open XML Document (*.docx)|*.docx|HWPX Document (*.hwpx)|*.hwpx|OpenDocument Text (*.odt)|*.odt";
                dialog.DefaultExt = "docx";
                dialog.FileName = baseName + ".docx";
            }
            else if (document.Kind == ConvertibleDocumentKind.Spreadsheet)
            {
                dialog.Filter =
                    "Excel Open XML Workbook (*.xlsx)|*.xlsx|OpenDocument Spreadsheet (*.ods)|*.ods";
                dialog.DefaultExt = "xlsx";
                dialog.FileName = baseName + ".xlsx";
            }
            else
            {
                dialog.Filter =
                    "PowerPoint Open XML Presentation (*.pptx)|*.pptx|OpenDocument Presentation (*.odp)|*.odp";
                dialog.DefaultExt = "pptx";
                dialog.FileName = baseName + ".pptx";
            }

            return dialog;
        }

        private static string ApplySelectedConversionExtension(
            string path,
            ConvertibleDocumentKind kind,
            int filterIndex)
        {
            string extension;

            if (kind == ConvertibleDocumentKind.Text)
                extension = filterIndex == 2 ? ".hwpx" : filterIndex == 3 ? ".odt" : ".docx";
            else if (kind == ConvertibleDocumentKind.Spreadsheet)
                extension = filterIndex == 2 ? ".ods" : ".xlsx";
            else
                extension = filterIndex == 2 ? ".odp" : ".pptx";

            if (!string.Equals(
                    Path.GetExtension(path),
                    extension,
                    StringComparison.OrdinalIgnoreCase))
            {
                path = Path.ChangeExtension(path, extension.TrimStart('.'));
            }

            return path;
        }
    }
}
