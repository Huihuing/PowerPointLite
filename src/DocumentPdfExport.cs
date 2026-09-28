using System;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private void ExportSupportedDocumentToPdf()
        {
            using (OpenFileDialog open = new OpenFileDialog())
            {
                open.Filter =
                    "Exportable documents (*.docx;*.xlsx;*.hwpx;*.hwp;*.odt;*.ods;*.odp)|*.docx;*.xlsx;*.hwpx;*.hwp;*.odt;*.ods;*.odp|" +
                    "Text documents (*.docx;*.hwpx;*.hwp;*.odt)|*.docx;*.hwpx;*.hwp;*.odt|" +
                    "Spreadsheets (*.xlsx;*.ods)|*.xlsx;*.ods|" +
                    "OpenDocument Presentation (*.odp)|*.odp|All files (*.*)|*.*";
                open.CheckFileExists = true;
                open.Multiselect = false;

                if (open.ShowDialog(this) != DialogResult.OK)
                    return;

                using (SaveFileDialog save = new SaveFileDialog())
                {
                    save.Filter = "PDF Document (*.pdf)|*.pdf";
                    save.DefaultExt = "pdf";
                    save.AddExtension = true;
                    save.FileName =
                        Path.GetFileNameWithoutExtension(open.FileName) + ".pdf";

                    if (save.ShowDialog(this) != DialogResult.OK)
                        return;

                    try
                    {
                        ExportDocumentPathToPdf(open.FileName, save.FileName);
                        status.Text =
                            "PDF exported: " + Path.GetFileName(save.FileName) +
                            "  •  raster export / source font files are not embedded";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            this,
                            ex.Message +
                            "\r\n\r\nThe source document was not modified.",
                            "PDF export failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static void ExportDocumentPathToPdf(
            string sourcePath,
            string outputPath)
        {
            if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                throw new FileNotFoundException("The source document was not found.", sourcePath);

            string extension = Path.GetExtension(sourcePath).ToLowerInvariant();

            if (extension == ".docx")
            {
                PdfExportService.ExportTextDocument(
                    DocxReader.Read(sourcePath),
                    outputPath);
                return;
            }

            if (extension == ".hwpx")
            {
                PdfExportService.ExportTextDocument(
                    HwpxReader.Read(sourcePath),
                    outputPath);
                return;
            }

            if (extension == ".hwp")
            {
                HwpReadResult result = HwpReader.Read(sourcePath);
                PdfExportService.ExportTextDocument(
                    result.Document,
                    outputPath);
                return;
            }

            if (extension == ".odt")
            {
                PdfExportService.ExportTextDocument(
                    OdtReader.Read(sourcePath),
                    outputPath);
                return;
            }

            if (extension == ".xlsx")
            {
                PdfExportService.ExportSpreadsheet(
                    XlsxReader.Read(sourcePath),
                    outputPath);
                return;
            }

            if (extension == ".ods")
            {
                PdfExportService.ExportSpreadsheet(
                    OdsReader.Read(sourcePath),
                    outputPath);
                return;
            }

            if (extension == ".odp")
            {
                PdfExportService.ExportPresentation(
                    OdpReader.Read(sourcePath),
                    outputPath);
                return;
            }

            throw new NotSupportedException(
                "PDF export is not available for this file type. " +
                "For PPTX, open it in the Viewer and use Ctrl+Shift+P so the existing renderer is used.");
        }
    }
}
