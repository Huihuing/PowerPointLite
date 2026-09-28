using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Shift | Keys.P))
            {
                ExportCurrentPresentationToPdf();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ExportCurrentPresentationToPdf()
        {
            if (renderedSlides.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Open and render a presentation first.",
                    "Export PDF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF Document (*.pdf)|*.pdf";
                dialog.DefaultExt = "pdf";
                dialog.AddExtension = true;
                dialog.FileName =
                    string.IsNullOrEmpty(currentFile)
                        ? "Presentation.pdf"
                        : Path.GetFileNameWithoutExtension(currentFile) + ".pdf";

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                List<PdfRasterPage> pages = new List<PdfRasterPage>();

                try
                {
                    for (int i = 0; i < renderedSlides.Count; i++)
                    {
                        string imagePath = renderedSlides[i];
                        if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                            continue;

                        using (Image source = LoadImageUnlocked(imagePath))
                        {
                            Bitmap bitmap = new Bitmap(
                                Math.Max(1, source.Width),
                                Math.Max(1, source.Height));

                            using (Graphics graphics = Graphics.FromImage(bitmap))
                            {
                                graphics.Clear(Color.White);
                                graphics.DrawImage(
                                    source,
                                    new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                            }

                            float pageWidth = 720f;
                            float pageHeight = Math.Max(
                                72f,
                                pageWidth * bitmap.Height / (float)bitmap.Width);

                            pages.Add(
                                new PdfRasterPage(
                                    bitmap,
                                    pageWidth,
                                    pageHeight));
                        }
                    }

                    if (pages.Count == 0)
                        throw new InvalidOperationException("No rendered slide pages were available for export.");

                    RasterPdfWriter.Write(dialog.FileName, pages);

                    status.Text =
                        "PDF exported: " + Path.GetFileName(dialog.FileName) +
                        "  •  raster export / no source font-file embedding";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        this,
                        ex.Message,
                        "PDF export failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    for (int i = 0; i < pages.Count; i++)
                    {
                        if (pages[i] != null)
                            pages[i].Dispose();
                    }
                }
            }
        }
    }
}
