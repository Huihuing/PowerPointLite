using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace PptxViewer
{
public sealed partial class MainForm : Form
    {
        private void PrintSlidesLayout(PrintLayoutMode mode)
        {
            if (renderedSlides.Count == 0)
                return;

            using (PrintDocument document = new PrintDocument())
            using (PrintDialog dialog = new PrintDialog())
            {
                int index = 0;
                document.DocumentName = string.IsNullOrEmpty(currentFile)
                    ? "Presentation"
                    : Path.GetFileName(currentFile);

                document.PrintPage += delegate(object sender, PrintPageEventArgs e)
                {
                    Rectangle bounds = e.MarginBounds;

                    if (mode == PrintLayoutMode.FullSlide)
                    {
                        DrawPrintSlide(e.Graphics, renderedSlides[index], bounds);
                        index++;
                    }
                    else if (mode == PrintLayoutMode.NotesPage)
                    {
                        int topHeight = (int)(bounds.Height * 0.58);
                        Rectangle slideRect = new Rectangle(bounds.Left, bounds.Top, bounds.Width, topHeight);
                        DrawPrintSlide(e.Graphics, renderedSlides[index], slideRect);

                        string notes = index < speakerNotes.Count ? speakerNotes[index] : "";

                        using (Font titleFont = new Font("Segoe UI", 11f, FontStyle.Bold))
                        using (Font notesFont = new Font("Segoe UI", 10f))
                        using (Brush brush = new SolidBrush(Color.Black))
                        {
                            float y = slideRect.Bottom + 16;
                            e.Graphics.DrawString("Slide " + (index + 1).ToString() + " Notes", titleFont, brush, bounds.Left, y);

                            RectangleF notesRect = new RectangleF(
                                bounds.Left,
                                y + 28,
                                bounds.Width,
                                Math.Max(1, bounds.Bottom - (y + 28)));

                            e.Graphics.DrawString(notes ?? "", notesFont, brush, notesRect);
                        }

                        index++;
                    }
                    else
                    {
                        int perPage =
                            mode == PrintLayoutMode.Handout2 ? 2 :
                            mode == PrintLayoutMode.Handout4 ? 4 :
                            6;

                        int cols = perPage == 2 ? 1 : 2;
                        int rows = (int)Math.Ceiling(perPage / (double)cols);
                        int cellW = bounds.Width / cols;
                        int cellH = bounds.Height / rows;

                        for (int slot = 0; slot < perPage && index < renderedSlides.Count; slot++, index++)
                        {
                            int col = slot % cols;
                            int row = slot / cols;

                            Rectangle cell = new Rectangle(
                                bounds.Left + col * cellW + 8,
                                bounds.Top + row * cellH + 8,
                                cellW - 16,
                                cellH - 28);

                            DrawPrintSlide(e.Graphics, renderedSlides[index], cell);

                            using (Font font = new Font("Segoe UI", 8f))
                            using (Brush brush = new SolidBrush(Color.Black))
                            {
                                e.Graphics.DrawString((index + 1).ToString(), font, brush, cell.Left, cell.Bottom + 2);
                            }
                        }
                    }

                    e.HasMorePages = index < renderedSlides.Count;
                };

                dialog.Document = document;
                dialog.UseEXDialog = true;

                if (dialog.ShowDialog(this) == DialogResult.OK)
                    document.Print();
            }
        }

        private static void DrawPrintSlide(Graphics graphics, string imagePath, Rectangle bounds)
        {
            using (Image image = LoadImageUnlocked(imagePath))
            {
                float scale = Math.Min(bounds.Width / (float)image.Width, bounds.Height / (float)image.Height);
                int w = Math.Max(1, (int)Math.Round(image.Width * scale));
                int h = Math.Max(1, (int)Math.Round(image.Height * scale));
                int x = bounds.Left + (bounds.Width - w) / 2;
                int y = bounds.Top + (bounds.Height - h) / 2;

                graphics.DrawImage(image, new Rectangle(x, y, w, h));
            }
        }

        private void PrintSlides()
        {
            if (renderedSlides.Count == 0)
                return;

            int printIndex = 0;

            using (PrintDocument document = new PrintDocument())
            using (PrintDialog dialog = new PrintDialog())
            {
                document.DocumentName = string.IsNullOrEmpty(currentFile)
                    ? "PowerPointLite"
                    : Path.GetFileName(currentFile);

                document.DefaultPageSettings.Landscape = true;
                dialog.Document = document;
                dialog.UseEXDialog = true;

                document.PrintPage += delegate(object sender, PrintPageEventArgs e)
                {
                    if (printIndex >= renderedSlides.Count)
                    {
                        e.HasMorePages = false;
                        return;
                    }

                    using (Image image = LoadImageUnlocked(renderedSlides[printIndex]))
                    {
                        Rectangle bounds = e.MarginBounds;
                        float sx = bounds.Width / (float)image.Width;
                        float sy = bounds.Height / (float)image.Height;
                        float scale = Math.Min(sx, sy);
                        int w = Math.Max(1, (int)Math.Round(image.Width * scale));
                        int h = Math.Max(1, (int)Math.Round(image.Height * scale));
                        int x = bounds.Left + (bounds.Width - w) / 2;
                        int y = bounds.Top + (bounds.Height - h) / 2;

                        e.Graphics.DrawImage(image, new Rectangle(x, y, w, h));
                    }

                    printIndex++;
                    e.HasMorePages = printIndex < renderedSlides.Count;
                };

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    try { document.Print(); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, ex.Message, "Print failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            if (e.Data == null)
                return;

            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0)
                return;

            string ext = Path.GetExtension(files[0]).ToLowerInvariant();

            if (ext == ".pptx" || ext == ".pptm" || ext == ".ppt")
                LoadPresentation(files[0]);
        }

        private void OnKeyDownMain(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O)
            {
                OpenFile();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.P)
            {
                PrintWithLayout();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.F)
            {
                FindSlidesDialog();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.G)
            {
                GoToSlideDialog();
                e.Handled = true;
                return;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.S)
            {
                ShowSlideSorter();
                e.Handled = true;
                return;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.N)
            {
                ToggleNotes();
                e.Handled = true;
                return;
            }

            if (e.KeyCode == Keys.F5)
            {
                StartSlideshow(e.Shift);
                e.Handled = true;
                return;
            }

            if (internalSlideShowMode)
            {
                if (e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9)
                {
                    slideNumberBuffer += ((int)e.KeyCode - (int)Keys.D0).ToString();
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9)
                {
                    slideNumberBuffer += ((int)e.KeyCode - (int)Keys.NumPad0).ToString();
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.Enter && slideNumberBuffer.Length > 0)
                {
                    int n;
                    if (int.TryParse(slideNumberBuffer, out n) && n >= 1 && n <= renderedSlides.Count)
                        ShowSlideWithTransition(n - 1, n - 1 >= currentIndex);

                    slideNumberBuffer = "";
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.B)
                {
                    ToggleBlankScreen(Color.Black);
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.W)
                {
                    ToggleBlankScreen(Color.White);
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.H)
                {
                    ShowNextHiddenSlide();
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.N || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Right ||
                    e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Space)
                {
                    RestoreBlankScreen();
                    Next();
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.P || e.KeyCode == Keys.Left || e.KeyCode == Keys.Up ||
                    e.KeyCode == Keys.PageUp || e.KeyCode == Keys.Back)
                {
                    RestoreBlankScreen();
                    Previous();
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.Escape)
                {
                    EndInternalSlideshow();
                    e.Handled = true;
                    return;
                }
            }

            if (e.KeyCode == Keys.Left || e.KeyCode == Keys.PageUp)
            {
                Previous();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right || e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Space)
            {
                Next();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Home && renderedSlides.Count > 0)
            {
                ShowSlide(0);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.End && renderedSlides.Count > 0)
            {
                ShowSlide(renderedSlides.Count - 1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F11)
            {
                ToggleFullscreen();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.T)
            {
                ToggleSidebar();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape && fullscreen)
            {
                ToggleFullscreen();
                e.Handled = true;
            }
        }

        private static string BuildCacheDirectory(string file)
        {
            FileInfo fi = new FileInfo(file);

            string keyInput =
                fi.FullName.ToLowerInvariant() + "|" +
                fi.Length.ToString() + "|" +
                fi.LastWriteTimeUtc.Ticks.ToString() + "|v122";

            byte[] raw = Encoding.UTF8.GetBytes(keyInput);
            string hash;

            using (SHA1 sha = SHA1.Create())
            {
                byte[] digest = sha.ComputeHash(raw);
                StringBuilder sb = new StringBuilder();

                for (int i = 0; i < digest.Length; i++)
                    sb.Append(digest[i].ToString("x2"));

                hash = sb.ToString();
            }

            string cacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PptxViewer",
                "Cache");

            return Path.Combine(cacheRoot, hash);
        }

        private static void ReleaseCom(object obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.FinalReleaseComObject(obj);
            }
            catch { }
        }

        private void ClearRenderedSlides()
        {
            if (viewer.Image != null)
            {
                Image img = viewer.Image;
                viewer.Image = null;
                img.Dispose();
            }

            foreach (Control c in thumbnails.Controls)
            {
                Panel p = c as Panel;
                if (p == null) continue;

                foreach (Control child in p.Controls)
                {
                    PictureBox pb = child as PictureBox;
                    if (pb != null && pb.Image != null)
                    {
                        Image img = pb.Image;
                        pb.Image = null;
                        img.Dispose();
                    }
                }
            }

            thumbnails.Controls.Clear();
            renderedSlides.Clear();
            interactiveRegions.Clear();
            transitionSpecs.Clear();
            animationSteps.Clear();
            speakerNotes.Clear();
            slideTexts.Clear();
            slideTitles.Clear();
            hiddenSlides.Clear();
            animationRevealCount = 0;
            blankScreenActive = false;

            if (blankScreenBackup != null)
            {
                blankScreenBackup.Dispose();
                blankScreenBackup = null;
            }

            if (notesBox != null)
                notesBox.Clear();
            internalSlideShowMode = false;
            currentIndex = -1;
        }

        private void OnClosing(object sender, FormClosingEventArgs e)
        {
            try { sidebarHideTimer.Stop(); } catch { }

            try
            {
                if (presenterView != null && !presenterView.IsDisposed)
                    presenterView.Close();
            }
            catch { }

            ClearRenderedSlides();
        }
    }
}
