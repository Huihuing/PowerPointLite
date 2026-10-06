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
        private void AnimateBetweenImages(
            Image oldImage,
            Image nextImage,
            string kind,
            int durationMs,
            bool forward)
        {
            AnimateBetweenImages(
                oldImage,
                nextImage,
                kind,
                durationMs,
                forward,
                "");
        }

        private void AnimateBetweenImages(
            Image oldImage,
            Image nextImage,
            string kind,
            int durationMs,
            bool forward,
            string direction)
        {
            if (nextImage == null)
                return;

            if (oldImage == null || durationMs <= 0 || kind == "cut")
            {
                ReplaceViewerImage(nextImage);
                if (fitMode) ApplyFit(); else ApplyZoom();
                return;
            }

            Bitmap oldCopy = new Bitmap(oldImage);
            Bitmap nextCopy = new Bitmap(nextImage);
            nextImage.Dispose();

            int width = Math.Max(1, nextCopy.Width);
            int height = Math.Max(1, nextCopy.Height);
            int frames = Math.Max(5, Math.Min(18, durationMs / 24));
            int delay = Math.Max(1, durationMs / frames);

            try
            {
                for (int i = 1; i <= frames; i++)
                {
                    float t = i / (float)frames;
                    Bitmap frame = new Bitmap(width, height, PixelFormat.Format32bppArgb);

                    using (Graphics g = Graphics.FromImage(frame))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(oldCopy, new Rectangle(0, 0, width, height));

                        if (kind == "wipe")
                        {
                            string dir = NormalizeTransitionDirection(
                                direction,
                                forward);

                            if (dir == "u" || dir == "d")
                            {
                                int reveal = Math.Max(
                                    1,
                                    (int)Math.Round(height * t));

                                Rectangle src = dir == "d"
                                    ? new Rectangle(0, 0, width, reveal)
                                    : new Rectangle(0, height - reveal, width, reveal);

                                g.DrawImage(
                                    nextCopy,
                                    src,
                                    src,
                                    GraphicsUnit.Pixel);
                            }
                            else
                            {
                                int reveal = Math.Max(
                                    1,
                                    (int)Math.Round(width * t));

                                Rectangle src = dir == "r"
                                    ? new Rectangle(0, 0, reveal, height)
                                    : new Rectangle(width - reveal, 0, reveal, height);

                                g.DrawImage(
                                    nextCopy,
                                    src,
                                    src,
                                    GraphicsUnit.Pixel);
                            }
                        }
                        else if (kind == "push")
                        {
                            string dir = NormalizeTransitionDirection(
                                direction,
                                forward);

                            g.Clear(Color.Black);

                            if (dir == "u" || dir == "d")
                            {
                                int shift = (int)Math.Round(height * t);

                                if (dir == "u")
                                {
                                    g.DrawImage(
                                        oldCopy,
                                        new Rectangle(0, -shift, width, height));
                                    g.DrawImage(
                                        nextCopy,
                                        new Rectangle(0, height - shift, width, height));
                                }
                                else
                                {
                                    g.DrawImage(
                                        oldCopy,
                                        new Rectangle(0, shift, width, height));
                                    g.DrawImage(
                                        nextCopy,
                                        new Rectangle(0, -height + shift, width, height));
                                }
                            }
                            else
                            {
                                int shift = (int)Math.Round(width * t);

                                if (dir == "l")
                                {
                                    g.DrawImage(
                                        oldCopy,
                                        new Rectangle(-shift, 0, width, height));
                                    g.DrawImage(
                                        nextCopy,
                                        new Rectangle(width - shift, 0, width, height));
                                }
                                else
                                {
                                    g.DrawImage(
                                        oldCopy,
                                        new Rectangle(shift, 0, width, height));
                                    g.DrawImage(
                                        nextCopy,
                                        new Rectangle(-width + shift, 0, width, height));
                                }
                            }
                        }
                        else
                        {
                            ColorMatrix matrix = new ColorMatrix();
                            matrix.Matrix33 = t;

                            using (ImageAttributes attrs = new ImageAttributes())
                            {
                                attrs.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                                g.DrawImage(
                                    nextCopy,
                                    new Rectangle(0, 0, width, height),
                                    0, 0, nextCopy.Width, nextCopy.Height,
                                    GraphicsUnit.Pixel,
                                    attrs);
                            }
                        }
                    }

                    ReplaceViewerImage(frame);
                    if (fitMode) ApplyFit(); else ApplyZoom();
                    Application.DoEvents();
                    Thread.Sleep(delay);
                }

                ReplaceViewerImage(new Bitmap(nextCopy));
                if (fitMode) ApplyFit(); else ApplyZoom();
            }
            finally
            {
                oldCopy.Dispose();
                nextCopy.Dispose();
            }
        }

        private static string NormalizeTransitionDirection(
            string direction,
            bool forward)
        {
            if (direction == "l" ||
                direction == "r" ||
                direction == "u" ||
                direction == "d")
            {
                return direction;
            }

            return forward ? "l" : "r";
        }

        private void LoadInteractiveMetadata()
        {
            interactiveRegions = new List<List<InternalPptxRenderer.InteractiveRegion>>();
            transitionSpecs = new List<InternalPptxRenderer.TransitionSpec>();
            animationSteps = new List<List<string>>();
            animationRevealCount = 0;

            if (string.IsNullOrEmpty(currentFile))
                return;

            string ext = Path.GetExtension(currentFile).ToLowerInvariant();
            if (ext != ".pptx" && ext != ".pptm")
                return;

            try
            {
                interactiveRegions = InternalPptxRenderer.ReadInteractiveRegions(
                    currentFile,
                    currentCacheDirectory);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Interactive metadata: " + ex.Message);
            }

            try
            {
                transitionSpecs = InternalPptxRenderer.ReadTransitions(currentFile);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Transition metadata: " + ex.Message);
            }

            try
            {
                animationSteps = InternalPptxRenderer.ReadAnimationSteps(currentFile);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Animation metadata: " + ex.Message);
            }
        }

        private void LoadViewerMetadata()
        {
            speakerNotes = new List<string>();
            slideTexts = new List<string>();
            slideTitles = new List<string>();
            hiddenSlides = new List<bool>();

            if (string.IsNullOrEmpty(currentFile))
                return;

            string ext = Path.GetExtension(currentFile).ToLowerInvariant();
            if (ext != ".pptx" && ext != ".pptm")
            {
                for (int i = 0; i < renderedSlides.Count; i++)
                {
                    speakerNotes.Add("");
                    slideTexts.Add("");
                    slideTitles.Add("Slide " + (i + 1).ToString());
                    hiddenSlides.Add(false);
                }
                return;
            }

            try
            {
                InternalPptxRenderer.ViewerMetadata meta =
                    InternalPptxRenderer.ReadViewerMetadata(currentFile);

                speakerNotes = meta.SpeakerNotes;
                slideTexts = meta.SlideTexts;
                slideTitles = meta.SlideTitles;
                hiddenSlides = meta.HiddenSlides;
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Viewer metadata: " + ex.Message);
            }

            while (speakerNotes.Count < renderedSlides.Count) speakerNotes.Add("");
            while (slideTexts.Count < renderedSlides.Count) slideTexts.Add("");
            while (slideTitles.Count < renderedSlides.Count)
                slideTitles.Add("Slide " + (slideTitles.Count + 1).ToString());
            while (hiddenSlides.Count < renderedSlides.Count) hiddenSlides.Add(false);
        }

        private InternalPptxRenderer.InteractiveRegion HitTestInteractiveRegion(Point point)
        {
            if (currentIndex < 0 ||
                currentIndex >= interactiveRegions.Count ||
                viewer.ClientSize.Width <= 0 ||
                viewer.ClientSize.Height <= 0)
            {
                return null;
            }

            float nx = point.X / (float)viewer.ClientSize.Width;
            float ny = point.Y / (float)viewer.ClientSize.Height;

            List<InternalPptxRenderer.InteractiveRegion> regions = interactiveRegions[currentIndex];

            for (int i = regions.Count - 1; i >= 0; i--)
            {
                if (regions[i].Bounds.Contains(nx, ny))
                    return regions[i];
            }

            return null;
        }

        private void OnViewerMouseMove(object sender, MouseEventArgs e)
        {
            InternalPptxRenderer.InteractiveRegion region = HitTestInteractiveRegion(e.Location);

            if (region == null)
            {
                viewer.Cursor = Cursors.Default;
                interactiveToolTip.SetToolTip(viewer, null);
                return;
            }

            viewer.Cursor = Cursors.Hand;
            interactiveToolTip.SetToolTip(
                viewer,
                string.IsNullOrEmpty(region.Tooltip) ? region.Target : region.Tooltip);
        }

        private void OnViewerMouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            if (internalSlideShowMode &&
                pointerTool != PointerTool.Arrow)
            {
                return;
            }

            InternalPptxRenderer.InteractiveRegion region = HitTestInteractiveRegion(e.Location);
            if (region == null)
            {
                if (internalSlideShowMode &&
                    CurrentSlideAllowsMouseAdvance())
                {
                    Next();
                }

                return;
            }

            try
            {
                if (region.Kind == "slide" && region.SlideIndex >= 0)
                {
                    if (internalSlideShowMode && fullscreen)
                        ShowSlideWithTransition(region.SlideIndex, region.SlideIndex >= currentIndex);
                    else
                        ShowSlide(region.SlideIndex);
                    return;
                }

                if (region.Kind == "endshow")
                {
                    if (fullscreen) ToggleFullscreen();
                    return;
                }

                if (region.Kind == "media" &&
                    !string.IsNullOrEmpty(region.Target))
                {
                    if (PortableMediaPlayerForm.TryShow(
                            this,
                            region.Target))
                    {
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(region.Target))
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = region.Target,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Open link/media failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void BuildThumbnails()
        {
            thumbnails.SuspendLayout();
            thumbnails.Controls.Clear();

            for (int i = 0; i < renderedSlides.Count; i++)
            {
                int index = i;

                Panel card = new Panel
                {
                    Width = 220,
                    Height = 155,
                    Margin = new Padding(5),
                    BackColor = Color.FromArgb(38, 42, 49),
                    Cursor = Cursors.Hand
                };

                PictureBox pic = new PictureBox
                {
                    Left = 8,
                    Top = 8,
                    Width = 204,
                    Height = 115,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Image = LoadImageUnlocked(renderedSlides[i]),
                    Cursor = Cursors.Hand
                };

                Label label = new Label
                {
                    Left = 8,
                    Top = 128,
                    Width = 204,
                    Height = 20,
                    TextAlign = ContentAlignment.MiddleCenter,
                    ForeColor = (i < hiddenSlides.Count && hiddenSlides[i])
                        ? Color.DarkGray
                        : Color.Gainsboro,
                    Text = (i < hiddenSlides.Count && hiddenSlides[i])
                        ? "Slide " + (i + 1).ToString() + "  [Hidden]"
                        : "Slide " + (i + 1).ToString(),
                    Cursor = Cursors.Hand
                };

                string titleText = i < slideTitles.Count ? slideTitles[i] : "";
                if (!string.IsNullOrWhiteSpace(titleText))
                    interactiveToolTip.SetToolTip(card, titleText);

                EventHandler click = delegate { ShowSlide(index); };
                card.Click += click;
                pic.Click += click;
                label.Click += click;

                card.Controls.Add(pic);
                card.Controls.Add(label);
                thumbnails.Controls.Add(card);
            }

            thumbnails.ResumeLayout();
        }

        private void ShowSlide(int index)
        {
            if (index < 0 || index >= renderedSlides.Count)
                return;

            currentIndex = index;
            animationRevealCount = 0;
            slideNumberBuffer = "";
            RestoreBlankScreen();

            ReplaceViewerImage(LoadImageUnlocked(renderedSlides[index]));
            UpdateNotesBox();

            if (fitMode) ApplyFit();
            else ApplyZoom();

            UpdateStatus();
            SyncPresenterView();
        }

        private void UpdateStatus()
        {
            if (currentIndex < 0)
                return;

            string hidden = currentIndex < hiddenSlides.Count && hiddenSlides[currentIndex]
                ? " | Hidden"
                : "";

            string notes = currentIndex < speakerNotes.Count &&
                           !string.IsNullOrWhiteSpace(speakerNotes[currentIndex])
                ? " | Notes"
                : "";

            status.Text = string.Format(
                "{0} | {1} / {2} | {3}%{4}{5}",
                activeEngine,
                currentIndex + 1,
                renderedSlides.Count,
                (int)Math.Round(zoom * 100),
                hidden,
                notes);
        }

        private static Image LoadImageUnlocked(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (Image temp = Image.FromStream(fs))
            {
                return new Bitmap(temp);
            }
        }

        private void ApplyFit()
        {
            if (viewer.Image == null)
                return;

            int viewportW = Math.Max(
                100,
                Math.Min(
                    mainSplit.Panel2.ClientSize.Width,
                    viewerPanel.ClientSize.Width) - 24);

            int viewportH = Math.Max(
                100,
                Math.Min(
                    mainSplit.Panel2.ClientSize.Height,
                    viewerPanel.ClientSize.Height) - 24);

            float sx = (float)viewportW / viewer.Image.Width;
            float sy = (float)viewportH / viewer.Image.Height;

            zoom = Math.Min(sx, sy);
            zoom = Math.Max(0.10f, Math.Min(3.00f, zoom));

            SetTrackZoom();
            ApplyZoom();
            viewerPanel.AutoScrollPosition = Point.Empty;
        }

        private void SetZoom(float value)
        {
            zoom = Math.Max(0.10f, Math.Min(3.00f, value));
            SetTrackZoom();
            ApplyZoom();
        }

        private void SetTrackZoom()
        {
            int z = (int)Math.Round(zoom * 100);
            z = Math.Max(zoomTrack.Minimum, Math.Min(zoomTrack.Maximum, z));
            zoomTrack.Value = z;
        }

        private void ApplyZoom()
        {
            if (viewer.Image == null)
                return;

            viewer.Width = Math.Max(1, (int)Math.Round(viewer.Image.Width * zoom));
            viewer.Height = Math.Max(1, (int)Math.Round(viewer.Image.Height * zoom));

            viewer.Left = Math.Max(0, (viewerPanel.ClientSize.Width - viewer.Width) / 2);
            viewer.Top = Math.Max(0, (viewerPanel.ClientSize.Height - viewer.Height) / 2);

            UpdateStatus();
        }

        private void OnViewerMouseWheel(object sender, MouseEventArgs e)
        {
            Point screenPoint = Cursor.Position;
            Point local = viewerPanel.PointToClient(screenPoint);

            if (!viewerPanel.ClientRectangle.Contains(local))
                return;

            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                fitMode = false;

                if (e.Delta > 0)
                    SetZoom(zoom + 0.10f);
                else if (e.Delta < 0)
                    SetZoom(zoom - 0.10f);
            }
            else
            {
                if (e.Delta < 0)
                    Next();
                else if (e.Delta > 0)
                    Previous();
            }

            HandledMouseEventArgs handled = e as HandledMouseEventArgs;
            if (handled != null)
                handled.Handled = true;
        }

        private void ToggleSidebar()
        {
            mainSplit.Panel1Collapsed = !mainSplit.Panel1Collapsed;

            if (fitMode)
                BeginInvoke((MethodInvoker)ApplyFit);
        }

        private void Previous()
        {
            if (internalSlideShowMode && TryAdvanceInternalAnimation(false))
                return;

            int target = internalSlideShowMode
                ? FindVisibleSlide(currentIndex - 1, -1)
                : currentIndex - 1;

            if (target >= 0)
            {
                if (internalSlideShowMode && fullscreen)
                    ShowSlideWithTransition(target, false);
                else
                    ShowSlide(target);
            }
        }
    }
}
