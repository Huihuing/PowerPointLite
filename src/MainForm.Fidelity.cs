using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private System.Windows.Forms.Timer pptxFidelityTimer;
        private string pptxFidelityCheckedSignature;
        private bool pptxFidelityRefreshRunning;

        private void InitializePptxFidelityFallback()
        {
            if (pptxFidelityTimer != null)
                return;

            pptxFidelityTimer = new System.Windows.Forms.Timer();
            pptxFidelityTimer.Interval = 700;
            pptxFidelityTimer.Tick += delegate
            {
                TryApplyPptxAlternateContentFallback();
            };
            pptxFidelityTimer.Start();

            FormClosed += delegate
            {
                if (pptxFidelityTimer != null)
                {
                    pptxFidelityTimer.Stop();
                    pptxFidelityTimer.Dispose();
                    pptxFidelityTimer = null;
                }
            };
        }

        private void TryApplyPptxAlternateContentFallback()
        {
            if (pptxFidelityRefreshRunning ||
                internalSlideShowMode ||
                string.IsNullOrEmpty(currentFile) ||
                !File.Exists(currentFile))
            {
                return;
            }

            string extension = Path.GetExtension(currentFile).ToLowerInvariant();
            if (extension != ".pptx" && extension != ".pptm")
                return;

            if (!string.Equals(
                    activeEngine,
                    "Internal OpenXML",
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string signature = BuildPptxFidelitySignature(currentFile);

            if (string.Equals(
                    pptxFidelityCheckedSignature,
                    signature,
                    StringComparison.Ordinal))
            {
                return;
            }

            pptxFidelityCheckedSignature = signature;

            if (!PptxRenderPreprocessor.NeedsAlternateContentFallback(currentFile))
                return;

            pptxFidelityRefreshRunning = true;
            Cursor previousCursor = Cursor;
            string prepared = null;
            Cursor = Cursors.WaitCursor;

            try
            {
                prepared = PptxRenderPreprocessor.PrepareForRendering(currentFile);
                if (string.Equals(
                        prepared,
                        currentFile,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                string baseCache = !string.IsNullOrEmpty(currentCacheDirectory)
                    ? currentCacheDirectory
                    : BuildCacheDirectory(currentFile);
                string compatibilityCache = Path.Combine(
                    baseCache,
                    "alternate-content");

                try
                {
                    if (Directory.Exists(compatibilityCache))
                        Directory.Delete(compatibilityCache, true);
                }
                catch
                {
                }

                Directory.CreateDirectory(compatibilityCache);

                List<string> slides = InternalPptxRenderer.Render(
                    prepared,
                    compatibilityCache);

                if (slides == null || slides.Count == 0)
                    return;

                int restoreIndex = currentIndex;

                ClearRenderedSlides();
                renderedSlides.AddRange(slides);
                LoadInteractiveMetadata();
                LoadViewerMetadata();
                BuildThumbnails();

                if (renderedSlides.Count > 0)
                {
                    if (restoreIndex < 0)
                        restoreIndex = 0;
                    if (restoreIndex >= renderedSlides.Count)
                        restoreIndex = renderedSlides.Count - 1;

                    ShowSlide(restoreIndex);
                }

                activeEngine = "Internal OpenXML";
                engineLabel.Text = UiLocalization.Text("Engine: Internal OpenXML");
                status.Text = UiLocalization.CurrentLanguage == AppLanguage.Korean
                    ? "Internal OpenXML · Office 호환 fallback 적용 · " +
                      renderedSlides.Count.ToString() + "개 슬라이드"
                    : "Internal OpenXML · Office compatibility fallback applied · " +
                      renderedSlides.Count.ToString() + " slides";
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "AlternateContent compatibility render failed: " +
                    ex.Message);
            }
            finally
            {
                PptxRenderPreprocessor.DeletePreparedFile(prepared, currentFile);
                Cursor = previousCursor;
                pptxFidelityRefreshRunning = false;
            }
        }

        private static string BuildPptxFidelitySignature(string path)
        {
            try
            {
                FileInfo info = new FileInfo(path);
                return info.FullName.ToLowerInvariant() + "|" +
                       info.Length.ToString() + "|" +
                       info.LastWriteTimeUtc.Ticks.ToString();
            }
            catch
            {
                return path ?? string.Empty;
            }
        }
    }
}
