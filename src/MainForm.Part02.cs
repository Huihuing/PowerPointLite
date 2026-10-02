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
        private void LoadPresentation(string file)
        {
            if (!File.Exists(file))
                return;

            DetectBackends();

            Cursor = Cursors.WaitCursor;
            Enabled = false;

            try
            {
                ClearRenderedSlides();

                currentFile = Path.GetFullPath(file);
                Text = "PowerPointLite 1.3 - " + Path.GetFileName(file);

                string cacheDir = BuildCacheDirectory(currentFile);
                Directory.CreateDirectory(cacheDir);
                currentCacheDirectory = cacheDir;

                List<string> slides = null;
                List<string> failures = new List<string>();

                // Highest fidelity path when Microsoft PowerPoint is installed.
                if (Type.GetTypeFromProgID("PowerPoint.Application") != null)
                {
                    try
                    {
                        status.Text = UiLocalization.Text("Rendering with Microsoft PowerPoint...");
                        Application.DoEvents();

                        slides = RenderWithPowerPoint(currentFile, cacheDir);
                        activeEngine = "Microsoft PowerPoint";
                    }
                    catch (Exception ex)
                    {
                        failures.Add("PowerPoint: " + ex.Message);
                    }
                }

                string sourceExtension = Path.GetExtension(currentFile).ToLowerInvariant();

                // LibreOffice is an optional, separately installed application.
                // When present, use its Impress renderer before the approximation
                // renderer because it usually understands a wider PPTX surface.
                if ((slides == null || slides.Count == 0) &&
                    !string.IsNullOrEmpty(libreOfficePath))
                {
                    try
                    {
                        status.Text = UiLocalization.Text("Rendering with LibreOffice Impress...");
                        Application.DoEvents();

                        slides = RenderWithLibreOffice(currentFile, cacheDir, libreOfficePath);
                        activeEngine = "LibreOffice";
                    }
                    catch (Exception ex)
                    {
                        failures.Add("LibreOffice: " + ex.Message);
                    }
                }

                // Self-contained fallback. This remains the required path when
                // neither PowerPoint nor LibreOffice is available.
                if (slides == null || slides.Count == 0)
                {
                    if (sourceExtension == ".pptx" || sourceExtension == ".pptm")
                    {
                        status.Text = UiLocalization.Text("Reading PPTX directly with Internal OpenXML...");
                        Application.DoEvents();

                        try
                        {
                            slides = InternalPptxRenderer.Render(currentFile, cacheDir);
                            activeEngine = "Internal OpenXML";
                        }
                        catch (Exception ex)
                        {
                            failures.Add("Internal OpenXML: " + ex.Message);
                        }
                    }
                }

                if (slides == null || slides.Count == 0)
                {
                    string detail = failures.Count > 0
                        ? "\r\n\r\n" + string.Join("\r\n\r\n", failures.ToArray())
                        : "";

                    throw new InvalidOperationException(
                        UiLocalization.Text("The presentation could not be rendered.") + detail);
                }

                renderedSlides.AddRange(slides);
                LoadInteractiveMetadata();
                LoadViewerMetadata();
                RecentFileStore.Add(currentFile);
                BuildThumbnails();

                if (renderedSlides.Count > 0)
                    ShowSlide(0);

                engineLabel.Text =
                    activeEngine == "Microsoft PowerPoint" ? "Engine: PowerPoint Native" :
                    activeEngine == "LibreOffice" ? "Engine: LibreOffice Impress" :
                    "Engine: Internal OpenXML";
                engineLabel.Text = UiLocalization.Text(engineLabel.Text);

                status.Text = activeEngine + " | " + renderedSlides.Count +
                    (UiLocalization.IsKorean ? "개 슬라이드" : " slides");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    UiLocalization.Text("Open failed"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                status.Text = UiLocalization.Text("Open failed");
            }
            finally
            {
                Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        private static List<string> RenderWithPowerPoint(string file, string cacheDir)
        {
            Type ppType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (ppType == null)
                throw new InvalidOperationException("Microsoft PowerPoint was not found.");

            object app = null;
            object presentations = null;
            object presentation = null;
            object slides = null;
            object pageSetup = null;

            List<string> result = new List<string>();

            try
            {
                app = Activator.CreateInstance(ppType);
                dynamic dApp = app;
                presentations = dApp.Presentations;
                dynamic dPresentations = presentations;

                presentation = dPresentations.Open(Path.GetFullPath(file), -1, 0, 0);

                dynamic dPresentation = presentation;
                slides = dPresentation.Slides;
                pageSetup = dPresentation.PageSetup;

                int count = (int)((dynamic)slides).Count;
                double sw = Convert.ToDouble(((dynamic)pageSetup).SlideWidth);
                double sh = Convert.ToDouble(((dynamic)pageSetup).SlideHeight);

                int targetW = 1920;
                int targetH = Math.Max(1, (int)Math.Round(targetW * (sh / sw)));

                for (int i = 1; i <= count; i++)
                {
                    string png = Path.Combine(cacheDir, "pp_" + i.ToString("D4") + ".png");

                    if (!File.Exists(png))
                    {
                        object slide = null;
                        try
                        {
                            slide = ((dynamic)slides).Item(i);
                            ((dynamic)slide).Export(png, "PNG", targetW, targetH);
                        }
                        finally
                        {
                            ReleaseCom(slide);
                        }
                    }

                    if (File.Exists(png))
                        result.Add(png);
                }

                return result;
            }
            finally
            {
                try { if (presentation != null) ((dynamic)presentation).Close(); } catch { }
                try { if (app != null) ((dynamic)app).Quit(); } catch { }

                ReleaseCom(pageSetup);
                ReleaseCom(slides);
                ReleaseCom(presentation);
                ReleaseCom(presentations);
                ReleaseCom(app);
            }
        }

        private static List<string> RenderWithLibreOffice(string originalFile, string cacheDir, string soffice)
        {
            if (string.IsNullOrEmpty(soffice) || !File.Exists(soffice))
                throw new InvalidOperationException("LibreOffice executable was not found.");

            string loDir = Path.Combine(cacheDir, "lo");
            Directory.CreateDirectory(loDir);

            string inputFile = originalFile;
            string extension = Path.GetExtension(originalFile).ToLowerInvariant();

            // Legacy binary PPT is converted only inside the cache directory.
            // The user's original file is never modified.
            if (extension == ".ppt")
            {
                string converted = Path.Combine(
                    loDir,
                    Path.GetFileNameWithoutExtension(originalFile) + ".pptx");

                if (!File.Exists(converted))
                {
                    string convertArgs =
                        "--headless --nologo --nodefault --nofirststartwizard " +
                        "--convert-to pptx --outdir " + Quote(loDir) + " " + Quote(originalFile);
                    RunProcess(soffice, convertArgs, 120000);
                }

                if (!File.Exists(converted))
                    throw new InvalidOperationException("LibreOffice did not create a temporary PPTX for the legacy PPT file.");

                inputFile = converted;
            }

            InternalPptxRenderer.PresentationInfo info =
                InternalPptxRenderer.ReadPresentationInfo(inputFile);

            if (info.SlideParts.Count == 0)
                throw new InvalidOperationException("No slides were found.");

            int targetWidth = 1920;
            int targetHeight = info.WidthEmu > 0
                ? Math.Max(1, (int)Math.Round(targetWidth * (info.HeightEmu / (double)info.WidthEmu)))
                : 1080;

            List<string> result = new List<string>();
            string expectedOutput = Path.Combine(
                loDir,
                Path.GetFileNameWithoutExtension(inputFile) + ".png");

            for (int slideIndex = 0; slideIndex < info.SlideParts.Count; slideIndex++)
            {
                string target = Path.Combine(
                    loDir,
                    "lo_" + (slideIndex + 1).ToString("D4") + ".png");

                if (!File.Exists(target))
                {
                    TryDeleteFile(expectedOutput);

                    // LibreOffice's graphic export filter exposes PageNumber,
                    // PixelWidth, PixelHeight and AntiAliasing. PageNumber is
                    // normally one-based for Impress export; if a particular
                    // installation does not produce output, retry zero-based.
                    bool exported = TryExportLibreOfficeSlide(
                        soffice,
                        inputFile,
                        loDir,
                        expectedOutput,
                        slideIndex + 1,
                        targetWidth,
                        targetHeight);

                    if (!exported)
                    {
                        TryDeleteFile(expectedOutput);
                        exported = TryExportLibreOfficeSlide(
                            soffice,
                            inputFile,
                            loDir,
                            expectedOutput,
                            slideIndex,
                            targetWidth,
                            targetHeight);
                    }

                    if (!exported || !File.Exists(expectedOutput))
                    {
                        throw new InvalidOperationException(
                            "LibreOffice did not export slide " +
                            (slideIndex + 1).ToString() + " as PNG.");
                    }

                    File.Copy(expectedOutput, target, true);
                    TryDeleteFile(expectedOutput);
                }

                if (File.Exists(target))
                    result.Add(target);
            }

            return result;
        }

        private static bool TryExportLibreOfficeSlide(
            string soffice,
            string inputFile,
            string outputDirectory,
            string expectedOutput,
            int pageNumber,
            int pixelWidth,
            int pixelHeight)
        {
            try
            {
                string filter =
                    "png:impress_png_Export:" +
                    "{\"PixelWidth\":{\"type\":\"long\",\"value\":\"" + pixelWidth.ToString() + "\"}," +
                    "\"PixelHeight\":{\"type\":\"long\",\"value\":\"" + pixelHeight.ToString() + "\"}," +
                    "\"PageNumber\":{\"type\":\"long\",\"value\":\"" + pageNumber.ToString() + "\"}," +
                    "\"AntiAliasing\":{\"type\":\"boolean\",\"value\":\"true\"}}";

                string args =
                    "--headless --nologo --nodefault --nofirststartwizard " +
                    "--convert-to " + Quote(filter) + " --outdir " +
                    Quote(outputDirectory) + " " + Quote(inputFile);

                RunProcess(soffice, args, 120000);
                return File.Exists(expectedOutput);
            }
            catch
            {
                return false;
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
            }
        }

        private static void RunProcess(string exe, string arguments, int timeoutMs)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = arguments,
                WorkingDirectory = Path.GetDirectoryName(exe),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (Process p = Process.Start(psi))
            {
                string stdout = p.StandardOutput.ReadToEnd();
                string stderr = p.StandardError.ReadToEnd();

                if (!p.WaitForExit(timeoutMs))
                {
                    try { p.Kill(); } catch { }
                    throw new TimeoutException("LibreOffice timed out.");
                }

                if (p.ExitCode != 0)
                    throw new InvalidOperationException("LibreOffice failed: " + stderr + stdout);
            }
        }

        private static string Quote(string value)
        {
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private void StartSlideshow(bool fromCurrent)
        {
            if (string.IsNullOrEmpty(currentFile) || !File.Exists(currentFile))
                return;

            int requestedIndex = fromCurrent
                ? Math.Max(0, currentIndex)
                : FindVisibleSlide(0, 1);

            if (requestedIndex < 0)
                requestedIndex = 0;

            Type ppType = Type.GetTypeFromProgID("PowerPoint.Application");

            if (ppType != null)
            {
                try
                {
                    dynamic app = Activator.CreateInstance(ppType);
                    app.Visible = -1;

                    dynamic pres = app.Presentations.Open(Path.GetFullPath(currentFile), -1, 0, 0);
                    dynamic settings = pres.SlideShowSettings;

                    settings.ShowType = 1;
                    settings.ShowWithAnimation = -1;
                    settings.ShowWithNarration = -1;
                    settings.ShowPresenterView = 0;

                    dynamic window = settings.Run();

                    if (fromCurrent && requestedIndex > 0)
                    {
                        try { window.View.GotoSlide(requestedIndex + 1, -1); } catch { }
                    }

                    return;
                }
                catch { }
            }

            if (!string.IsNullOrEmpty(libreOfficePath))
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = libreOfficePath,
                        UseShellExecute = true,
                        Arguments = "--show=" +
                            (fromCurrent ? Math.Max(1, requestedIndex + 1).ToString() : "1") +
                            " " + Quote(currentFile)
                    };

                    Process.Start(psi);
                    return;
                }
                catch { }
            }

            internalSlideShowMode = true;
            animationRevealCount = 0;
            slideNumberBuffer = "";

            if (currentIndex != requestedIndex)
                ShowSlide(requestedIndex);

            if (!fullscreen)
                ToggleFullscreen();

            PrepareInternalAnimationState();
        }

        private void PrepareInternalAnimationState()
        {
            if (!internalSlideShowMode || currentIndex < 0)
                return;

            animationRevealCount = 0;

            if (currentIndex < animationSteps.Count &&
                animationSteps[currentIndex] != null &&
                animationSteps[currentIndex].Count > 0 &&
                !string.IsNullOrEmpty(currentCacheDirectory))
            {
                try
                {
                    string frame = InternalPptxRenderer.RenderAnimationState(
                        currentFile,
                        currentCacheDirectory,
                        currentIndex,
                        animationRevealCount);

                    if (!string.IsNullOrEmpty(frame) && File.Exists(frame))
                        ReplaceViewerImage(LoadImageUnlocked(frame));
                }
                catch (Exception ex)
                {
                    CrashReporter.WriteLine("Animation start: " + ex.Message);
                }
            }

            if (fitMode)
                ApplyFit();
            else
                ApplyZoom();
        }

        private bool TryAdvanceInternalAnimation(bool forward)
        {
            if (!internalSlideShowMode || currentIndex < 0 || currentIndex >= animationSteps.Count)
                return false;

            List<string> steps = animationSteps[currentIndex];
            if (steps == null || steps.Count == 0)
                return false;

            if (forward)
            {
                if (animationRevealCount >= steps.Count)
                    return false;

                animationRevealCount++;
            }
            else
            {
                if (animationRevealCount <= 0)
                    return false;

                animationRevealCount--;
            }

            try
            {
                string frame = InternalPptxRenderer.RenderAnimationState(
                    currentFile,
                    currentCacheDirectory,
                    currentIndex,
                    animationRevealCount);

                if (!string.IsNullOrEmpty(frame) && File.Exists(frame))
                {
                    Image next = LoadImageUnlocked(frame);
                    AnimateBetweenImages(viewer.Image, next, "fade", 160, false);
                    return true;
                }
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Animation step: " + ex.Message);
            }

            return false;
        }

        private void ReplaceViewerImage(Image next)
        {
            Image old = viewer.Image;
            viewer.Image = next;
            if (old != null && !object.ReferenceEquals(old, next))
                old.Dispose();
        }

        private void ShowSlideWithTransition(int index, bool forward)
        {
            if (transitionAnimating || index < 0 || index >= renderedSlides.Count)
                return;

            Image next = LoadImageUnlocked(renderedSlides[index]);
            string kind = "cut";
            int duration = 220;

            if (index < transitionSpecs.Count && transitionSpecs[index] != null)
            {
                kind = transitionSpecs[index].Kind;
                duration = transitionSpecs[index].DurationMs;
            }

            transitionAnimating = true;

            try
            {
                AnimateBetweenImages(viewer.Image, next, kind, duration, forward);
                currentIndex = index;
                animationRevealCount = 0;

                if (internalSlideShowMode)
                    PrepareInternalAnimationState();

                UpdateStatus();
            }
            finally
            {
                transitionAnimating = false;
            }
        }
    }
}
