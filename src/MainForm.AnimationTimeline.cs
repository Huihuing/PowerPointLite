using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private List<InternalPptxRenderer.SlideAnimationTimeline> enhancedAnimationTimelines =
            new List<InternalPptxRenderer.SlideAnimationTimeline>();
        private string enhancedAnimationTimelineSourceFile;
        private System.Windows.Forms.Timer enhancedAnimationTimer;
        private bool enhancedAnimationAdvancing;
        private int enhancedAnimationObservedSlide = -1;
        private InternalPptxRenderer.AnimationStepSpec enhancedPendingStep;
        private int enhancedPendingTarget = -1;
        private DateTime enhancedPendingDue = DateTime.MinValue;

        private void EnsureEnhancedAnimationRuntime()
        {
            if (enhancedAnimationTimer != null)
                return;

            enhancedAnimationTimer = new System.Windows.Forms.Timer();
            enhancedAnimationTimer.Interval = 50;
            enhancedAnimationTimer.Tick += delegate
            {
                PollEnhancedAnimationTimeline();
            };
            enhancedAnimationTimer.Start();

            FormClosed += delegate
            {
                try
                {
                    if (enhancedAnimationTimer != null)
                    {
                        enhancedAnimationTimer.Stop();
                        enhancedAnimationTimer.Dispose();
                        enhancedAnimationTimer = null;
                    }
                }
                catch { }
            };
        }

        private void RefreshEnhancedAnimationTimelineMetadata()
        {
            string file = currentFile;
            if (string.Equals(
                    file,
                    enhancedAnimationTimelineSourceFile,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            enhancedAnimationTimelineSourceFile = file;
            enhancedAnimationTimelines =
                new List<InternalPptxRenderer.SlideAnimationTimeline>();
            enhancedAnimationObservedSlide = -1;
            CancelEnhancedPendingAnimation();

            if (string.IsNullOrEmpty(file) || !File.Exists(file))
                return;

            string extension = Path.GetExtension(file).ToLowerInvariant();
            if (extension != ".pptx" && extension != ".pptm")
                return;

            try
            {
                enhancedAnimationTimelines =
                    InternalPptxRenderer.ReadAnimationTimelines(file);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Animation timeline metadata: " + ex.Message);
            }
        }

        private bool PrepareEnhancedAnimationTimelineState()
        {
            EnsureEnhancedAnimationRuntime();
            RefreshEnhancedAnimationTimelineMetadata();

            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();
            if (timeline == null || timeline.Steps.Count == 0)
                return false;

            CancelEnhancedPendingAnimation();
            animationRevealCount = 0;
            enhancedAnimationObservedSlide = currentIndex;

            try
            {
                string frame =
                    InternalPptxRenderer.RenderAnimationTimelineState(
                        currentFile,
                        currentCacheDirectory,
                        currentIndex,
                        0);

                if (!string.IsNullOrEmpty(frame) && File.Exists(frame))
                    ReplaceViewerImage(LoadImageUnlocked(frame));
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Animation timeline start: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool TryAdvanceEnhancedAnimationTimeline(bool forward)
        {
            EnsureEnhancedAnimationRuntime();
            RefreshEnhancedAnimationTimelineMetadata();

            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();
            if (timeline == null || timeline.Steps.Count == 0)
                return false;

            if (enhancedAnimationAdvancing)
                return true;

            if (forward)
            {
                if (animationRevealCount >= timeline.Steps.Count)
                    return false;

                if (enhancedPendingTarget >= 0)
                    return true;

                InternalPptxRenderer.AnimationStepSpec step =
                    timeline.Steps[animationRevealCount];

                if (step != null &&
                    !string.IsNullOrEmpty(
                        step.TriggerShapeId))
                {
                    return true;
                }

                QueueEnhancedAnimationStep(
                    step,
                    animationRevealCount + 1,
                    true);
                return true;
            }

            CancelEnhancedPendingAnimation();

            if (animationRevealCount <= 0)
                return false;

            InternalPptxRenderer.AnimationStepSpec reverseStep =
                timeline.Steps[Math.Max(0, animationRevealCount - 1)];
            ApplyEnhancedAnimationState(
                reverseStep,
                animationRevealCount - 1,
                false);
            return true;
        }

        private bool HasPendingEnhancedSpecificTrigger()
        {
            RefreshEnhancedAnimationTimelineMetadata();

            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();

            if (timeline == null ||
                animationRevealCount < 0 ||
                animationRevealCount >=
                    timeline.Steps.Count)
            {
                return false;
            }

            InternalPptxRenderer.AnimationStepSpec step =
                timeline.Steps[
                    animationRevealCount];

            return step != null &&
                step.RequiresClick &&
                !string.IsNullOrEmpty(
                    step.TriggerShapeId);
        }

        private bool TryAdvanceEnhancedAnimationTrigger(
            Point point)
        {
            if (!internalSlideShowMode ||
                currentIndex < 0 ||
                viewer.ClientSize.Width <= 0 ||
                viewer.ClientSize.Height <= 0)
            {
                return false;
            }

            EnsureEnhancedAnimationRuntime();
            RefreshEnhancedAnimationTimelineMetadata();

            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();

            if (timeline == null ||
                animationRevealCount < 0 ||
                animationRevealCount >=
                    timeline.Steps.Count)
            {
                return false;
            }

            InternalPptxRenderer.AnimationStepSpec step =
                timeline.Steps[
                    animationRevealCount];

            if (step == null ||
                !step.RequiresClick ||
                string.IsNullOrEmpty(
                    step.TriggerShapeId))
            {
                return false;
            }

            if (enhancedAnimationAdvancing ||
                enhancedPendingTarget >= 0)
            {
                return true;
            }

            if (currentIndex >=
                    animationShapeRegions.Count)
            {
                return false;
            }

            float nx;
            float ny;

            if (!TryNormalizeViewerPoint(
                    point,
                    out nx,
                    out ny))
            {
                return false;
            }

            List<InternalPptxRenderer.ShapeRegion> regions =
                animationShapeRegions[
                    currentIndex];

            for (int i =
                     regions.Count - 1;
                 i >= 0;
                 i--)
            {
                InternalPptxRenderer.ShapeRegion region =
                    regions[i];

                if (region == null ||
                    !string.Equals(
                        region.ShapeId,
                        step.TriggerShapeId,
                        StringComparison.OrdinalIgnoreCase) ||
                    !region.Bounds.Contains(
                        nx,
                        ny))
                {
                    continue;
                }

                QueueEnhancedAnimationStep(
                    step,
                    animationRevealCount + 1,
                    true);
                return true;
            }

            return false;
        }

        private void PollEnhancedAnimationTimeline()
        {
            if (!internalSlideShowMode ||
                currentIndex < 0 ||
                enhancedAnimationAdvancing)
            {
                if (!internalSlideShowMode)
                    CancelEnhancedPendingAnimation();
                return;
            }

            RefreshEnhancedAnimationTimelineMetadata();

            if (enhancedAnimationObservedSlide != currentIndex)
            {
                enhancedAnimationObservedSlide = currentIndex;
                CancelEnhancedPendingAnimation();
                return;
            }

            if (enhancedPendingTarget >= 0)
            {
                if (DateTime.Now >= enhancedPendingDue)
                {
                    InternalPptxRenderer.AnimationStepSpec pending =
                        enhancedPendingStep;
                    int target = enhancedPendingTarget;
                    enhancedPendingStep = null;
                    enhancedPendingTarget = -1;
                    enhancedPendingDue = DateTime.MinValue;
                    ApplyEnhancedAnimationState(pending, target, true);
                }
                return;
            }

            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();
            if (timeline == null ||
                animationRevealCount < 0 ||
                animationRevealCount >= timeline.Steps.Count)
            {
                return;
            }

            InternalPptxRenderer.AnimationStepSpec next =
                timeline.Steps[animationRevealCount];
            if (next == null || next.RequiresClick)
                return;

            QueueEnhancedAnimationStep(
                next,
                animationRevealCount + 1,
                true);
        }

        private void QueueEnhancedAnimationStep(
            InternalPptxRenderer.AnimationStepSpec step,
            int targetStep,
            bool forward)
        {
            if (step == null)
                return;

            int delay = forward
                ? Math.Max(0, Math.Min(60000, step.AutoStartDelayMs))
                : 0;

            if (delay <= 0)
            {
                ApplyEnhancedAnimationState(step, targetStep, forward);
                return;
            }

            enhancedPendingStep = step;
            enhancedPendingTarget = targetStep;
            enhancedPendingDue = DateTime.Now.AddMilliseconds(delay);
        }

        private void ApplyEnhancedAnimationState(
            InternalPptxRenderer.AnimationStepSpec step,
            int targetStep,
            bool forward)
        {
            if (step == null || enhancedAnimationAdvancing)
                return;

            enhancedAnimationAdvancing = true;

            try
            {
                int visualDuration = forward
                    ? Math.Max(
                        80,
                        Math.Min(
                            10000,
                            Math.Max(
                                1,
                                step.TotalDurationMs -
                                step.AutoStartDelayMs)))
                    : 120;

                bool playedObjectFrames = false;

                if (forward)
                {
                    int frameCount = Math.Max(
                        5,
                        Math.Min(
                            14,
                            visualDuration / 55));

                    List<string> frames =
                        InternalPptxRenderer.RenderAnimationTimelineStepFrames(
                            currentFile,
                            currentCacheDirectory,
                            currentIndex,
                            Math.Max(0, targetStep - 1),
                            frameCount);

                    if (frames != null && frames.Count > 0)
                    {
                        PlayEnhancedAnimationFrames(
                            frames,
                            visualDuration);
                        playedObjectFrames = true;
                    }
                }

                if (!playedObjectFrames)
                {
                    string frame =
                        InternalPptxRenderer.RenderAnimationTimelineState(
                            currentFile,
                            currentCacheDirectory,
                            currentIndex,
                            targetStep);

                    if (string.IsNullOrEmpty(frame) ||
                        !File.Exists(frame))
                    {
                        return;
                    }

                    Image next = LoadImageUnlocked(frame);
                    string kind = forward
                        ? NormalizeEnhancedAnimationKind(step.VisualKind)
                        : "fade";

                    AnimateBetweenImages(
                        viewer.Image,
                        next,
                        kind,
                        visualDuration,
                        forward);
                }

                animationRevealCount = targetStep;
                enhancedAnimationObservedSlide = currentIndex;
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Animation timeline step: " + ex.Message);
            }
            finally
            {
                enhancedAnimationAdvancing = false;
            }
        }

        private void PlayEnhancedAnimationFrames(
            List<string> frames,
            int totalDurationMs)
        {
            if (frames == null || frames.Count == 0)
                return;

            int delay = Math.Max(
                1,
                totalDurationMs /
                Math.Max(1, frames.Count));

            for (int i = 0; i < frames.Count; i++)
            {
                string path = frames[i];
                if (string.IsNullOrEmpty(path) ||
                    !File.Exists(path))
                {
                    continue;
                }

                ReplaceViewerImage(
                    LoadImageUnlocked(path));

                if (fitMode)
                    ApplyFit();
                else
                    ApplyZoom();

                Application.DoEvents();
                System.Threading.Thread.Sleep(delay);
            }
        }

        private static string NormalizeEnhancedAnimationKind(string kind)
        {
            if (kind == "push") return "push";
            if (kind == "wipe") return "wipe";
            return "fade";
        }

        private InternalPptxRenderer.SlideAnimationTimeline
            CurrentEnhancedAnimationTimeline()
        {
            if (currentIndex < 0 ||
                currentIndex >= enhancedAnimationTimelines.Count)
            {
                return null;
            }

            return enhancedAnimationTimelines[currentIndex];
        }

        private void CancelEnhancedPendingAnimation()
        {
            enhancedPendingStep = null;
            enhancedPendingTarget = -1;
            enhancedPendingDue = DateTime.MinValue;
        }
    }
}
