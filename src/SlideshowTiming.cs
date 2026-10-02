using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private System.Windows.Forms.Timer slideTimingTimer;
        private List<InternalPptxRenderer.SlideAdvanceSpec> slideAdvanceSpecs =
            new List<InternalPptxRenderer.SlideAdvanceSpec>();

        private InternalPptxRenderer.SlideShowSettingsSpec slideShowSettings =
            new InternalPptxRenderer.SlideShowSettingsSpec();

        private string slideTimingSourceFile;
        private int timingObservedSlide = -1;
        private DateTime timingSlideStartedAt = DateTime.MinValue;
        private bool timingAdvanceInProgress;
        private bool slideTimingInitialized;

        private void InitializeSlideShowTiming()
        {
            if (slideTimingInitialized)
                return;

            slideTimingInitialized = true;
            slideTimingTimer = new System.Windows.Forms.Timer();
            slideTimingTimer.Interval = 100;
            slideTimingTimer.Tick += delegate
            {
                PollSlideTiming();
            };
            slideTimingTimer.Start();

            FormClosed += delegate
            {
                try
                {
                    slideTimingTimer.Stop();
                    slideTimingTimer.Dispose();
                }
                catch
                {
                }
            };
        }

        private void PollSlideTiming()
        {
            if (timingAdvanceInProgress)
                return;

            RefreshSlideTimingMetadataIfNeeded();

            if (!internalSlideShowMode || currentIndex < 0)
            {
                timingObservedSlide = -1;
                timingSlideStartedAt = DateTime.MinValue;
                return;
            }

            if (currentIndex != timingObservedSlide)
            {
                timingObservedSlide = currentIndex;
                timingSlideStartedAt = DateTime.Now;
                return;
            }

            if (slideShowSettings != null && !slideShowSettings.UseTimings)
                return;

            if (currentIndex >= slideAdvanceSpecs.Count)
                return;

            InternalPptxRenderer.SlideAdvanceSpec spec = slideAdvanceSpecs[currentIndex];
            if (spec == null || spec.AutoAdvanceMs < 0)
                return;

            double elapsedMs = (DateTime.Now - timingSlideStartedAt).TotalMilliseconds;
            if (elapsedMs < spec.AutoAdvanceMs)
                return;

            timingAdvanceInProgress = true;

            try
            {
                // A slide-level advTm means "leave this slide". Calling Next()
                // here would consume one on-click animation step first, causing
                // recorded slide timings to stall on slides that contain builds.
                // Bypass animation-step navigation only for this timer-driven
                // slide transition; normal mouse/keyboard Next still consumes
                // animation steps before advancing the slide.
                AdvanceSlideFromRecordedTiming();
            }
            finally
            {
                timingSlideStartedAt = DateTime.Now;
                timingAdvanceInProgress = false;
            }
        }

        private void AdvanceSlideFromRecordedTiming()
        {
            if (!internalSlideShowMode || renderedSlides.Count == 0)
                return;

            CancelEnhancedPendingAnimation();

            int target = FindVisibleSlide(currentIndex + 1, 1);
            if (target >= 0 && target < renderedSlides.Count)
            {
                if (fullscreen)
                    ShowSlideWithTransition(target, true);
                else
                    ShowSlide(target);
                return;
            }

            if (InternalSlideShowLoops())
            {
                int first = FindVisibleSlide(0, 1);
                if (first >= 0)
                {
                    if (fullscreen)
                        ShowSlideWithTransition(first, true);
                    else
                        ShowSlide(first);
                }
                return;
            }

            EndInternalSlideshow();
        }

        private void RefreshSlideTimingMetadataIfNeeded()
        {
            string file = currentFile;

            if (string.Equals(
                    file,
                    slideTimingSourceFile,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            slideTimingSourceFile = file;
            slideAdvanceSpecs = new List<InternalPptxRenderer.SlideAdvanceSpec>();
            slideShowSettings = new InternalPptxRenderer.SlideShowSettingsSpec();
            timingObservedSlide = -1;
            timingSlideStartedAt = DateTime.MinValue;

            if (string.IsNullOrEmpty(file) || !File.Exists(file))
                return;

            string extension = Path.GetExtension(file).ToLowerInvariant();
            if (extension != ".pptx" && extension != ".pptm")
                return;

            try
            {
                slideAdvanceSpecs = InternalPptxRenderer.ReadSlideAdvanceSpecs(file);
                slideShowSettings = InternalPptxRenderer.ReadSlideShowSettings(file);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine("Slide timing metadata: " + ex.Message);
            }
        }

        private bool InternalSlideShowLoops()
        {
            RefreshSlideTimingMetadataIfNeeded();
            return slideShowSettings != null && slideShowSettings.Loop;
        }

        private bool CurrentSlideAllowsMouseAdvance()
        {
            // Transition advClick controls whether a click may LEAVE the slide.
            // It must not suppress click-triggered builds on the current slide.
            // Return true while a click animation remains so OnViewerMouseClick
            // can call Next(), which consumes the animation before slide motion.
            if (HasPendingClickAnimationBuild())
                return true;

            RefreshSlideTimingMetadataIfNeeded();

            if (currentIndex < 0 || currentIndex >= slideAdvanceSpecs.Count)
                return true;

            InternalPptxRenderer.SlideAdvanceSpec spec = slideAdvanceSpecs[currentIndex];
            return spec == null || spec.AdvanceOnClick;
        }

        private bool HasPendingClickAnimationBuild()
        {
            RefreshEnhancedAnimationTimelineMetadata();
            InternalPptxRenderer.SlideAnimationTimeline timeline =
                CurrentEnhancedAnimationTimeline();

            if (timeline != null &&
                animationRevealCount >= 0 &&
                animationRevealCount < timeline.Steps.Count)
            {
                InternalPptxRenderer.AnimationStepSpec next =
                    timeline.Steps[animationRevealCount];
                if (next != null && next.RequiresClick)
                    return true;
            }

            if (currentIndex >= 0 &&
                currentIndex < animationSteps.Count &&
                animationSteps[currentIndex] != null &&
                animationRevealCount < animationSteps[currentIndex].Count)
            {
                return true;
            }

            return false;
        }
    }
}
