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

        private string slideTimingSourceFile;
        private int timingObservedSlide = -1;
        private DateTime timingSlideStartedAt = DateTime.MinValue;
        private bool timingAdvanceInProgress;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            slideTimingTimer =
                new System.Windows.Forms.Timer();

            slideTimingTimer.Interval = 100;
            slideTimingTimer.Tick += delegate
            {
                PollSlideTiming();
            };
            slideTimingTimer.Start();
        }

        private void PollSlideTiming()
        {
            if (timingAdvanceInProgress)
                return;

            RefreshSlideTimingMetadataIfNeeded();

            if (!internalSlideShowMode ||
                currentIndex < 0)
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

            if (currentIndex >= slideAdvanceSpecs.Count)
                return;

            InternalPptxRenderer.SlideAdvanceSpec spec =
                slideAdvanceSpecs[currentIndex];

            if (spec == null ||
                spec.AutoAdvanceMs < 0)
            {
                return;
            }

            double elapsedMs =
                (DateTime.Now - timingSlideStartedAt)
                .TotalMilliseconds;

            if (elapsedMs < spec.AutoAdvanceMs)
                return;

            timingAdvanceInProgress = true;

            try
            {
                timingSlideStartedAt = DateTime.Now;
                Next();
            }
            finally
            {
                timingAdvanceInProgress = false;
            }
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
            slideAdvanceSpecs =
                new List<InternalPptxRenderer.SlideAdvanceSpec>();

            timingObservedSlide = -1;
            timingSlideStartedAt = DateTime.MinValue;

            if (string.IsNullOrEmpty(file) ||
                !File.Exists(file))
            {
                return;
            }

            string extension =
                Path.GetExtension(file).ToLowerInvariant();

            if (extension != ".pptx" &&
                extension != ".pptm")
            {
                return;
            }

            try
            {
                slideAdvanceSpecs =
                    InternalPptxRenderer.ReadSlideAdvanceSpecs(file);
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Slide timing metadata: " + ex.Message);
            }
        }

        private bool CurrentSlideAllowsMouseAdvance()
        {
            RefreshSlideTimingMetadataIfNeeded();

            if (currentIndex < 0 ||
                currentIndex >= slideAdvanceSpecs.Count)
            {
                return true;
            }

            InternalPptxRenderer.SlideAdvanceSpec spec =
                slideAdvanceSpecs[currentIndex];

            return spec == null || spec.AdvanceOnClick;
        }
    }
}
