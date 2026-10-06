using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class AnimationTimingDiagnostics
    {
        private const string PresentationNamespace =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

        public static void CreateAndValidate(string outputDirectory)
        {
            if (string.IsNullOrEmpty(outputDirectory))
                throw new ArgumentException("Output directory is required.", "outputDirectory");

            outputDirectory = Path.GetFullPath(outputDirectory);
            Directory.CreateDirectory(outputDirectory);

            string packagePath = Path.Combine(
                outputDirectory,
                "TEST_ANIMATION_TIMING.pptx");
            string renderDirectory = Path.Combine(
                outputDirectory,
                "render");

            DeleteFile(packagePath);
            DeleteDirectory(renderDirectory);
            Directory.CreateDirectory(renderDirectory);

            PresentationDocument document =
                PresentationDocument.CreateNew("Animation Timing Self Test");
            PresentationSlide slide = document.Slides[0];

            PresentationShape first = slide.AddShape(
                PresentationShapeKind.RoundedRectangle);
            first.Name = "Entrance target";
            first.X = 1371600;
            first.Y = 2743200;
            first.Width = 3200400;
            first.Height = 1371600;
            first.FillColorHex = "5B8CFF";

            PresentationShape second = slide.AddShape(
                PresentationShapeKind.Ellipse);
            second.Name = "Exit target";
            second.X = 7315200;
            second.Y = 2743200;
            second.Width = 1828800;
            second.Height = 1371600;
            second.FillColorHex = "31B6A1";

            PresentationPackageWriter.Save(document, packagePath);
            InjectSyntheticTiming(packagePath);
            ValidateTimeline(packagePath, renderDirectory);
        }

        private static void InjectSyntheticTiming(string packagePath)
        {
            using (ZipArchive archive = ZipFile.Open(
                packagePath,
                ZipArchiveMode.Update))
            {
                const string partName = "ppt/slides/slide1.xml";
                ZipArchiveEntry entry = archive.GetEntry(partName);
                if (entry == null)
                    throw new InvalidOperationException("Synthetic slide1.xml is missing.");

                XmlDocument slide = new XmlDocument();
                slide.PreserveWhitespace = false;
                using (Stream input = entry.Open())
                    slide.Load(input);

                XmlNode oldTiming = FindFirst(slide.DocumentElement, "timing");
                if (oldTiming != null && oldTiming.ParentNode != null)
                    oldTiming.ParentNode.RemoveChild(oldTiming);

                XmlDocumentFragment fragment = slide.CreateDocumentFragment();
                fragment.InnerXml = BuildTimingXml();
                slide.DocumentElement.AppendChild(fragment);

                entry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(
                    partName,
                    CompressionLevel.Optimal);
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                settings.Indent = false;

                using (Stream output = replacement.Open())
                using (XmlWriter writer = XmlWriter.Create(output, settings))
                    slide.Save(writer);
            }
        }

        private static string BuildTimingXml()
        {
            return
                "<p:timing xmlns:p=\"" + PresentationNamespace + "\">" +
                "<p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" nodeType=\"tmRoot\">" +
                "<p:childTnLst><p:seq><p:cTn id=\"2\" dur=\"indefinite\" nodeType=\"mainSeq\">" +
                "<p:childTnLst>" +

                // Step 1: click entrance.
                "<p:par><p:cTn id=\"10\" dur=\"450\" nodeType=\"clickEffect\" presetClass=\"entr\" presetID=\"1\">" +
                "<p:stCondLst><p:cond evt=\"onClick\" delay=\"0\"/></p:stCondLst>" +
                "<p:childTnLst><p:animEffect transition=\"in\" filter=\"fade\">" +
                "<p:cBhvr><p:cTn id=\"11\" dur=\"450\"/><p:tgtEl><p:spTgt spid=\"2\"/></p:tgtEl></p:cBhvr>" +
                "</p:animEffect></p:childTnLst></p:cTn></p:par>" +

                // Same step: emphasis starts with previous after 80 ms.
                "<p:par><p:cTn id=\"20\" dur=\"300\" accel=\"25000\" decel=\"15000\" nodeType=\"withEffect\" presetClass=\"emph\" presetID=\"3\">" +
                "<p:stCondLst><p:cond evt=\"onBegin\" delay=\"80\"/></p:stCondLst>" +
                "<p:childTnLst><p:animClr>" +
                "<p:to><a:srgbClr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"FF6633\"/></p:to>" +
                "<p:cBhvr><p:cTn id=\"21\" dur=\"300\"/><p:tgtEl><p:spTgt spid=\"2\"/></p:tgtEl></p:cBhvr>" +
                "</p:animClr></p:childTnLst></p:cTn></p:par>" +

                // Same step: exit starts after previous with an extra delay.
                "<p:par><p:cTn id=\"30\" dur=\"250\" nodeType=\"afterEffect\" presetClass=\"exit\" presetID=\"10\">" +
                "<p:stCondLst><p:cond evt=\"onEnd\" delay=\"120\"/></p:stCondLst>" +
                "<p:childTnLst><p:animEffect transition=\"out\" filter=\"fade\">" +
                "<p:cBhvr><p:cTn id=\"31\" dur=\"250\"/><p:tgtEl><p:spTgt spid=\"3\"/></p:tgtEl></p:cBhvr>" +
                "</p:animEffect></p:childTnLst></p:cTn></p:par>" +

                // Step 2: another click starts a motion-path approximation.
                "<p:par><p:cTn id=\"40\" dur=\"600\" nodeType=\"clickEffect\" presetClass=\"path\" presetID=\"1\" autoRev=\"1\">" +
                "<p:stCondLst><p:cond evt=\"onClick\" delay=\"40\"/></p:stCondLst>" +
                "<p:childTnLst><p:animMotion path=\"M 0 0 L 0.2 0\">" +
                "<p:cBhvr><p:cTn id=\"41\" dur=\"600\"/><p:tgtEl><p:spTgt spid=\"2\"/></p:tgtEl></p:cBhvr>" +
                "</p:animMotion></p:childTnLst></p:cTn></p:par>" +
                "<p:par><p:cTn id=\"50\" dur=\"600\" nodeType=\"withEffect\" presetClass=\"emph\" presetID=\"6\">" +
                "<p:stCondLst><p:cond evt=\"onBegin\" delay=\"0\"/></p:stCondLst>" +
                "<p:childTnLst><p:animScale><p:by x=\"125000\" y=\"115000\"/>" +
                "<p:cBhvr><p:cTn id=\"51\" dur=\"600\"/><p:tgtEl><p:spTgt spid=\"2\"/></p:tgtEl></p:cBhvr>" +
                "</p:animScale></p:childTnLst></p:cTn></p:par>" +

                "<p:par><p:cTn id=\"60\" dur=\"600\" nodeType=\"withEffect\" presetClass=\"emph\" presetID=\"8\">" +
                "<p:stCondLst><p:cond evt=\"onBegin\" delay=\"0\"/></p:stCondLst>" +
                "<p:childTnLst><p:animRot by=\"10800000\">" +
                "<p:cBhvr><p:cTn id=\"61\" dur=\"600\"/><p:tgtEl><p:spTgt spid=\"2\"/></p:tgtEl></p:cBhvr>" +
                "</p:animRot></p:childTnLst></p:cTn></p:par>" +


                "</p:childTnLst></p:cTn></p:seq></p:childTnLst>" +
                "</p:cTn></p:par></p:tnLst></p:timing>";
        }

        private static void ValidateTimeline(
            string packagePath,
            string renderDirectory)
        {
            List<InternalPptxRenderer.SlideAnimationTimeline> timelines =
                InternalPptxRenderer.ReadAnimationTimelines(packagePath);

            if (timelines == null || timelines.Count != 1)
                throw new InvalidOperationException("Animation timeline slide count is incorrect.");

            InternalPptxRenderer.SlideAnimationTimeline timeline = timelines[0];
            if (timeline == null || timeline.Steps.Count != 2)
            {
                throw new InvalidOperationException(
                    "Expected two animation click groups in the synthetic timing tree.");
            }

            InternalPptxRenderer.AnimationStepSpec first = timeline.Steps[0];
            InternalPptxRenderer.AnimationStepSpec second = timeline.Steps[1];

            if (!first.RequiresClick || first.Actions.Count != 3)
                throw new InvalidOperationException("The first animation step was not grouped correctly.");

            RequireActionClass(first, "entrance");
            RequireActionClass(first, "emphasis");
            RequireActionClass(first, "exit");
            RequireActionClass(second, "motion");

            bool foundColor = false;
            bool foundEasing = false;

            for (int i = 0;
                 i < first.Actions.Count;
                 i++)
            {
                InternalPptxRenderer.AnimationActionSpec action =
                    first.Actions[i];

                if (action.HasColor &&
                    action.ColorTo.R == 0xFF &&
                    action.ColorTo.G == 0x66 &&
                    action.ColorTo.B == 0x33)
                {
                    foundColor = true;
                }

                if (Math.Abs(
                        action.Acceleration -
                        0.25f) < 0.001f &&
                    Math.Abs(
                        action.Deceleration -
                        0.15f) < 0.001f)
                {
                    foundEasing = true;
                }
            }

            if (!foundColor ||
                !foundEasing)
            {
                throw new InvalidOperationException(
                    "Color emphasis or acceleration/deceleration timing was not parsed.");
            }

            if (!second.RequiresClick ||
                second.Actions.Count != 3)
            {
                throw new InvalidOperationException(
                    "The second animation step was not parsed correctly.");
            }

            bool foundScale = false;
            bool foundRotation = false;

            for (int i = 0;
                 i < second.Actions.Count;
                 i++)
            {
                if (second.Actions[i].HasScale)
                {
                    foundScale =
                        Math.Abs(
                            second.Actions[i].ScaleToX -
                            1.25f) < 0.01f;
                }

                if (second.Actions[i].HasRotation)
                {
                    foundRotation =
                        Math.Abs(
                            second.Actions[i].RotationToDegrees -
                            180f) < 0.1f;
                }
            }

            if (!foundScale ||
                !foundRotation)
            {
                throw new InvalidOperationException(
                    "Scale/rotation animation properties were not parsed.");
            }

            if (first.TotalDurationMs < 450)
                throw new InvalidOperationException("Animation duration/after-previous timing was not retained.");

            if (second.AutoStartDelayMs != 40)
                throw new InvalidOperationException("Animation start delay was not retained.");

            for (int state = 0; state <= 2; state++)
            {
                string frame = InternalPptxRenderer.RenderAnimationTimelineState(
                    packagePath,
                    renderDirectory,
                    0,
                    state);

                if (string.IsNullOrEmpty(frame) ||
                    !File.Exists(frame) ||
                    new FileInfo(frame).Length <= 0)
                {
                    throw new InvalidOperationException(
                        "Animation timeline render state " + state.ToString() + " was not created.");
                }
            }

            List<string> progressFrames =
                InternalPptxRenderer.RenderAnimationTimelineStepFrames(
                    packagePath,
                    renderDirectory,
                    0,
                    0,
                    8);

            if (progressFrames == null ||
                progressFrames.Count != 8)
            {
                throw new InvalidOperationException(
                    "Object-level animation progress frames were not created.");
            }

            for (int i = 0; i < progressFrames.Count; i++)
            {
                if (string.IsNullOrEmpty(progressFrames[i]) ||
                    !File.Exists(progressFrames[i]) ||
                    new FileInfo(progressFrames[i]).Length <= 0)
                {
                    throw new InvalidOperationException(
                        "Animation progress frame " +
                        i.ToString() +
                        " is missing.");
                }
            }

            List<string> transformedFrames =
                InternalPptxRenderer.RenderAnimationTimelineStepFrames(
                    packagePath,
                    renderDirectory,
                    0,
                    1,
                    8);

            if (transformedFrames == null ||
                transformedFrames.Count != 8)
            {
                throw new InvalidOperationException(
                    "Motion/scale/rotation animation frames were not created.");
            }

            for (int i = 0;
                 i < transformedFrames.Count;
                 i++)
            {
                if (string.IsNullOrEmpty(
                        transformedFrames[i]) ||
                    !File.Exists(
                        transformedFrames[i]) ||
                    new FileInfo(
                        transformedFrames[i])
                        .Length <= 0)
                {
                    throw new InvalidOperationException(
                        "Transformed animation frame " +
                        i.ToString() +
                        " is missing.");
                }
            }
        }

        private static void RequireActionClass(
            InternalPptxRenderer.AnimationStepSpec step,
            string effectClass)
        {
            for (int i = 0; i < step.Actions.Count; i++)
            {
                if (string.Equals(
                        step.Actions[i].EffectClass,
                        effectClass,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            throw new InvalidOperationException(
                "Synthetic timing did not preserve the expected " + effectClass + " action.");
        }

        private static XmlNode FindFirst(XmlNode node, string localName)
        {
            if (node == null)
                return null;
            if (node.LocalName == localName)
                return node;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirst(node.ChildNodes[i], localName);
                if (found != null)
                    return found;
            }
            return null;
        }

        private static void DeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }

        private static void DeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
            }
            catch { }
        }
    }
}
