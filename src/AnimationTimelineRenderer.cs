using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        public sealed class AnimationActionSpec
        {
            public string ShapeId = string.Empty;
            public string EffectClass = "unknown";
            public string Trigger = "onClick";
            public string VisualKind = "fade";
            public string PresetClass = string.Empty;
            public string PresetId = string.Empty;
            public string PresetSubtype = string.Empty;
            public int DelayMs;
            public int DurationMs = 300;
            public int StartOffsetMs;
            public int RepeatCount = 1;
            public bool AutoReverse;
        }

        public sealed class AnimationStepSpec
        {
            public bool RequiresClick = true;
            public int AutoStartDelayMs;
            public int TotalDurationMs = 300;
            public string VisualKind = "fade";
            public readonly List<AnimationActionSpec> Actions =
                new List<AnimationActionSpec>();
        }

        public sealed class SlideAnimationTimeline
        {
            public readonly List<AnimationStepSpec> Steps =
                new List<AnimationStepSpec>();
        }

        public static List<SlideAnimationTimeline> ReadAnimationTimelines(
            string file)
        {
            List<SlideAnimationTimeline> result =
                new List<SlideAnimationTimeline>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    XmlDocument document = LoadXml(zip, info.SlideParts[i]);
                    result.Add(ParseAnimationTimeline(document));
                }
            }

            return result;
        }

        public static string RenderAnimationTimelineState(
            string file,
            string cacheDir,
            int slideIndex,
            int completedSteps)
        {
            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);
                if (slideIndex < 0 || slideIndex >= info.SlideParts.Count)
                    return null;

                XmlDocument slideDocument = LoadXml(zip, info.SlideParts[slideIndex]);
                SlideAnimationTimeline timeline = ParseAnimationTimeline(slideDocument);
                completedSteps = Math.Max(
                    0,
                    Math.Min(timeline.Steps.Count, completedSteps));

                HashSet<string> hidden = BuildAnimationHiddenSet(
                    timeline,
                    completedSteps);

                int width = 1920;
                int height = (int)Math.Round(
                    width * ((double)info.HeightEmu / info.WidthEmu));
                if (height < 200 || height > 4000)
                    height = 1080;

                string output = Path.Combine(
                    cacheDir,
                    "timeline_" + (slideIndex + 1).ToString("D4") +
                    "_" + completedSteps.ToString("D3") + ".png");

                if (!File.Exists(output))
                {
                    using (Bitmap bitmap = RenderSlide(
                        zip,
                        info,
                        info.SlideParts[slideIndex],
                        width,
                        height,
                        hidden))
                    {
                        if (completedSteps > 0 &&
                            completedSteps <= timeline.Steps.Count)
                        {
                            OverlayTransientAnimationState(
                                bitmap,
                                slideDocument,
                                info,
                                timeline.Steps[completedSteps - 1]);
                        }

                        bitmap.Save(output, ImageFormat.Png);
                    }
                }

                return output;
            }
        }

        private static SlideAnimationTimeline ParseAnimationTimeline(
            XmlDocument document)
        {
            SlideAnimationTimeline timeline = new SlideAnimationTimeline();
            if (document == null)
                return timeline;

            XmlNode timing = FindFirst(document, "timing");
            if (timing == null)
                return timeline;

            List<AnimationActionSpec> actions =
                new List<AnimationActionSpec>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            List<XmlNode> targets = FindAll(timing, "spTgt");

            for (int i = 0; i < targets.Count; i++)
            {
                XmlNode target = targets[i];
                string shapeId = GetAttr(target, "spid");
                if (string.IsNullOrEmpty(shapeId))
                    continue;

                XmlNode effect = FindAnimationEffectNode(target);
                XmlNode timingNode = FindAnimationTimingNode(effect ?? target);
                string timingId = timingNode == null
                    ? string.Empty
                    : GetAttr(timingNode, "id") ?? string.Empty;
                string key = shapeId + "|" + timingId + "|" +
                    (effect == null ? target.ParentNode == null ? string.Empty : target.ParentNode.LocalName : effect.LocalName);

                if (!seen.Add(key))
                    continue;

                AnimationActionSpec action = BuildAnimationAction(
                    shapeId,
                    effect,
                    timingNode);
                actions.Add(action);
            }

            BuildAnimationSteps(actions, timeline);
            return timeline;
        }

        private static AnimationActionSpec BuildAnimationAction(
            string shapeId,
            XmlNode effect,
            XmlNode timingNode)
        {
            AnimationActionSpec action = new AnimationActionSpec();
            action.ShapeId = shapeId ?? string.Empty;

            if (timingNode != null)
            {
                action.PresetClass = GetAttr(timingNode, "presetClass") ?? string.Empty;
                action.PresetId = GetAttr(timingNode, "presetID") ?? string.Empty;
                action.PresetSubtype = GetAttr(timingNode, "presetSubtype") ?? string.Empty;
                action.Trigger = ReadAnimationTrigger(timingNode);
                action.DelayMs = ReadAnimationDelay(timingNode);
                action.DurationMs = ReadAnimationDuration(timingNode);
                action.RepeatCount = ReadAnimationRepeatCount(timingNode);
                action.AutoReverse = IsTrueValue(GetAttr(timingNode, "autoRev"));
            }

            action.EffectClass = ClassifyAnimationEffect(
                effect,
                action.PresetClass);
            action.VisualKind = ChooseAnimationVisualKind(
                effect,
                action);
            return action;
        }

        private static XmlNode FindAnimationEffectNode(XmlNode target)
        {
            XmlNode node = target;
            while (node != null)
            {
                string name = node.LocalName;
                if (name == "animEffect" ||
                    name == "anim" ||
                    name == "set" ||
                    name == "animMotion" ||
                    name == "animScale" ||
                    name == "animRot" ||
                    name == "animClr" ||
                    name == "cmd")
                {
                    return node;
                }

                if (name == "timing")
                    break;
                node = node.ParentNode;
            }
            return null;
        }

        private static XmlNode FindAnimationTimingNode(XmlNode node)
        {
            XmlNode current = node;
            XmlNode fallback = null;

            while (current != null)
            {
                if (current.LocalName == "cTn")
                {
                    if (fallback == null)
                        fallback = current;

                    string nodeType = GetAttr(current, "nodeType");
                    string presetClass = GetAttr(current, "presetClass");
                    if (!string.IsNullOrEmpty(nodeType) ||
                        !string.IsNullOrEmpty(presetClass))
                    {
                        return current;
                    }
                }

                if (current.LocalName == "timing")
                    break;
                current = current.ParentNode;
            }

            return fallback;
        }

        private static string ClassifyAnimationEffect(
            XmlNode effect,
            string presetClass)
        {
            if (string.Equals(presetClass, "entr", StringComparison.OrdinalIgnoreCase))
                return "entrance";
            if (string.Equals(presetClass, "exit", StringComparison.OrdinalIgnoreCase))
                return "exit";
            if (string.Equals(presetClass, "emph", StringComparison.OrdinalIgnoreCase))
                return "emphasis";
            if (string.Equals(presetClass, "path", StringComparison.OrdinalIgnoreCase))
                return "motion";

            if (effect == null)
                return "entrance";

            if (effect.LocalName == "animMotion")
                return "motion";
            if (effect.LocalName == "animClr" ||
                effect.LocalName == "animScale" ||
                effect.LocalName == "animRot" ||
                effect.LocalName == "anim")
                return "emphasis";

            if (effect.LocalName == "animEffect")
            {
                string transition = GetAttr(effect, "transition");
                if (transition == "out") return "exit";
                if (transition == "in") return "entrance";
            }

            if (effect.LocalName == "set")
            {
                XmlNode strValue = FindFirst(effect, "strVal");
                string value = GetAttr(strValue, "val");
                if (!string.IsNullOrEmpty(value) &&
                    value.IndexOf("hidden", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "exit";
                return "entrance";
            }

            return "emphasis";
        }

        private static string ReadAnimationTrigger(XmlNode timingNode)
        {
            string nodeType = GetAttr(timingNode, "nodeType") ?? string.Empty;
            if (nodeType == "clickEffect" || nodeType == "clickPar")
                return "onClick";
            if (nodeType == "withEffect")
                return "withPrevious";
            if (nodeType == "afterEffect")
                return "afterPrevious";

            XmlNode stCondLst = DirectChild(timingNode, "stCondLst");
            if (stCondLst != null)
            {
                for (int i = 0; i < stCondLst.ChildNodes.Count; i++)
                {
                    XmlNode condition = stCondLst.ChildNodes[i];
                    if (condition.LocalName != "cond")
                        continue;

                    string evt = GetAttr(condition, "evt") ?? string.Empty;
                    if (evt == "onClick" || evt == "onNext")
                        return "onClick";
                    if (evt == "onBegin")
                        return "withPrevious";
                    if (evt == "onEnd")
                        return "afterPrevious";

                    string delay = GetAttr(condition, "delay");
                    if (string.Equals(delay, "indefinite", StringComparison.OrdinalIgnoreCase))
                        return "onClick";
                }
            }

            return "onClick";
        }

        private static int ReadAnimationDelay(XmlNode timingNode)
        {
            XmlNode stCondLst = DirectChild(timingNode, "stCondLst");
            if (stCondLst == null)
                return 0;

            for (int i = 0; i < stCondLst.ChildNodes.Count; i++)
            {
                XmlNode condition = stCondLst.ChildNodes[i];
                if (condition.LocalName != "cond")
                    continue;

                string delay = GetAttr(condition, "delay");
                int parsed;
                if (!string.IsNullOrEmpty(delay) &&
                    int.TryParse(delay, out parsed))
                {
                    return Math.Max(0, Math.Min(60000, parsed));
                }
            }
            return 0;
        }

        private static int ReadAnimationDuration(XmlNode timingNode)
        {
            if (timingNode == null)
                return 300;

            string duration = GetAttr(timingNode, "dur");
            int parsed;
            if (string.IsNullOrEmpty(duration) ||
                !int.TryParse(duration, out parsed))
            {
                parsed = 300;
            }

            parsed = Math.Max(0, Math.Min(60000, parsed));
            string speed = GetAttr(timingNode, "spd");
            int speedValue;
            if (!string.IsNullOrEmpty(speed) &&
                int.TryParse(speed, out speedValue) &&
                speedValue > 0)
            {
                parsed = (int)Math.Round(parsed * 100000.0 / speedValue);
            }

            return Math.Max(0, Math.Min(60000, parsed));
        }

        private static int ReadAnimationRepeatCount(XmlNode timingNode)
        {
            string repeat = GetAttr(timingNode, "repeatCount");
            if (string.IsNullOrEmpty(repeat) ||
                string.Equals(repeat, "indefinite", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            double value;
            if (!double.TryParse(
                    repeat,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value))
                return 1;

            if (value > 1000.0)
                value /= 1000.0;

            return Math.Max(1, Math.Min(20, (int)Math.Round(value)));
        }

        private static bool IsTrueValue(string value)
        {
            return value == "1" ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static string ChooseAnimationVisualKind(
            XmlNode effect,
            AnimationActionSpec action)
        {
            if (action.EffectClass == "motion")
                return "push";

            string filter = effect == null
                ? string.Empty
                : GetAttr(effect, "filter") ?? string.Empty;
            string lower = filter.ToLowerInvariant();

            if (lower.IndexOf("wipe") >= 0 ||
                lower.IndexOf("blinds") >= 0 ||
                lower.IndexOf("strips") >= 0 ||
                lower.IndexOf("checker") >= 0)
                return "wipe";
            if (lower.IndexOf("fly") >= 0 ||
                lower.IndexOf("crawl") >= 0 ||
                lower.IndexOf("peek") >= 0)
                return "push";

            return "fade";
        }

        private static void BuildAnimationSteps(
            List<AnimationActionSpec> actions,
            SlideAnimationTimeline timeline)
        {
            AnimationStepSpec current = null;
            int currentEnd = 0;

            for (int i = 0; i < actions.Count; i++)
            {
                AnimationActionSpec action = actions[i];
                bool newStep = current == null || action.Trigger == "onClick";

                if (newStep)
                {
                    current = new AnimationStepSpec();
                    current.RequiresClick = action.Trigger == "onClick";
                    current.AutoStartDelayMs = action.DelayMs;
                    current.TotalDurationMs = 0;
                    current.VisualKind = action.VisualKind;
                    timeline.Steps.Add(current);
                    currentEnd = 0;
                }

                if (action.Trigger == "afterPrevious")
                    action.StartOffsetMs = currentEnd + action.DelayMs;
                else
                    action.StartOffsetMs = action.DelayMs;

                int repeats = Math.Max(1, action.RepeatCount);
                int duration = action.DurationMs * repeats;
                if (action.AutoReverse)
                    duration *= 2;

                int actionEnd = action.StartOffsetMs + duration;
                currentEnd = Math.Max(currentEnd, actionEnd);
                current.TotalDurationMs = Math.Max(
                    current.TotalDurationMs,
                    Math.Max(1, actionEnd));

                if (action.VisualKind == "push")
                    current.VisualKind = "push";
                else if (action.VisualKind == "wipe" && current.VisualKind != "push")
                    current.VisualKind = "wipe";

                current.Actions.Add(action);
            }
        }

        private static HashSet<string> BuildAnimationHiddenSet(
            SlideAnimationTimeline timeline,
            int completedSteps)
        {
            HashSet<string> hidden =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int s = 0; s < timeline.Steps.Count; s++)
            {
                AnimationStepSpec step = timeline.Steps[s];
                for (int a = 0; a < step.Actions.Count; a++)
                {
                    AnimationActionSpec action = step.Actions[a];
                    if (action.EffectClass == "entrance" &&
                        !string.IsNullOrEmpty(action.ShapeId))
                    {
                        hidden.Add(action.ShapeId);
                    }
                }
            }

            for (int s = 0; s < completedSteps && s < timeline.Steps.Count; s++)
            {
                AnimationStepSpec step = timeline.Steps[s];
                for (int a = 0; a < step.Actions.Count; a++)
                {
                    AnimationActionSpec action = step.Actions[a];
                    if (string.IsNullOrEmpty(action.ShapeId))
                        continue;

                    if (action.EffectClass == "entrance")
                        hidden.Remove(action.ShapeId);
                    else if (action.EffectClass == "exit")
                        hidden.Add(action.ShapeId);
                }
            }

            return hidden;
        }

        private static void OverlayTransientAnimationState(
            Bitmap bitmap,
            XmlDocument slideDocument,
            PresentationInfo info,
            AnimationStepSpec step)
        {
            if (bitmap == null || slideDocument == null || step == null)
                return;

            Dictionary<string, RectangleF> bounds =
                ReadSlideShapeBounds(slideDocument, info, bitmap.Width, bitmap.Height);

            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;

                for (int i = 0; i < step.Actions.Count; i++)
                {
                    AnimationActionSpec action = step.Actions[i];
                    if (action.EffectClass != "emphasis" &&
                        action.EffectClass != "motion")
                        continue;

                    RectangleF rect;
                    if (!bounds.TryGetValue(action.ShapeId, out rect))
                        continue;

                    if (action.EffectClass == "motion")
                    {
                        using (Pen pen = new Pen(Color.FromArgb(165, 70, 125, 230), 3f))
                        {
                            pen.DashStyle = DashStyle.Dash;
                            g.DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);
                            g.DrawLine(
                                pen,
                                rect.Left + rect.Width * 0.25f,
                                rect.Top + rect.Height / 2f,
                                rect.Right + Math.Min(36f, rect.Width * 0.2f),
                                rect.Top + rect.Height / 2f);
                        }
                    }
                    else
                    {
                        using (Pen glow = new Pen(Color.FromArgb(190, 245, 170, 40), 5f))
                            g.DrawRectangle(glow, rect.X, rect.Y, rect.Width, rect.Height);
                    }
                }
            }
        }

        private static Dictionary<string, RectangleF> ReadSlideShapeBounds(
            XmlDocument slideDocument,
            PresentationInfo info,
            int width,
            int height)
        {
            Dictionary<string, RectangleF> result =
                new Dictionary<string, RectangleF>(StringComparer.OrdinalIgnoreCase);
            if (slideDocument == null || info == null)
                return result;

            XmlNode tree = FindFirst(slideDocument, "spTree");
            if (tree == null)
                return result;

            TransformContext root = new TransformContext();
            root.Ax = (double)width / info.WidthEmu;
            root.Ay = (double)height / info.HeightEmu;
            root.Bx = 0;
            root.By = 0;
            CollectAnimationShapeBounds(tree, root, result);
            return result;
        }

        private static void CollectAnimationShapeBounds(
            XmlNode container,
            TransformContext context,
            Dictionary<string, RectangleF> result)
        {
            if (container == null)
                return;

            for (int i = 0; i < container.ChildNodes.Count; i++)
            {
                XmlNode child = container.ChildNodes[i];
                if (child.LocalName == "grpSp")
                {
                    CollectAnimationShapeBounds(
                        child,
                        BuildGroupContext(child, context),
                        result);
                    continue;
                }

                if (child.LocalName != "sp" &&
                    child.LocalName != "pic" &&
                    child.LocalName != "graphicFrame" &&
                    child.LocalName != "cxnSp")
                    continue;

                string id = GetShapeId(child);
                RectangleF rect;
                if (!string.IsNullOrEmpty(id) && TryGetRect(child, context, out rect))
                    result[id] = rect;
            }
        }
    }
}
