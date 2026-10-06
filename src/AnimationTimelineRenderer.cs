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
            public string VisualDirection = "fromLeft";
            public string MotionPath = string.Empty;
            public string PresetClass = string.Empty;
            public string PresetId = string.Empty;
            public string PresetSubtype = string.Empty;
            public bool HasScale;
            public float ScaleFromX = 1f;
            public float ScaleFromY = 1f;
            public float ScaleToX = 1f;
            public float ScaleToY = 1f;
            public bool HasRotation;
            public float RotationFromDegrees;
            public float RotationToDegrees;
            public bool HasColor;
            public Color ColorFrom = Color.Empty;
            public Color ColorTo = Color.Empty;
            public float Acceleration;
            public float Deceleration;
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
                action.Acceleration = ReadAnimationPercentAttribute(
                    timingNode,
                    "accel");
                action.Deceleration = ReadAnimationPercentAttribute(
                    timingNode,
                    "decel");
            }

            action.EffectClass = ClassifyAnimationEffect(
                effect,
                action.PresetClass);
            action.VisualKind = ChooseAnimationVisualKind(
                effect,
                action);
            action.VisualDirection = ChooseAnimationVisualDirection(
                effect,
                action);
            if (effect != null &&
                effect.LocalName == "animMotion")
            {
                action.MotionPath =
                    GetAttr(
                        effect,
                        "path") ??
                    string.Empty;
            }
            else if (effect != null &&
                effect.LocalName == "animScale")
            {
                ReadAnimationScale(
                    effect,
                    action);
            }
            else if (effect != null &&
                effect.LocalName == "animRot")
            {
                ReadAnimationRotation(
                    effect,
                    action);
            }
            else if (effect != null &&
                effect.LocalName == "animClr")
            {
                ReadAnimationColor(
                    effect,
                    action);
            }

            return action;
        }

        private static void ReadAnimationScale(
            XmlNode effect,
            AnimationActionSpec action)
        {
            if (effect == null ||
                action == null)
            {
                return;
            }

            action.HasScale = true;

            float fromX = 1f;
            float fromY = 1f;
            float toX = 1f;
            float toY = 1f;
            bool hasFrom =
                TryReadAnimationScalePoint(
                    DirectChild(
                        effect,
                        "from"),
                    out fromX,
                    out fromY);
            bool hasTo =
                TryReadAnimationScalePoint(
                    DirectChild(
                        effect,
                        "to"),
                    out toX,
                    out toY);

            float byX;
            float byY;
            bool hasBy =
                TryReadAnimationScalePoint(
                    DirectChild(
                        effect,
                        "by"),
                    out byX,
                    out byY);

            if (hasFrom)
            {
                action.ScaleFromX = fromX;
                action.ScaleFromY = fromY;
            }

            if (hasTo)
            {
                action.ScaleToX = toX;
                action.ScaleToY = toY;
            }
            else if (hasBy)
            {
                action.ScaleToX =
                    hasFrom
                        ? action.ScaleFromX *
                            byX
                        : byX;
                action.ScaleToY =
                    hasFrom
                        ? action.ScaleFromY *
                            byY
                        : byY;
            }
            else
            {
                action.ScaleToX = 1.12f;
                action.ScaleToY = 1.12f;
            }
        }

        private static bool TryReadAnimationScalePoint(
            XmlNode node,
            out float x,
            out float y)
        {
            x = 1f;
            y = 1f;

            if (node == null)
                return false;

            string rawX =
                GetAttr(
                    node,
                    "x");
            string rawY =
                GetAttr(
                    node,
                    "y");

            bool hasX =
                TryParseAnimationScaleValue(
                    rawX,
                    out x);
            bool hasY =
                TryParseAnimationScaleValue(
                    rawY,
                    out y);

            if (!hasX &&
                hasY)
            {
                x = y;
            }
            else if (hasX &&
                !hasY)
            {
                y = x;
            }

            return hasX || hasY;
        }

        private static bool TryParseAnimationScaleValue(
            string raw,
            out float value)
        {
            value = 1f;

            if (string.IsNullOrEmpty(raw))
                return false;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                float percent;
                if (float.TryParse(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out percent))
                {
                    value =
                        percent /
                        100f;
                    return true;
                }

                return false;
            }

            float parsed;
            if (!float.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return false;
            }

            if (Math.Abs(parsed) > 10f)
            {
                parsed /=
                    100000f;
            }

            value =
                Math.Max(
                    0.02f,
                    Math.Min(
                        8f,
                        parsed));
            return true;
        }

        private static void ReadAnimationRotation(
            XmlNode effect,
            AnimationActionSpec action)
        {
            if (effect == null ||
                action == null)
            {
                return;
            }

            action.HasRotation = true;

            float from = 0f;
            float to = 0f;
            float by = 0f;
            bool hasFrom =
                TryParseAnimationAngle(
                    GetAttr(
                        effect,
                        "from"),
                    out from);
            bool hasTo =
                TryParseAnimationAngle(
                    GetAttr(
                        effect,
                        "to"),
                    out to);
            bool hasBy =
                TryParseAnimationAngle(
                    GetAttr(
                        effect,
                        "by"),
                    out by);

            action.RotationFromDegrees =
                hasFrom
                    ? from
                    : 0f;

            if (hasTo)
            {
                action.RotationToDegrees =
                    to;
            }
            else if (hasBy)
            {
                action.RotationToDegrees =
                    action.RotationFromDegrees +
                    by;
            }
            else
            {
                action.RotationToDegrees =
                    360f;
            }
        }

        private static bool TryParseAnimationAngle(
            string raw,
            out float degrees)
        {
            degrees = 0f;

            if (string.IsNullOrEmpty(raw))
                return false;

            float parsed;
            if (!float.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return false;
            }

            if (Math.Abs(parsed) > 720f)
            {
                parsed /=
                    60000f;
            }

            degrees =
                Math.Max(
                    -3600f,
                    Math.Min(
                        3600f,
                        parsed));
            return true;
        }

        private static float ReadAnimationPercentAttribute(
            XmlNode node,
            string name)
        {
            if (node == null)
                return 0f;

            string raw =
                GetAttr(
                    node,
                    name);

            if (string.IsNullOrEmpty(raw))
                return 0f;

            float parsed;
            if (!float.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return 0f;
            }

            if (Math.Abs(parsed) > 1f)
                parsed /= 100000f;

            return Math.Max(
                0f,
                Math.Min(
                    1f,
                    parsed));
        }

        private static void ReadAnimationColor(
            XmlNode effect,
            AnimationActionSpec action)
        {
            if (effect == null ||
                action == null)
            {
                return;
            }

            Color from;
            Color to;
            Color by;

            bool hasFrom =
                TryReadAnimationColorNode(
                    DirectChild(
                        effect,
                        "from"),
                    out from);

            bool hasTo =
                TryReadAnimationColorNode(
                    DirectChild(
                        effect,
                        "to"),
                    out to);

            bool hasBy =
                TryReadAnimationColorNode(
                    DirectChild(
                        effect,
                        "by"),
                    out by);

            if (!hasTo && hasBy)
            {
                to = by;
                hasTo = true;
            }

            if (!hasTo)
                return;

            action.HasColor = true;
            action.ColorFrom =
                hasFrom
                    ? from
                    : Color.Empty;
            action.ColorTo = to;
        }

        private static bool TryReadAnimationColorNode(
            XmlNode node,
            out Color color)
        {
            color = Color.Empty;

            if (node == null)
                return false;

            if (node.LocalName == "srgbClr")
            {
                Color? parsed =
                    ParseHexColor(
                        GetAttr(
                            node,
                            "val"));

                if (parsed.HasValue)
                {
                    color =
                        parsed.Value;
                    return true;
                }
            }

            if (node.LocalName == "sysClr")
            {
                string raw =
                    GetAttr(
                        node,
                        "lastClr");

                Color? parsed =
                    ParseHexColor(raw);

                if (parsed.HasValue)
                {
                    color =
                        parsed.Value;
                    return true;
                }
            }

            if (node.LocalName == "prstClr")
            {
                string raw =
                    GetAttr(
                        node,
                        "val");

                if (!string.IsNullOrEmpty(raw))
                {
                    Color named =
                        Color.FromName(raw);

                    if (named.A > 0)
                    {
                        color = named;
                        return true;
                    }
                }
            }

            if (node.LocalName == "rgb")
            {
                int r;
                int g;
                int b;

                if (TryReadAnimationColorComponent(
                        GetAttr(node, "r"),
                        out r) &&
                    TryReadAnimationColorComponent(
                        GetAttr(node, "g"),
                        out g) &&
                    TryReadAnimationColorComponent(
                        GetAttr(node, "b"),
                        out b))
                {
                    color =
                        Color.FromArgb(
                            255,
                            r,
                            g,
                            b);
                    return true;
                }
            }

            for (int i = 0;
                 i < node.ChildNodes.Count;
                 i++)
            {
                if (TryReadAnimationColorNode(
                        node.ChildNodes[i],
                        out color))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryReadAnimationColorComponent(
            string raw,
            out int value)
        {
            value = 0;

            if (string.IsNullOrEmpty(raw))
                return false;

            float parsed;
            if (!float.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return false;
            }

            if (parsed > 255f)
            {
                parsed =
                    parsed *
                    255f /
                    100000f;
            }

            value =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(parsed)));
            return true;
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

        private static string ChooseAnimationVisualDirection(
            XmlNode effect,
            AnimationActionSpec action)
        {
            string text = string.Empty;

            if (effect != null)
            {
                text =
                    (GetAttr(effect, "filter") ?? string.Empty) + " " +
                    (GetAttr(effect, "path") ?? string.Empty);
            }

            text += " " + (action == null ? string.Empty : action.PresetSubtype);
            text = text.ToLowerInvariant();

            if (text.IndexOf("right") >= 0)
                return "fromRight";
            if (text.IndexOf("top") >= 0 ||
                text.IndexOf("up") >= 0)
                return "fromTop";
            if (text.IndexOf("bottom") >= 0 ||
                text.IndexOf("down") >= 0)
                return "fromBottom";
            return "fromLeft";
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

        public static List<string> RenderAnimationTimelineStepFrames(
            string file,
            string cacheDir,
            int slideIndex,
            int stepIndex,
            int frameCount)
        {
            List<string> outputs = new List<string>();
            if (string.IsNullOrEmpty(file) ||
                string.IsNullOrEmpty(cacheDir))
                return outputs;

            frameCount = Math.Max(3, Math.Min(18, frameCount));
            Directory.CreateDirectory(cacheDir);

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);
                if (slideIndex < 0 || slideIndex >= info.SlideParts.Count)
                    return outputs;

                XmlDocument slideDocument =
                    LoadXml(zip, info.SlideParts[slideIndex]);
                SlideAnimationTimeline timeline =
                    ParseAnimationTimeline(slideDocument);

                if (stepIndex < 0 || stepIndex >= timeline.Steps.Count)
                    return outputs;

                AnimationStepSpec step = timeline.Steps[stepIndex];

                HashSet<string> beforeHidden =
                    BuildAnimationHiddenSet(timeline, stepIndex);
                HashSet<string> afterHidden =
                    BuildAnimationHiddenSet(timeline, stepIndex + 1);

                int width = 1920;
                int height = (int)Math.Round(
                    width * ((double)info.HeightEmu / info.WidthEmu));
                if (height < 200 || height > 4000)
                    height = 1080;

                HashSet<string> cleanHidden =
                    new HashSet<string>(
                        beforeHidden,
                        StringComparer.OrdinalIgnoreCase);

                for (int actionIndex = 0;
                     actionIndex < step.Actions.Count;
                     actionIndex++)
                {
                    AnimationActionSpec currentAction =
                        step.Actions[actionIndex];

                    if (currentAction != null &&
                        !string.IsNullOrEmpty(
                            currentAction.ShapeId))
                    {
                        cleanHidden.Add(
                            currentAction.ShapeId);
                    }
                }

                using (Bitmap before = RenderSlide(
                    zip,
                    info,
                    info.SlideParts[slideIndex],
                    width,
                    height,
                    beforeHidden))
                using (Bitmap after = RenderSlide(
                    zip,
                    info,
                    info.SlideParts[slideIndex],
                    width,
                    height,
                    afterHidden))
                using (Bitmap clean = RenderSlide(
                    zip,
                    info,
                    info.SlideParts[slideIndex],
                    width,
                    height,
                    cleanHidden))
                {
                    Dictionary<string, RectangleF> bounds =
                        ReadSlideShapeBounds(
                            slideDocument,
                            info,
                            width,
                            height);

                    for (int i = 1; i <= frameCount; i++)
                    {
                        string output = Path.Combine(
                            cacheDir,
                            "timeline_" +
                            (slideIndex + 1).ToString("D4") +
                            "_step_" +
                            (stepIndex + 1).ToString("D3") +
                            "_frame_" +
                            i.ToString("D3") +
                            ".png");

                        outputs.Add(output);
                        if (File.Exists(output))
                            continue;

                        float progress = i / (float)frameCount;

                        using (Bitmap frame =
                            ComposeAnimationProgressFrame(
                                clean,
                                before,
                                after,
                                bounds,
                                step,
                                progress))
                        {
                            frame.Save(output, ImageFormat.Png);
                        }
                    }
                }
            }

            return outputs;
        }

        private static Bitmap ComposeAnimationProgressFrame(
            Bitmap clean,
            Bitmap before,
            Bitmap after,
            Dictionary<string, RectangleF> bounds,
            AnimationStepSpec step,
            float progress)
        {
            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            Bitmap frame =
                new Bitmap(
                    clean.Width,
                    clean.Height,
                    PixelFormat.Format32bppArgb);

            using (Graphics g =
                Graphics.FromImage(frame))
            {
                g.SmoothingMode =
                    SmoothingMode.AntiAlias;
                g.InterpolationMode =
                    InterpolationMode.HighQualityBicubic;

                g.DrawImage(
                    clean,
                    new Rectangle(
                        0,
                        0,
                        frame.Width,
                        frame.Height));

                for (int i = 0;
                     i < step.Actions.Count;
                     i++)
                {
                    AnimationActionSpec action =
                        step.Actions[i];

                    if (action == null ||
                        string.IsNullOrEmpty(
                            action.ShapeId))
                    {
                        continue;
                    }

                    RectangleF rectF;
                    if (!bounds.TryGetValue(
                            action.ShapeId,
                            out rectF))
                    {
                        continue;
                    }

                    Rectangle rect =
                        ClampAnimationRect(
                            rectF,
                            frame.Width,
                            frame.Height);

                    if (rect.Width <= 0 ||
                        rect.Height <= 0)
                    {
                        continue;
                    }

                    float localProgress =
                        CalculateAnimationActionProgress(
                            action,
                            step,
                            progress);

                    if (action.EffectClass ==
                        "entrance")
                    {
                        if (localProgress > 0f)
                        {
                            DrawEntranceAnimationRegion(
                                g,
                                after,
                                rect,
                                action,
                                localProgress);
                        }
                    }
                    else if (action.EffectClass ==
                        "exit")
                    {
                        DrawExitAnimationRegion(
                            g,
                            before,
                            clean,
                            rect,
                            action,
                            localProgress);
                    }
                    else if (action.EffectClass ==
                        "motion")
                    {
                        DrawMotionAnimationProgress(
                            g,
                            before,
                            clean,
                            rect,
                            action,
                            localProgress,
                            frame.Width,
                            frame.Height);
                    }
                    else if (action.EffectClass ==
                        "emphasis")
                    {
                        DrawEmphasisAnimationProgress(
                            g,
                            before,
                            clean,
                            rect,
                            action,
                            localProgress);
                    }
                    else
                    {
                        DrawAnimationRegion(
                            g,
                            before,
                            rect,
                            rect,
                            1f);
                    }
                }
            }

            return frame;
        }

        private static float CalculateAnimationActionProgress(
            AnimationActionSpec action,
            AnimationStepSpec step,
            float stepProgress)
        {
            int total = Math.Max(1, step.TotalDurationMs);
            double elapsed = total * stepProgress;
            double start = Math.Max(0, action.StartOffsetMs);

            int repeats = Math.Max(1, action.RepeatCount);
            double duration =
                Math.Max(1, action.DurationMs) * repeats;
            if (action.AutoReverse)
                duration *= 2.0;

            if (elapsed <= start)
                return 0f;

            if (elapsed >= start + duration)
            {
                return action.AutoReverse
                    ? 0f
                    : 1f;
            }

            double value = (elapsed - start) / duration;

            if (action.AutoReverse)
            {
                double doubled = value * 2.0;
                value = doubled <= 1.0
                    ? doubled
                    : 2.0 - doubled;
            }

            float normalized =
                (float)Math.Max(
                    0.0,
                    Math.Min(
                        1.0,
                        value));

            return ApplyAnimationEasing(
                normalized,
                action.Acceleration,
                action.Deceleration);
        }

        private static float ApplyAnimationEasing(
            float progress,
            float acceleration,
            float deceleration)
        {
            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            acceleration =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        acceleration));

            deceleration =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        deceleration));

            float sum =
                acceleration +
                deceleration;

            if (sum > 0.98f)
            {
                float scale =
                    0.98f /
                    sum;

                acceleration *= scale;
                deceleration *= scale;
            }

            if (acceleration <= 0.0001f &&
                deceleration <= 0.0001f)
            {
                return progress;
            }

            double constant =
                1.0 -
                acceleration -
                deceleration;

            double area =
                constant +
                acceleration *
                    0.5 +
                deceleration *
                    0.5;

            if (area <= 0.000001)
                return progress;

            double velocity =
                1.0 /
                area;

            double t =
                progress;

            double value;

            if (acceleration > 0f &&
                t < acceleration)
            {
                value =
                    0.5 *
                    velocity /
                    acceleration *
                    t *
                    t;
            }
            else if (t <=
                acceleration +
                constant)
            {
                value =
                    0.5 *
                    velocity *
                    acceleration +
                    velocity *
                    (t -
                     acceleration);
            }
            else
            {
                double before =
                    0.5 *
                    velocity *
                    acceleration +
                    velocity *
                    constant;

                double u =
                    t -
                    acceleration -
                    constant;

                if (deceleration <=
                    0.0001f)
                {
                    value =
                        before +
                        velocity *
                        u;
                }
                else
                {
                    value =
                        before +
                        velocity *
                        u -
                        0.5 *
                        velocity /
                        deceleration *
                        u *
                        u;
                }
            }

            return (float)Math.Max(
                0.0,
                Math.Min(
                    1.0,
                    value));
        }

        private static Rectangle ClampAnimationRect(
            RectangleF source,
            int width,
            int height)
        {
            int left = Math.Max(
                0,
                Math.Min(
                    width,
                    (int)Math.Floor(source.Left)));
            int top = Math.Max(
                0,
                Math.Min(
                    height,
                    (int)Math.Floor(source.Top)));
            int right = Math.Max(
                left,
                Math.Min(
                    width,
                    (int)Math.Ceiling(source.Right)));
            int bottom = Math.Max(
                top,
                Math.Min(
                    height,
                    (int)Math.Ceiling(source.Bottom)));

            return Rectangle.FromLTRB(
                left,
                top,
                right,
                bottom);
        }

        private static void DrawEntranceAnimationRegion(
            Graphics g,
            Bitmap after,
            Rectangle rect,
            AnimationActionSpec action,
            float progress)
        {
            string kind = action.VisualKind ?? "fade";

            if (kind == "wipe")
            {
                Rectangle reveal = GetAnimationRevealRect(
                    rect,
                    action.VisualDirection,
                    progress);
                if (reveal.Width > 0 && reveal.Height > 0)
                    DrawAnimationRegion(
                        g,
                        after,
                        reveal,
                        reveal,
                        1f);
                return;
            }

            if (kind == "push")
            {
                Rectangle destination =
                    GetAnimationPushDestination(
                        rect,
                        action.VisualDirection,
                        progress,
                        false);

                GraphicsState state = g.Save();
                g.SetClip(rect);
                DrawAnimationRegion(
                    g,
                    after,
                    destination,
                    rect,
                    1f);
                g.Restore(state);
                return;
            }

            DrawAnimationRegion(
                g,
                after,
                rect,
                rect,
                progress);
        }

        private static void DrawExitAnimationRegion(
            Graphics g,
            Bitmap before,
            Bitmap background,
            Rectangle rect,
            AnimationActionSpec action,
            float progress)
        {
            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            string kind =
                action.VisualKind ??
                "fade";

            if (kind == "wipe")
            {
                Rectangle remaining =
                    GetAnimationRevealRect(
                        rect,
                        action.VisualDirection,
                        1f -
                        progress);

                if (remaining.Width > 0 &&
                    remaining.Height > 0)
                {
                    DrawAnimationRegion(
                        g,
                        before,
                        remaining,
                        remaining,
                        1f);
                }

                return;
            }

            if (kind == "push")
            {
                Rectangle destination =
                    GetAnimationPushDestination(
                        rect,
                        action.VisualDirection,
                        progress,
                        true);

                GraphicsState state =
                    g.Save();

                try
                {
                    g.SetClip(rect);

                    DrawAnimationRegion(
                        g,
                        before,
                        destination,
                        rect,
                        1f);
                }
                finally
                {
                    g.Restore(state);
                }

                return;
            }

            DrawAnimationRegion(
                g,
                before,
                rect,
                rect,
                1f -
                progress);
        }

        private static Rectangle GetAnimationRevealRect(
            Rectangle rect,
            string direction,
            float progress)
        {
            progress = Math.Max(0f, Math.Min(1f, progress));

            if (direction == "fromRight")
            {
                int width = Math.Max(
                    1,
                    (int)Math.Round(rect.Width * progress));
                return new Rectangle(
                    rect.Right - width,
                    rect.Top,
                    width,
                    rect.Height);
            }

            if (direction == "fromTop")
            {
                int height = Math.Max(
                    1,
                    (int)Math.Round(rect.Height * progress));
                return new Rectangle(
                    rect.Left,
                    rect.Top,
                    rect.Width,
                    height);
            }

            if (direction == "fromBottom")
            {
                int height = Math.Max(
                    1,
                    (int)Math.Round(rect.Height * progress));
                return new Rectangle(
                    rect.Left,
                    rect.Bottom - height,
                    rect.Width,
                    height);
            }

            int revealWidth = Math.Max(
                1,
                (int)Math.Round(rect.Width * progress));
            return new Rectangle(
                rect.Left,
                rect.Top,
                revealWidth,
                rect.Height);
        }

        private static Rectangle GetAnimationPushDestination(
            Rectangle rect,
            string direction,
            float progress,
            bool exiting)
        {
            progress = Math.Max(0f, Math.Min(1f, progress));

            float amount = exiting
                ? progress
                : 1f - progress;

            int dx = 0;
            int dy = 0;

            if (direction == "fromRight")
                dx = (int)Math.Round(rect.Width * amount);
            else if (direction == "fromTop")
                dy = -(int)Math.Round(rect.Height * amount);
            else if (direction == "fromBottom")
                dy = (int)Math.Round(rect.Height * amount);
            else
                dx = -(int)Math.Round(rect.Width * amount);

            return new Rectangle(
                rect.X + dx,
                rect.Y + dy,
                rect.Width,
                rect.Height);
        }

        private static void DrawAnimationRegion(
            Graphics g,
            Bitmap source,
            Rectangle destination,
            Rectangle sourceRect,
            float alpha)
        {
            alpha = Math.Max(0f, Math.Min(1f, alpha));

            using (ImageAttributes attributes =
                new ImageAttributes())
            {
                ColorMatrix matrix = new ColorMatrix();
                matrix.Matrix33 = alpha;
                attributes.SetColorMatrix(
                    matrix,
                    ColorMatrixFlag.Default,
                    ColorAdjustType.Bitmap);

                g.DrawImage(
                    source,
                    destination,
                    sourceRect.X,
                    sourceRect.Y,
                    sourceRect.Width,
                    sourceRect.Height,
                    GraphicsUnit.Pixel,
                    attributes);
            }
        }

        private static void DrawEmphasisAnimationProgress(
            Graphics g,
            Bitmap source,
            Bitmap background,
            Rectangle rect,
            AnimationActionSpec action,
            float progress)
        {
            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            double pulse =
                Math.Sin(
                    Math.PI *
                    progress);

            float scaleX =
                action.HasScale
                    ? LerpAnimationValue(
                        action.ScaleFromX,
                        action.ScaleToX,
                        progress)
                    : 1f +
                        (float)pulse *
                        0.08f;

            float scaleY =
                action.HasScale
                    ? LerpAnimationValue(
                        action.ScaleFromY,
                        action.ScaleToY,
                        progress)
                    : scaleX;

            float rotation =
                action.HasRotation
                    ? LerpAnimationValue(
                        action.RotationFromDegrees,
                        action.RotationToDegrees,
                        progress)
                    : 0f;

            scaleX =
                Math.Max(
                    0.05f,
                    Math.Min(
                        8f,
                        scaleX));
            scaleY =
                Math.Max(
                    0.05f,
                    Math.Min(
                        8f,
                        scaleY));

            float cx =
                rect.Left +
                rect.Width /
                2f;
            float cy =
                rect.Top +
                rect.Height /
                2f;

            Rectangle destination =
                new Rectangle(
                    (int)Math.Round(
                        cx -
                        rect.Width *
                        scaleX /
                        2f),
                    (int)Math.Round(
                        cy -
                        rect.Height *
                        scaleY /
                        2f),
                    Math.Max(
                        1,
                        (int)Math.Round(
                            rect.Width *
                            scaleX)),
                    Math.Max(
                        1,
                        (int)Math.Round(
                            rect.Height *
                            scaleY)));

            GraphicsState state =
                g.Save();

            try
            {
                if (Math.Abs(rotation) >
                    0.001f)
                {
                    g.TranslateTransform(
                        cx,
                        cy);
                    g.RotateTransform(
                        rotation);
                    g.TranslateTransform(
                        -cx,
                        -cy);
                }

                using (Bitmap objectLayer =
                    CreateAnimationObjectLayer(
                        source,
                        background,
                        rect,
                        action.HasColor
                            ? (Color?)action.ColorTo
                            : null,
                        action.ColorFrom,
                        progress))
                {
                    if (objectLayer != null)
                    {
                        g.DrawImage(
                            objectLayer,
                            destination,
                            0,
                            0,
                            objectLayer.Width,
                            objectLayer.Height,
                            GraphicsUnit.Pixel);
                    }
                    else
                    {
                        DrawAnimationRegion(
                            g,
                            source,
                            destination,
                            rect,
                            1f);
                    }
                }
            }
            finally
            {
                g.Restore(state);
            }

            if (!action.HasScale &&
                !action.HasRotation &&
                !action.HasColor &&
                pulse > 0.01)
            {
                int alpha =
                    Math.Max(
                        0,
                        Math.Min(
                            180,
                            (int)Math.Round(
                                170.0 *
                                pulse)));

                float expand =
                    (float)(
                        3.0 +
                        4.0 *
                        pulse);

                using (Pen glow =
                    new Pen(
                        Color.FromArgb(
                            alpha,
                            245,
                            170,
                            40),
                        (float)(
                            2.0 +
                            3.0 *
                            pulse)))
                {
                    g.DrawRectangle(
                        glow,
                        destination.X -
                            expand,
                        destination.Y -
                            expand,
                        destination.Width +
                            expand *
                            2f,
                        destination.Height +
                            expand *
                            2f);
                }
            }
        }

        private static float LerpAnimationValue(
            float from,
            float to,
            float progress)
        {
            return from +
                (to -
                 from) *
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));
        }

        private static void DrawMotionAnimationProgress(
            Graphics g,
            Bitmap source,
            Bitmap background,
            Rectangle rect,
            AnimationActionSpec action,
            float progress,
            int width,
            int height)
        {
            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            float dx = 0f;
            float dy = 0f;

            float pathX;
            float pathY;

            bool hasPathPoint =
                TryEvaluateMotionPath(
                    action.MotionPath,
                    progress,
                    out pathX,
                    out pathY);

            if (hasPathPoint)
            {
                float startX;
                float startY;

                if (!TryEvaluateMotionPath(
                        action.MotionPath,
                        0f,
                        out startX,
                        out startY))
                {
                    startX = 0f;
                    startY = 0f;
                }

                dx =
                    pathX -
                    startX;
                dy =
                    pathY -
                    startY;
            }
            else
            {
                TryReadMotionPathDelta(
                    action.MotionPath,
                    out dx,
                    out dy);

                dx *= progress;
                dy *= progress;
            }

            if (Math.Abs(dx) <
                    0.0001f &&
                Math.Abs(dy) <
                    0.0001f &&
                progress > 0f)
            {
                float fallback =
                    0.18f *
                    progress;

                if (action.VisualDirection ==
                    "fromRight")
                {
                    dx = fallback;
                }
                else if (action.VisualDirection ==
                    "fromTop")
                {
                    dy = -fallback;
                }
                else if (action.VisualDirection ==
                    "fromBottom")
                {
                    dy = fallback;
                }
                else
                {
                    dx = -fallback;
                }
            }

            int offsetX =
                (int)Math.Round(
                    dx *
                    width);

            int offsetY =
                (int)Math.Round(
                    dy *
                    height);

            Rectangle destination =
                new Rectangle(
                    rect.X +
                        offsetX,
                    rect.Y +
                        offsetY,
                    rect.Width,
                    rect.Height);

            using (Bitmap objectLayer =
                CreateAnimationObjectLayer(
                    source,
                    background,
                    rect,
                    null,
                    Color.Empty,
                    progress))
            {
                if (objectLayer != null)
                {
                    g.DrawImage(
                        objectLayer,
                        destination,
                        0,
                        0,
                        objectLayer.Width,
                        objectLayer.Height,
                        GraphicsUnit.Pixel);
                }
                else
                {
                    DrawAnimationRegion(
                        g,
                        source,
                        destination,
                        rect,
                        1f);
                }
            }
        }

        private static Bitmap CreateAnimationObjectLayer(
            Bitmap source,
            Bitmap background,
            Rectangle rect,
            Color? targetColor,
            Color explicitFrom,
            float progress)
        {
            if (source == null ||
                background == null ||
                rect.Width <= 0 ||
                rect.Height <= 0)
            {
                return null;
            }

            Rectangle bounds =
                Rectangle.Intersect(
                    new Rectangle(
                        0,
                        0,
                        source.Width,
                        source.Height),
                    rect);

            if (bounds.Width <= 0 ||
                bounds.Height <= 0)
            {
                return null;
            }

            Bitmap layer =
                new Bitmap(
                    bounds.Width,
                    bounds.Height,
                    PixelFormat.Format32bppArgb);

            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            for (int y = 0;
                 y < bounds.Height;
                 y++)
            {
                for (int x = 0;
                     x < bounds.Width;
                     x++)
                {
                    Color sourceColor =
                        source.GetPixel(
                            bounds.Left + x,
                            bounds.Top + y);

                    Color backgroundColor =
                        background.GetPixel(
                            bounds.Left + x,
                            bounds.Top + y);

                    int difference =
                        Math.Max(
                            Math.Abs(
                                sourceColor.R -
                                backgroundColor.R),
                            Math.Max(
                                Math.Abs(
                                    sourceColor.G -
                                    backgroundColor.G),
                                Math.Abs(
                                    sourceColor.B -
                                    backgroundColor.B)));

                    if (difference < 3)
                    {
                        layer.SetPixel(
                            x,
                            y,
                            Color.Transparent);
                        continue;
                    }

                    int alpha =
                        Math.Max(
                            24,
                            Math.Min(
                                255,
                                difference *
                                10));

                    Color resultColor =
                        sourceColor;

                    if (targetColor.HasValue)
                    {
                        Color from =
                            explicitFrom.IsEmpty
                                ? sourceColor
                                : explicitFrom;

                        Color to =
                            targetColor.Value;

                        resultColor =
                            Color.FromArgb(
                                sourceColor.A,
                                (int)Math.Round(
                                    from.R +
                                    (to.R -
                                     from.R) *
                                    progress),
                                (int)Math.Round(
                                    from.G +
                                    (to.G -
                                     from.G) *
                                    progress),
                                (int)Math.Round(
                                    from.B +
                                    (to.B -
                                     from.B) *
                                    progress));
                    }

                    layer.SetPixel(
                        x,
                        y,
                        Color.FromArgb(
                            Math.Min(
                                resultColor.A,
                                alpha),
                            resultColor.R,
                            resultColor.G,
                            resultColor.B));
                }
            }

            return layer;
        }

        internal static bool TryEvaluateMotionPath(
            string pathData,
            float progress,
            out float x,
            out float y)
        {
            x = 0f;
            y = 0f;

            if (string.IsNullOrEmpty(
                    pathData))
            {
                return false;
            }

            progress =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        progress));

            using (GraphicsPath path =
                BuildSvgPath(
                    pathData,
                    new RectangleF(
                        0f,
                        0f,
                        1f,
                        1f),
                    0f,
                    0f,
                    1f,
                    1f))
            {
                if (path == null ||
                    path.PointCount < 2)
                {
                    return false;
                }

                path.Flatten();

                PointF[] points =
                    path.PathPoints;

                byte[] types =
                    path.PathTypes;

                if (points == null ||
                    points.Length < 2)
                {
                    return false;
                }

                double total = 0.0;

                for (int i = 1;
                     i < points.Length;
                     i++)
                {
                    if ((types[i] &
                         (byte)PathPointType.PathTypeMask) ==
                        (byte)PathPointType.Start)
                    {
                        continue;
                    }

                    double dx =
                        points[i].X -
                        points[i - 1].X;
                    double dy =
                        points[i].Y -
                        points[i - 1].Y;

                    total +=
                        Math.Sqrt(
                            dx * dx +
                            dy * dy);
                }

                if (total <= 0.0000001)
                    return false;

                double target =
                    total *
                    progress;

                double traversed = 0.0;

                PointF first =
                    points[0];

                if (progress <= 0f)
                {
                    x = first.X;
                    y = first.Y;
                    return true;
                }

                for (int i = 1;
                     i < points.Length;
                     i++)
                {
                    if ((types[i] &
                         (byte)PathPointType.PathTypeMask) ==
                        (byte)PathPointType.Start)
                    {
                        continue;
                    }

                    PointF from =
                        points[i - 1];
                    PointF to =
                        points[i];

                    double segX =
                        to.X -
                        from.X;
                    double segY =
                        to.Y -
                        from.Y;

                    double length =
                        Math.Sqrt(
                            segX * segX +
                            segY * segY);

                    if (length <= 0.0000001)
                        continue;

                    if (traversed +
                        length >=
                        target)
                    {
                        double local =
                            (target -
                             traversed) /
                            length;

                        x =
                            (float)(
                                from.X +
                                segX *
                                local);
                        y =
                            (float)(
                                from.Y +
                                segY *
                                local);
                        return true;
                    }

                    traversed += length;
                }

                PointF last =
                    points[
                        points.Length -
                        1];

                x = last.X;
                y = last.Y;
                return true;
            }
        }

        private static void TryReadMotionPathDelta(
            string path,
            out float dx,
            out float dy)
        {
            dx = 0f;
            dy = 0f;

            if (string.IsNullOrEmpty(path))
                return;

            List<double> values = new List<double>();
            string token = string.Empty;

            for (int i = 0; i <= path.Length; i++)
            {
                char ch = i < path.Length
                    ? path[i]
                    : ' ';

                bool numeric =
                    (ch >= '0' && ch <= '9') ||
                    ch == '-' ||
                    ch == '+' ||
                    ch == '.' ||
                    ch == 'e' ||
                    ch == 'E';

                if (numeric)
                {
                    token += ch;
                }
                else if (token.Length > 0)
                {
                    double value;
                    if (double.TryParse(
                            token,
                            NumberStyles.Float,
                            CultureInfo.InvariantCulture,
                            out value))
                    {
                        values.Add(value);
                    }
                    token = string.Empty;
                }
            }

            if (values.Count < 4)
                return;

            dx = (float)(
                values[values.Count - 2] -
                values[0]);
            dy = (float)(
                values[values.Count - 1] -
                values[1]);
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
