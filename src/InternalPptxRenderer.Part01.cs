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
internal static partial class InternalPptxRenderer
    {
        private const double EmuPerInch = 914400.0;
        private static readonly Dictionary<string, string> ThemeFonts =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public sealed class PresentationInfo
        {
            public long WidthEmu;
            public long HeightEmu;
            public readonly List<string> SlideParts = new List<string>();
            public readonly Dictionary<string, Color> ThemeColors =
                new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
        }

        public sealed class InteractiveRegion
        {
            public RectangleF Bounds;
            public string Kind;
            public string Target;
            public string Tooltip;
            public int SlideIndex = -1;
        }

        public sealed class TransitionSpec
        {
            public string Kind = "cut";
            public string Direction = "";
            public int DurationMs;
        }

        public sealed class ViewerMetadata
        {
            public readonly List<string> SpeakerNotes = new List<string>();
            public readonly List<string> SlideTexts = new List<string>();
            public readonly List<string> SlideTitles = new List<string>();
            public readonly List<bool> HiddenSlides = new List<bool>();
        }

        private sealed class RelationshipInfo
        {
            public string Id;
            public string Type;
            public string Target;
            public string TargetMode;
            public string ResolvedTarget;
        }

        private sealed class TransformContext
        {
            public double Ax;
            public double Bx;
            public double Ay;
            public double By;

            public float X(long emu) { return (float)(Ax * emu + Bx); }
            public float Y(long emu) { return (float)(Ay * emu + By); }
            public float W(long emu) { return (float)(Ax * emu); }
            public float H(long emu) { return (float)(Ay * emu); }
        }

        public static List<string> Render(string file, string cacheDir)
        {
            string ext = Path.GetExtension(file).ToLowerInvariant();

            if (ext != ".pptx" && ext != ".pptm")
                throw new InvalidOperationException(
                    "Internal OpenXML supports PPTX/PPTM. Legacy binary PPT needs PowerPoint or LibreOffice.");

            List<string> result = new List<string>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                if (info.SlideParts.Count == 0)
                    throw new InvalidOperationException("The PPTX contains no slides.");

                int width = 1920;
                int height = (int)Math.Round(width * ((double)info.HeightEmu / info.WidthEmu));

                if (height < 200 || height > 4000)
                    height = 1080;

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    string output = Path.Combine(
                        cacheDir,
                        "internal_" + (i + 1).ToString("D4") + ".png");

                    if (!File.Exists(output))
                    {
                        using (Bitmap bmp = RenderSlide(zip, info, info.SlideParts[i], width, height))
                        {
                            bmp.Save(output, ImageFormat.Png);
                        }
                    }

                    result.Add(output);
                }
            }

            return result;
        }

        public static PresentationInfo ReadPresentationInfo(string file)
        {
            using (ZipArchive zip = ZipFile.OpenRead(file))
                return ReadPresentationInfo(zip);
        }

        public static ViewerMetadata ReadViewerMetadata(string file)
        {
            ViewerMetadata result = new ViewerMetadata();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    string slidePart = info.SlideParts[i];
                    XmlDocument slideDoc = LoadXml(zip, slidePart);

                    bool hidden = false;
                    string title = "";
                    string allText = "";

                    if (slideDoc != null)
                    {
                        XmlNode root = slideDoc.DocumentElement;
                        string show = GetAttr(root, "show");

                        hidden =
                            show == "0" ||
                            string.Equals(show, "false", StringComparison.OrdinalIgnoreCase);

                        allText = ExtractText(slideDoc);
                        title = ExtractSlideTitle(slideDoc);

                        if (string.IsNullOrWhiteSpace(title))
                        {
                            string firstLine = allText ?? "";
                            int newline = firstLine.IndexOf('\n');

                            if (newline >= 0)
                                firstLine = firstLine.Substring(0, newline);

                            title = firstLine.Trim();

                            if (title.Length > 80)
                                title = title.Substring(0, 80);
                        }
                    }

                    result.HiddenSlides.Add(hidden);
                    result.SlideTexts.Add(allText ?? "");
                    result.SlideTitles.Add(
                        string.IsNullOrWhiteSpace(title)
                        ? "Slide " + (i + 1).ToString()
                        : title);

                    string notesPart = FindRelationshipTargetByType(
                        zip,
                        RelationshipPart(slidePart),
                        slidePart,
                        "notesSlide");

                    XmlDocument notesDoc = string.IsNullOrEmpty(notesPart)
                        ? null
                        : LoadXml(zip, notesPart);

                    result.SpeakerNotes.Add(
                        notesDoc == null
                        ? ""
                        : ExtractNotesText(notesDoc));
                }
            }

            return result;
        }

        private static string ExtractSlideTitle(XmlDocument slideDoc)
        {
            XmlNode spTree = FindFirst(slideDoc, "spTree");
            if (spTree == null)
                return "";

            foreach (XmlNode sp in FindAll(spTree, "sp"))
            {
                XmlNode ph = FindFirst(sp, "ph");
                string type = GetAttr(ph, "type");

                if (type == "title" || type == "ctrTitle")
                {
                    string text = ExtractText(sp);
                    if (!string.IsNullOrWhiteSpace(text))
                        return text.Trim().Replace("\r", " ").Replace("\n", " ");
                }
            }

            return "";
        }

        private static string ExtractNotesText(XmlDocument notesDoc)
        {
            XmlNode spTree = FindFirst(notesDoc, "spTree");
            if (spTree == null)
                return ExtractText(notesDoc);

            StringBuilder sb = new StringBuilder();

            foreach (XmlNode sp in FindAll(spTree, "sp"))
            {
                XmlNode ph = FindFirst(sp, "ph");
                string type = GetAttr(ph, "type");

                if (type == "sldImg" || type == "sldNum" || type == "dt" || type == "ftr" || type == "hdr")
                    continue;

                string text = ExtractText(sp);

                if (!string.IsNullOrWhiteSpace(text))
                {
                    if (sb.Length > 0)
                        sb.AppendLine();
                    sb.Append(text.Trim());
                }
            }

            return sb.ToString();
        }

        private static string ExtractText(XmlNode node)
        {
            if (node == null)
                return "";

            StringBuilder sb = new StringBuilder();

            foreach (XmlNode p in FindAll(node, "p"))
            {
                StringBuilder line = new StringBuilder();
                foreach (XmlNode t in FindAll(p, "t"))
                    line.Append(t.InnerText);

                if (line.Length > 0)
                {
                    if (sb.Length > 0) sb.AppendLine();
                    sb.Append(line.ToString());
                }
            }

            if (sb.Length == 0)
            {
                foreach (XmlNode t in FindAll(node, "t"))
                {
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(t.InnerText);
                }
            }

            return sb.ToString();
        }

        public static List<List<InteractiveRegion>> ReadInteractiveRegions(string file, string cacheDir)
        {
            List<List<InteractiveRegion>> result = new List<List<InteractiveRegion>>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    string slidePart = info.SlideParts[i];
                    XmlDocument doc = LoadXml(zip, slidePart);
                    List<InteractiveRegion> regions = new List<InteractiveRegion>();
                    result.Add(regions);

                    if (doc == null)
                        continue;

                    Dictionary<string, RelationshipInfo> rels =
                        LoadRelationshipInfos(zip, RelationshipPart(slidePart), slidePart);

                    TransformContext root = new TransformContext
                    {
                        Ax = 1.0 / info.WidthEmu,
                        Bx = 0,
                        Ay = 1.0 / info.HeightEmu,
                        By = 0
                    };

                    XmlNode spTree = FindFirst(doc, "spTree");
                    if (spTree == null)
                        continue;

                    List<string> layers = GetSlideVisualLayers(zip, slidePart);
                    Dictionary<string, RectangleF> placeholderRects = BuildPlaceholderRectangles(zip, layers, root);

                    CollectInteractiveRegions(zip, spTree, rels, root, info, cacheDir, regions, placeholderRects, i);
                }
            }

            return result;
        }

        public static List<TransitionSpec> ReadTransitions(string file)
        {
            List<TransitionSpec> result = new List<TransitionSpec>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    TransitionSpec spec = new TransitionSpec();
                    XmlDocument doc = LoadXml(zip, info.SlideParts[i]);
                    XmlNode transition = FindFirst(doc, "transition");

                    if (transition == null)
                    {
                        result.Add(spec);
                        continue;
                    }

                    string speed = GetAttr(transition, "spd");
                    spec.DurationMs = speed == "slow" ? 650 : speed == "med" ? 400 : 240;

                    int explicitDuration;
                    string durationText = GetAttr(transition, "dur");
                    if (!string.IsNullOrEmpty(durationText) &&
                        int.TryParse(durationText, out explicitDuration) &&
                        explicitDuration > 0 &&
                        explicitDuration <= 60000)
                    {
                        spec.DurationMs = explicitDuration;
                    }

                    spec.Kind = "fade";

                    foreach (XmlNode child in transition.ChildNodes)
                    {
                        string name = child.LocalName;
                        if (name == "cut") spec.Kind = "cut";
                        else if (name == "wipe" || name == "cover" || name == "uncover") spec.Kind = "wipe";
                        else if (name == "push" || name == "pull") spec.Kind = "push";
                        else if (name == "fade" || name == "dissolve") spec.Kind = "fade";
                        else if (name == "morph") spec.Kind = "fade";
                        else if (name != "sndAc") spec.Kind = "fade";

                        string direction = GetAttr(child, "dir");
                        if (direction == "l" ||
                            direction == "r" ||
                            direction == "u" ||
                            direction == "d")
                        {
                            spec.Direction = direction;
                        }
                    }

                    result.Add(spec);
                }
            }

            return result;
        }

        public static List<List<string>> ReadAnimationSteps(string file)
        {
            List<List<string>> result = new List<List<string>>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);

                for (int i = 0; i < info.SlideParts.Count; i++)
                {
                    List<string> steps = new List<string>();
                    result.Add(steps);

                    XmlDocument doc = LoadXml(zip, info.SlideParts[i]);
                    XmlNode timing = FindFirst(doc, "timing");
                    if (timing == null) continue;

                    foreach (XmlNode target in FindAll(timing, "spTgt"))
                    {
                        string spid = GetAttr(target, "spid");
                        if (!string.IsNullOrEmpty(spid) && IsEntranceAnimationTarget(target) && !steps.Contains(spid))
                            steps.Add(spid);
                    }
                }
            }

            return result;
        }

        public static string RenderAnimationState(string file, string cacheDir, int slideIndex, int revealCount)
        {
            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info = ReadPresentationInfo(zip);
                if (slideIndex < 0 || slideIndex >= info.SlideParts.Count)
                    return null;

                List<List<string>> allSteps = ReadAnimationStepsFromZip(zip, info);
                List<string> steps = slideIndex < allSteps.Count ? allSteps[slideIndex] : new List<string>();

                revealCount = Math.Max(0, Math.Min(steps.Count, revealCount));
                HashSet<string> hidden = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = revealCount; i < steps.Count; i++) hidden.Add(steps[i]);

                int width = 1920;
                int height = (int)Math.Round(width * ((double)info.HeightEmu / info.WidthEmu));
                if (height < 200 || height > 4000) height = 1080;

                string output = Path.Combine(cacheDir,
                    "anim_" + (slideIndex + 1).ToString("D4") + "_" + revealCount.ToString("D3") + ".png");

                if (!File.Exists(output))
                {
                    using (Bitmap bmp = RenderSlide(zip, info, info.SlideParts[slideIndex], width, height, hidden))
                        bmp.Save(output, ImageFormat.Png);
                }

                return output;
            }
        }

        private static List<List<string>> ReadAnimationStepsFromZip(ZipArchive zip, PresentationInfo info)
        {
            List<List<string>> result = new List<List<string>>();

            for (int i = 0; i < info.SlideParts.Count; i++)
            {
                List<string> steps = new List<string>();
                result.Add(steps);

                XmlDocument doc = LoadXml(zip, info.SlideParts[i]);
                XmlNode timing = FindFirst(doc, "timing");
                if (timing == null) continue;

                foreach (XmlNode target in FindAll(timing, "spTgt"))
                {
                    string spid = GetAttr(target, "spid");
                    if (!string.IsNullOrEmpty(spid) && IsEntranceAnimationTarget(target) && !steps.Contains(spid))
                        steps.Add(spid);
                }
            }

            return result;
        }

        private static bool IsEntranceAnimationTarget(XmlNode target)
        {
            XmlNode node = target;
            bool sawSet = false;

            while (node != null)
            {
                if (node.LocalName == "animEffect")
                {
                    string transition = GetAttr(node, "transition");
                    if (transition == "out") return false;
                    if (transition == "in") return true;
                }

                if (node.LocalName == "set") sawSet = true;
                if (node.LocalName == "timing") break;
                node = node.ParentNode;
            }

            return sawSet;
        }

        private static void CollectInteractiveRegions(
            ZipArchive zip, XmlNode container, Dictionary<string, RelationshipInfo> rels,
            TransformContext ctx, PresentationInfo info, string cacheDir,
            List<InteractiveRegion> regions, Dictionary<string, RectangleF> placeholderRects,
            int currentSlideIndex)
        {
            foreach (XmlNode child in container.ChildNodes)
            {
                string name = child.LocalName;

                if (name == "grpSp")
                {
                    TransformContext group = BuildGroupContext(child, ctx);
                    CollectInteractiveRegions(zip, child, rels, group, info, cacheDir, regions, placeholderRects, currentSlideIndex);
                    continue;
                }

                if (name != "sp" && name != "pic" && name != "graphicFrame" && name != "cxnSp")
                    continue;

                RectangleF rect;
                if (!TryGetRect(child, ctx, out rect) && !TryGetPlaceholderRect(child, placeholderRects, out rect))
                    continue;

                InteractiveRegion region = BuildInteractiveRegion(zip, child, rels, rect, info, cacheDir, currentSlideIndex);
                if (region != null) regions.Add(region);
            }
        }
    }
}
