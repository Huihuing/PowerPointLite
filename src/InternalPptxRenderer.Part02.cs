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
        private static InteractiveRegion BuildInteractiveRegion(
            ZipArchive zip,
            XmlNode node,
            Dictionary<string, RelationshipInfo> rels,
            RectangleF bounds,
            PresentationInfo info,
            string cacheDir,
            int currentSlideIndex)
        {
            foreach (XmlNode hlink in FindAll(node, "hlinkClick"))
            {
                string rid = GetRelationshipId(hlink);
                string action = GetAttr(hlink, "action");

                if (!string.IsNullOrEmpty(rid) && rels.ContainsKey(rid))
                {
                    RelationshipInfo rel = rels[rid];
                    InteractiveRegion region = new InteractiveRegion();
                    region.Bounds = bounds;
                    region.Tooltip = GetAttr(hlink, "tooltip");

                    if (rel.Type != null && rel.Type.EndsWith("/slide", StringComparison.OrdinalIgnoreCase))
                    {
                        region.Kind = "slide";
                        region.SlideIndex = info.SlideParts.IndexOf(rel.ResolvedTarget);
                        region.Target = rel.ResolvedTarget;
                    }
                    else
                    {
                        region.Kind = "link";
                        region.Target = rel.TargetMode == "External" ? rel.Target : rel.ResolvedTarget;
                    }

                    return region;
                }

                if (!string.IsNullOrEmpty(action) && action.StartsWith("ppaction://", StringComparison.OrdinalIgnoreCase))
                {
                    string lower = action.ToLowerInvariant();
                    InteractiveRegion nav = new InteractiveRegion();
                    nav.Bounds = bounds;
                    nav.Kind = "slide";
                    nav.Tooltip = "PowerPoint navigation action";

                    if (lower.IndexOf("jump=nextslide") >= 0)
                        nav.SlideIndex = Math.Min(info.SlideParts.Count - 1, currentSlideIndex + 1);
                    else if (lower.IndexOf("jump=previousslide") >= 0)
                        nav.SlideIndex = Math.Max(0, currentSlideIndex - 1);
                    else if (lower.IndexOf("jump=firstslide") >= 0)
                        nav.SlideIndex = 0;
                    else if (lower.IndexOf("jump=lastslide") >= 0)
                        nav.SlideIndex = Math.Max(0, info.SlideParts.Count - 1);
                    else if (lower.IndexOf("endshow") >= 0)
                    {
                        nav.Kind = "endshow";
                        nav.SlideIndex = -1;
                    }
                    else
                        nav = null;

                    if (nav != null)
                        return nav;
                }
            }

            foreach (string mediaName in new string[] { "videoFile", "audioFile", "media" })
            {
                foreach (XmlNode media in FindAll(node, mediaName))
                {
                    string rid = GetRelationshipId(media);
                    if (string.IsNullOrEmpty(rid) || !rels.ContainsKey(rid))
                        continue;

                    RelationshipInfo rel = rels[rid];
                    string part = rel.ResolvedTarget;
                    ZipArchiveEntry entry = zip.GetEntry(part);

                    if (entry == null)
                        continue;

                    string mediaDir = Path.Combine(cacheDir, "media");
                    Directory.CreateDirectory(mediaDir);

                    string filename = Path.GetFileName(part);
                    string output = Path.Combine(mediaDir, filename);

                    if (!File.Exists(output))
                    {
                        using (Stream input = entry.Open())
                        using (FileStream outputStream = File.Create(output))
                            input.CopyTo(outputStream);
                    }

                    return new InteractiveRegion
                    {
                        Bounds = bounds,
                        Kind = "media",
                        Target = output,
                        Tooltip = "Play media: " + filename
                    };
                }
            }

            return null;
        }

        private static PresentationInfo ReadPresentationInfo(ZipArchive zip)
        {
            PresentationInfo info = new PresentationInfo
            {
                WidthEmu = 12192000,
                HeightEmu = 6858000
            };

            LoadThemeColors(zip, info.ThemeColors);
            LoadThemeFonts(zip);

            XmlDocument presentation = LoadXml(zip, "ppt/presentation.xml");
            if (presentation == null)
                throw new InvalidOperationException("ppt/presentation.xml is missing.");

            XmlNode sldSz = FindFirst(presentation, "sldSz");
            if (sldSz != null)
            {
                info.WidthEmu = GetLong(sldSz, "cx", info.WidthEmu);
                info.HeightEmu = GetLong(sldSz, "cy", info.HeightEmu);
            }

            Dictionary<string, string> rels =
                LoadRelationships(zip, "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml");

            foreach (XmlNode node in FindAll(presentation, "sldId"))
            {
                string rid = GetRelationshipId(node);

                if (!string.IsNullOrEmpty(rid) && rels.ContainsKey(rid))
                    info.SlideParts.Add(rels[rid]);
            }

            if (info.SlideParts.Count == 0)
            {
                List<string> fallback = new List<string>();

                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    string name = NormalizePart(entry.FullName);

                    if (name.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                        name.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        fallback.Add(name);
                    }
                }

                fallback.Sort(StringComparer.OrdinalIgnoreCase);
                info.SlideParts.AddRange(fallback);
            }

            return info;
        }

        private static Bitmap RenderSlide(
            ZipArchive zip,
            PresentationInfo info,
            string slidePart,
            int width,
            int height)
        {
            return RenderSlide(zip, info, slidePart, width, height, null);
        }

        private static Bitmap RenderSlide(
            ZipArchive zip,
            PresentationInfo info,
            string slidePart,
            int width,
            int height,
            HashSet<string> hiddenShapeIds)
        {
            Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            try
            {
                float physicalWidthInches = (float)(info.WidthEmu / EmuPerInch);
                float dpi = physicalWidthInches > 0.1f ? width / physicalWidthInches : 144f;
                dpi = Math.Max(72f, Math.Min(300f, dpi));
                bmp.SetResolution(dpi, dpi);
            }
            catch { }

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                TransformContext root = new TransformContext
                {
                    Ax = (double)width / info.WidthEmu,
                    Bx = 0,
                    Ay = (double)height / info.HeightEmu,
                    By = 0
                };

                List<string> layers = GetSlideVisualLayers(zip, slidePart);
                PrepareRichTextInheritance(zip, layers);
                g.Clear(Color.White);

                for (int i = 0; i < layers.Count; i++)
                {
                    string bgPart = layers[i];
                    XmlDocument bgDoc = LoadXml(zip, bgPart);
                    Dictionary<string, string> bgRels = LoadRelationships(zip, RelationshipPart(bgPart), bgPart);

                    DrawBackgroundLayer(
                        zip,
                        g,
                        bgDoc,
                        bgRels,
                        new RectangleF(0, 0, width, height),
                        info.ThemeColors);
                }

                Dictionary<string, RectangleF> placeholderRects = BuildPlaceholderRectangles(zip, layers, root);
                int slideNumber = Math.Max(1, info.SlideParts.IndexOf(slidePart) + 1);

                for (int i = 0; i < layers.Count; i++)
                {
                    string part = layers[i];
                    XmlDocument doc = LoadXml(zip, part);
                    if (doc == null) continue;

                    Dictionary<string, string> rels = LoadRelationships(zip, RelationshipPart(part), part);
                    XmlNode spTree = FindFirst(doc, "spTree");
                    if (spTree == null) continue;

                    bool inheritedLayer =
                        part.IndexOf("/slideMasters/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        part.IndexOf("/slideLayouts/", StringComparison.OrdinalIgnoreCase) >= 0;

                    DrawContainer(
                        zip,
                        g,
                        spTree,
                        rels,
                        root,
                        info.ThemeColors,
                        inheritedLayer,
                        placeholderRects,
                        hiddenShapeIds,
                        slideNumber);
                }
            }

            return bmp;
        }

        private static List<string> GetSlideVisualLayers(ZipArchive zip, string slidePart)
        {
            List<string> layers = new List<string>();
            Dictionary<string, string> slideRels = LoadRelationships(zip, RelationshipPart(slidePart), slidePart);

            string layout = FindRelationshipTargetByType(zip, RelationshipPart(slidePart), slidePart, "slideLayout");
            string master = null;

            if (!string.IsNullOrEmpty(layout))
                master = FindRelationshipTargetByType(zip, RelationshipPart(layout), layout, "slideMaster");

            if (!string.IsNullOrEmpty(master)) layers.Add(master);
            if (!string.IsNullOrEmpty(layout)) layers.Add(layout);
            layers.Add(slidePart);

            return layers;
        }

        private static string FindRelationshipTargetByType(
            ZipArchive zip,
            string relPart,
            string sourcePart,
            string typeSuffix)
        {
            XmlDocument relDoc = LoadXml(zip, relPart);
            if (relDoc == null) return null;

            foreach (XmlNode rel in FindAll(relDoc, "Relationship"))
            {
                string type = GetAttr(rel, "Type");
                string target = GetAttr(rel, "Target");

                if (!string.IsNullOrEmpty(type) && type.EndsWith("/" + typeSuffix, StringComparison.OrdinalIgnoreCase))
                    return ResolvePart(sourcePart, target);
            }

            return null;
        }

        private static void DrawContainer(
            ZipArchive zip,
            Graphics g,
            XmlNode container,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme,
            bool inheritedLayer,
            Dictionary<string, RectangleF> placeholderRects,
            HashSet<string> hiddenShapeIds,
            int slideNumber)
        {
            foreach (XmlNode child in container.ChildNodes)
            {
                string name = child.LocalName;

                if (!inheritedLayer && hiddenShapeIds != null)
                {
                    string shapeId = GetShapeId(child);
                    if (!string.IsNullOrEmpty(shapeId) && hiddenShapeIds.Contains(shapeId))
                        continue;
                }

                if (name == "sp")
                {
                    if (inheritedLayer && ShouldSkipInheritedPlaceholder(child)) continue;
                    DrawShape(zip, g, child, rels, ctx, theme, placeholderRects, slideNumber);
                }
                else if (name == "pic") DrawPicture(zip, g, child, rels, ctx);
                else if (name == "cxnSp") DrawConnector(g, child, ctx, theme);
                else if (name == "graphicFrame") DrawGraphicFrame(zip, g, child, rels, ctx, theme);
                else if (name == "grpSp")
                {
                    TransformContext childCtx = BuildGroupContext(child, ctx);
                    DrawContainer(zip, g, child, rels, childCtx, theme, inheritedLayer, placeholderRects, hiddenShapeIds, slideNumber);
                }
            }
        }

        private static string GetShapeId(XmlNode node)
        {
            if (node == null) return null;
            XmlNode cNvPr = FindFirst(node, "cNvPr");
            return cNvPr != null ? GetAttr(cNvPr, "id") : null;
        }

        private static Dictionary<string, RectangleF> BuildPlaceholderRectangles(
            ZipArchive zip,
            List<string> layers,
            TransformContext root)
        {
            Dictionary<string, RectangleF> result = new Dictionary<string, RectangleF>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < layers.Count - 1; i++)
            {
                XmlDocument doc = LoadXml(zip, layers[i]);
                if (doc == null) continue;

                XmlNode spTree = FindFirst(doc, "spTree");
                if (spTree == null) continue;

                foreach (XmlNode sp in FindAll(spTree, "sp"))
                {
                    XmlNode ph = FindFirst(sp, "ph");
                    if (ph == null) continue;

                    RectangleF rect;
                    if (!TryGetRect(sp, root, out rect)) continue;

                    string idx = GetAttr(ph, "idx");
                    string type = GetAttr(ph, "type");

                    if (!string.IsNullOrEmpty(idx)) result["idx:" + idx] = rect;
                    if (!string.IsNullOrEmpty(type)) result["type:" + type] = rect;
                    if (string.IsNullOrEmpty(idx) && string.IsNullOrEmpty(type)) result["type:body"] = rect;
                }
            }

            return result;
        }

        private static bool TryGetPlaceholderRect(
            XmlNode shape,
            Dictionary<string, RectangleF> placeholderRects,
            out RectangleF rect)
        {
            rect = RectangleF.Empty;
            if (placeholderRects == null) return false;

            XmlNode ph = FindFirst(shape, "ph");
            if (ph == null) return false;

            string idx = GetAttr(ph, "idx");
            string type = GetAttr(ph, "type");

            if (!string.IsNullOrEmpty(idx) && placeholderRects.TryGetValue("idx:" + idx, out rect)) return true;
            if (!string.IsNullOrEmpty(type) && placeholderRects.TryGetValue("type:" + type, out rect)) return true;
            if (placeholderRects.TryGetValue("type:body", out rect)) return true;
            if (placeholderRects.TryGetValue("type:obj", out rect)) return true;
            return false;
        }

        private static bool IsPlaceholder(XmlNode shape)
        {
            return FindFirst(shape, "ph") != null;
        }

        private static bool ShouldSkipInheritedPlaceholder(XmlNode shape)
        {
            XmlNode ph = FindFirst(shape, "ph");
            if (ph == null) return false;

            string type = GetAttr(ph, "type");
            if (type == "dt" || type == "ftr" || type == "sldNum") return false;
            return true;
        }

        private static TransformContext BuildGroupContext(XmlNode group, TransformContext parent)
        {
            XmlNode grpSpPr = DirectChild(group, "grpSpPr");
            XmlNode xfrm = grpSpPr != null ? DirectChild(grpSpPr, "xfrm") : null;
            if (xfrm == null) return parent;

            XmlNode off = DirectChild(xfrm, "off");
            XmlNode ext = DirectChild(xfrm, "ext");
            XmlNode chOff = DirectChild(xfrm, "chOff");
            XmlNode chExt = DirectChild(xfrm, "chExt");

            long ox = GetLong(off, "x", 0);
            long oy = GetLong(off, "y", 0);
            long ew = GetLong(ext, "cx", 1);
            long eh = GetLong(ext, "cy", 1);
            long cox = GetLong(chOff, "x", 0);
            long coy = GetLong(chOff, "y", 0);
            long cew = Math.Max(1, GetLong(chExt, "cx", ew));
            long ceh = Math.Max(1, GetLong(chExt, "cy", eh));

            double sx = (double)ew / cew;
            double sy = (double)eh / ceh;

            return new TransformContext
            {
                Ax = parent.Ax * sx,
                Bx = parent.Ax * (ox - cox * sx) + parent.Bx,
                Ay = parent.Ay * sy,
                By = parent.Ay * (oy - coy * sy) + parent.By
            };
        }
    }
}
